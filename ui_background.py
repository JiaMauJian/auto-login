"""
背景工作：瀏覽器執行緒的生命週期、登入／讀取／寫入、Excel 開啟狀態輪詢、
現金算法的問答。全部丟到背景執行緒的東西都在這裡協調。
"""

import datetime
import os
import queue
import threading
import traceback
from pathlib import Path
from tkinter import filedialog

from playwright.sync_api import Error as PlaywrightError
from playwright.sync_api import sync_playwright

import excel_io
import ledger as ledger_mod
import order_fill
import planner
import fetch as fetch_mod
from fetch import collect, login_only
from login import app_dir, configure_browsers_path, open_context
from ui_common import ask_cash_method, ask_confirm, show_error, show_warning

# 背景做的三件事，講給人聽的名字。收尾出錯時要說得出是哪一步壞掉的。
STEP_NAMES = {"logged_in": "登入", "fetched": "讀取", "written": "寫入", "logged_out": "登出",
              "order_stock_list": "讀取ＯＯ持股", "order_plans": "讀取試算",
              "order_filled": "下單填單",
              "order_dialog_closed": "委託確認視窗關閉偵測",
              "order_price_refresh": "多輪出清重讀持股",
              "order_odd_cancelled": "出清零股自動撤單",
              "order_rates": "帳戶報酬率補讀", "excel_layout": "Excel 版面錨點檢查",
              "order_quotes_fetched": "查詢委買賣", "pending_fetched": "查詢掛單",
              "pending_cancelled": "取消掛單"}

# 瀏覽器起不來時，錯誤視窗最上面那段人話。traceback 講的是 Playwright 的內部狀況，
# 對使用者沒有意義，真正能動手的只有下面這兩件事。
BROWSER_HINT = (
    "瀏覽器沒能開起來，所以這次的動作沒有做。常見的原因有兩個：\n"
    "・這台電腦上找不到指定的瀏覽器（.env 的 BROWSER_CHANNEL 指到沒裝的版本），"
    "或 Playwright 的 Chromium 還沒下載完\n"
    "・USER_DATA_DIR 那個資料夾正被另一個 Chrome 視窗開著 —— 同一個資料夾不能同時"
    "給兩個 Chrome 用，請把 Chrome 全部關掉再按一次"
)


def _browser_alive(context):
    """瀏覽器是不是還開著（使用者有沒有自己把它關掉）。"""
    try:
        return any(not pg.is_closed() for pg in context.pages)
    except PlaywrightError:
        return False


def _pump_browser(context):
    """
    背景執行緒閒置時呼叫：隨便找一個活著的分頁進 Playwright 一下，逼它去處理
    瀏覽器送來的事件。見 login.wait_until_finished 的說明。

    還沒登入、context 是 None 時什麼也不做；context 已經被使用者關掉、
    或這次呼叫剛好碰上分頁正在關閉，都當作沒事發生，下一輪再試。
    """
    if context is None:
        return
    try:
        pages = [pg for pg in context.pages if not pg.is_closed()]
        if pages:
            pages[0].wait_for_timeout(100)
    except PlaywrightError:
        pass


def _read_excel_after_fetch(records, path):
    """背景執行緒用：登入抓完網頁資料後，順便把 Excel 現值讀出來。只回傳純資料，不回傳任何 COM 物件。"""
    import pythoncom

    excel = workbook = None
    pythoncom.CoInitialize()
    try:
        with excel_io.opened(path, False) as (excel, workbook, attached):
            sheets, sheet_errors = {}, {}
            for record in records:
                name = record.get("sheet_name")
                if not name or record["problems"]:
                    continue
                sheet, error = excel_io.find_sheet(workbook, name)
                if sheet is None:
                    sheet_errors[name] = error
                else:
                    sheets[name] = excel_io.read_sheet(sheet)
    finally:
        # COM 參考要在 CoUninitialize 之前放掉。反過來的話，物件會在
        # 「已經沒有 COM 的執行緒」上被回收，那是未定義行為。
        sheet = excel = workbook = None
        pythoncom.CoUninitialize()

    return {"records": records, "sheets": sheets, "sheet_errors": sheet_errors, "attached": attached}


def _error_text(payload):
    """錯誤視窗的內容。看得懂的說明放最上面，原始 traceback 留在下面備查。"""
    detail = payload["error"][-1500:]
    hint = payload.get("hint")
    return f"{hint}\n\n────────────────\n{detail}" if hint else detail


class UiBackgroundMixin:
    # ---------- 背景工作 ----------

    def _selected_accounts(self):
        choice = self.account_choice.current()
        numbered = list(enumerate(self.accounts, start=1))
        return numbered if choice <= 0 else [numbered[choice - 1]]

    def _account_choices(self):
        """範圍選單上的每一列。名字知道了就寫名字 —— 20 組時光看「第 7 組」不知道是誰。"""
        choices = ["全部"]
        for i, account in enumerate(self.accounts, start=1):
            name = self.trader_of.get(i)
            label = f"第 {i} 組"
            if name:
                label += f"　{name}" + ("（模擬）" if account.get("fake") else "")
            choices.append(label)
        return choices

    def _refresh_account_choices(self):
        """
        登入或讀取之後名字才知道，選單上那幾列要跟著補上去。選中的那一列不動。

        掛單分頁的範圍選單也在這裡一起刷。它本來只在切到那個分頁的時候刷
        （_on_tab_changed），結果是「人一直待在掛單分頁按登入」的話，那個選單
        停在登入前的樣子——只剩「全部帳戶」、名字一個都沒有，而且不會自己好
        （2026/08/31 使用者遇到）。兩個選單問的是同一件事「現在知道哪幾位」，
        就該在同一個地方一起更新。

        下單分頁的「執行帳戶」**不在這裡**：2026/09/01 起那份清單來自 Excel 分頁
        （見 ui_order.refresh_order_accounts），跟登入拿到誰的名字無關。
        """
        keep = max(self.account_choice.current(), 0)
        self.account_choice.configure(values=self._account_choices())
        self.account_choice.current(keep)
        self._apply_scope_state()
        self._refresh_fetch_button()
        self._refresh_pending_scope()

    def _scope_ready(self):
        """範圍下拉能不能讓人手動切到某一組。

        名字不到齊之前不給選：「登入」跟「更新全部帳戶」選在「全部」時
        本來就會把每一組都登入一輪、順便知道名字（見 _on_logged_in），
        這裡等的正是那一輪做完 —— 不然使用者可能在誰都還沒登入過的時候，
        手動切到一個從沒動過的組別，按鈕只能落到「更新（第 N 組）帳戶」
        那種叫不出名字的備援文字。模擬帳號一開機名字就有了，天生就緒。
        """
        return self.excel_open and not self.busy and all(
            i in self.trader_of for i in range(1, len(self.accounts) + 1))

    def _apply_scope_state(self):
        self.account_choice.configure(state="readonly" if self._scope_ready() else "disabled")

    def _scope_order(self):
        """這次要做第幾組；None 代表全部。

        不管只有一組還是有很多組，選單停在「全部」就是全部 —— 第一次
        一定是全部讀取，之後要各別讀取得自己把選單切到那一組，按鈕文字
        才會跟著換成「更新（某某）帳戶」或「更新（第 N 組）帳戶」。
        """
        choice = self.account_choice.current()
        return choice if choice > 0 else None

    def _scope_name(self):
        """這次要做的是哪一位交易人；全部、或名字還不知道時是 None。"""
        order = self._scope_order()
        return None if order is None else self.trader_of.get(order)

    def _refresh_fetch_button(self):
        """
        按鈕上的字就是「按下去會動到誰」，一律用「更新」這個動詞。更新的是什麼
        （Excel 上的持股與現金餘額）兩個狀態都一樣，不寫在按鈕上，寫在旁邊那句
        Hint（見 ui_layout._build_toolbar）。

        2026/08/30 從「讀取」改成「更新」。「讀取」只講了這一按的前半段——它讀
        網頁**而且**會把 E/F/B8 覆蓋掉，一顆會改你試算表的按鈕不該叫「讀取」，
        那是三個候選字裡最不準的一個。動詞也就跟著分頁名（更新）對上了。

        名字還不知道（那一組沒登入過）時寫「第 3 組」而不是硬掰一個名字：
        這種時候按下去確實只做那一組，只是程式還說不出他是誰。正常流程會先
        按「登入」，那一步就會把名字補上，所以這條路很少真的走到。

        原本還有第三個狀態「登入+讀取全部帳戶」（還有帳號沒登入過時），用來標
        「這一按會順便補做登入、因此慢一個數量級」。2026/08/30 拿掉：「登入」
        搬到分頁列上面那條跨分頁常駐的列，從哪一頁都看得到、也不可能再被漏按
        （見 ui_layout._build_session_bar）；而真的漏按了，按下去的狀態列本來
        就會說「還沒登入的話瀏覽器會自己開起來」（見 start_fetch），資訊沒有
        因此消失，只是從按鈕挪到按下去的那一刻。少一個狀態，這顆按鈕就只剩
        「做誰」一個變數。
        """
        order = self._scope_order()
        if order is None:
            text = "更新全部帳戶"
        else:
            name = self.trader_of.get(order)
            text = f"更新（{name}）帳戶" if name else f"更新（第 {order} 組）帳戶"
        self.fetch_button.configure(text=text)

    def _on_scope_changed(self, _event=None):
        """上面換了範圍，左邊名單也跳到那一位 —— 兩邊是同一個選擇的兩個入口。"""
        self._refresh_fetch_button()
        name = self._scope_name()
        if name and name != self.current_sheet and name in self._shown():
            self.current_sheet = name
            self.people.selection_set(name)
            self.people.see(name)
            self._fill_right()

    def _sync_scope_to_person(self):
        """
        左邊名單換人，上面的範圍跟著換成他。

        對不到帳號時範圍不動（名單是網頁資料長出來的，正常情況一定對得到）
        —— 硬把範圍留在別人身上也不會說謊，按鈕上寫的一直是範圍裡的那個人。
        """
        order = next((i for i, name in self.trader_of.items() if name == self.current_sheet), None)
        if order is not None and order <= len(self.accounts):
            if self.account_choice.current() != order:
                self.account_choice.current(order)
        self._refresh_fetch_button()

    def _ensure_browser_thread(self):
        if self.browser_thread is None or not self.browser_thread.is_alive():
            self.browser_thread = threading.Thread(target=self._browser_worker, daemon=True)
            self.browser_thread.start()

    def start_login(self):
        """
        登入**一律全部**，不看更新分頁那個「範圍」。

        2026/08/30 使用者確認：不會單獨只登入一組。實際動線是開盤前按一次登入
        （順便把交易人姓名收齊，見 _on_logged_in），之後整天都在「讀取（某人）
        帳戶」——而那一顆本來就會順便把還沒登入的那組登進去（見 fetch.collect），
        所以「只登入某一組」這件事沒有任何入口消失。範圍留在更新分頁只管讀取
        （見 ui_layout._build_toolbar），這一顆搬到分頁列上面那條跨分頁常駐列
        （見 ui_layout._build_session_bar），兩邊不再共用同一個選擇。

        不叫 _require_excel()：登入根本不碰 Excel（見 ui_sync._sync_buttons 那顆
        按鈕的說明），帳號清單來自 .env，連 path 都只是順手帶著給「讀取」那條
        共用路解包用的，登入這一支不會拿它做任何事（見 _browser_worker）。
        """
        if self.busy:
            return
        if not self.accounts:
            show_error(self.root, "沒有帳號", "請先在 .env 填入 TBB_ID_1 / TBB_PASSWORD_1。")
            return
        if self.ledger_error:
            show_error(self.root, "紀錄檔有問題", self.ledger_error)
            return

        self._ensure_browser_thread()
        self._set_busy(True, "登入中，瀏覽器會自己開起來，請不要關掉它…")
        self.browser_waiting += 1
        self.browser_cmd_queue.put(("login", (list(enumerate(self.accounts, start=1)), self.path)))

    def start_fetch(self):
        # 看 _excel_in_use() 而不是只看 self.busy：下單分頁那幾條路也在用 COM 動
        # 同一份活頁簿（見那個述詞的說明）。
        if self._excel_in_use() or not self._require_excel():
            return
        if not self.accounts:
            show_error(self.root, "沒有帳號", "請先在 .env 填入 TBB_ID_1 / TBB_PASSWORD_1。")
            return
        if self.ledger_error:
            show_error(self.root, "紀錄檔有問題", self.ledger_error)
            return

        # 這一輪要做的是誰，按下去的當下就記起來：等結果回來的這幾十秒裡，
        # 使用者隨時可能在左邊名單上點別人（範圍會跟著換），報告卻是在講剛才那一輪。
        # 先問今天用哪一種算法（每天第一次），因為它決定這一輪要不要多查銀行餘額。
        # 使用者取消就整個收手——不登入、不讀取、不進 busy 狀態，維持按下去之前的樣子。
        self.today = datetime.date.today()
        if not self._maybe_ask_cash_method():
            return

        who = self.round_target = self._scope_name()
        self._ensure_browser_thread()
        self._set_busy(True, f"讀取{f'（{who}）' if who else ''}中，"
                             f"還沒登入的話瀏覽器會自己開起來，請不要關掉它…"
                             f"操作中不要改 Excel 的現金餘額。")
        self.browser_waiting += 1
        self.browser_cmd_queue.put((
            "fetch",
            (self._selected_accounts(), self.path,
             self.cash_method.get() == planner.METHOD_BANK),
        ))

    def start_logout_all(self):
        """
        「全部登出」：不管現在是誰的回合，把整個瀏覽器 session 收掉。

        跟登入/讀取不一樣，這顆不靠 excel_open 擋——沒開 Excel 也可能已經開著
        瀏覽器登入著（例如剛按過「登入」還沒選檔案），一樣該登得出去、關得掉。
        瀏覽器根本沒開過（執行緒不在或還沒起來）就不必勞師動眾丟一個指令進去，
        直接在這裡回話。
        """
        if self.busy:
            return
        if self.browser_thread is None or not self.browser_thread.is_alive():
            self._say("瀏覽器沒有開著，不用登出。")
            return
        if not ask_confirm(
                self.root,
                "登出並關閉瀏覽器",
                # 只留這一句（2026/09/01 使用者指定）。原本下面還接一句「瀏覽器裡任何你
                # 自己開的分頁（例如看盤視窗）也會一起關掉」——標題已經寫著「關閉瀏覽器」，
                # 那句是在解釋一件標題上已經講完的事。
                "確定要登出所有帳號並關閉瀏覽器嗎？",
                confirm_style="primary"):
            return

        self._set_busy(True, "登出中，瀏覽器會自己關掉，請稍候…")
        self.browser_waiting += 1
        self.browser_cmd_queue.put(("logout", None))

    def _browser_worker(self):
        """
        背景：整個瀏覽器 session 的生命週期都在這個執行緒裡，一直活到使用者自己把
        瀏覽器關掉，或整個介面關閉為止。

        每組帳號的分頁與 cookie 都收在這個執行緒手上的 store 裡（見 fetch.new_store）
        —— 「只更新某一位」能不重登就查得到資料，靠的就是它活得跟瀏覽器一樣久。

        不能每次按鈕都開新執行緒各開各的瀏覽器 —— Playwright 的同步 API 底層用
        greenlet 綁死建立它的那個執行緒，換一個執行緒去操作同一個 context 會直接
        壞掉（cannot switch to a different thread）。所以瀏覽器只能養在「一個專屬、
        活得夠久」的執行緒裡，「登入」「讀取」都只是丟一個指令進 queue
        給它處理，差別只在要不要順便查資料、更新 Excel。

        每一個指令都一定要回一則結果，這個執行緒也一定不能因為某次失敗就結束。
        主執行緒是收到那則結果才解除「登入中…」的 —— 不回話等於介面永遠卡在那裡，
        登入與讀取兩顆按鈕都是灰的，使用者只能關掉重開。所以 Playwright 改用
        start() 而不是 with：連「它自己起不來」（瀏覽器沒裝好、profile 正被另一個
        Chrome 佔著）都要變成一則看得到的「登入失敗」，而不是一個在背景執行緒裡
        炸掉、誰也看不到的例外。裝好瀏覽器之後再按一次就會重試，不必重開程式。

        登入與讀取只差在呼叫哪一個函式、回話時說自己是哪一種，所以走同一段程式碼
        —— 錯誤處理只有一份，不會有「登入那條路修好了、讀取那條還留著舊寫法」。
        """
        playwright = context = browser = None
        # 每組帳號的分頁、cookie、身分，還有「現在瀏覽器帶著誰的 cookie」。
        # 「一次只更新一位」全靠它：換人時把那一組登入時收下來的 cookie 換回去，
        # 不必重登（見 fetch.new_store）。跟著瀏覽器一起生、一起死。
        store = fetch_mod.new_store()
        # 下單分頁「依序執行」用的：填完一筆、開出委託確認視窗之後的那個 page，
        # 只有這個執行緒能碰它（見下面閒置輪詢那段跟 ui_order_exec._order_dialog_closed）。
        # 換帳號重開瀏覽器的每條路都要記得清掉它，不然舊帳號那頁早就不知道飛去
        # 哪裡了，還在照著它問「視窗關了沒」。
        order_watch_page = None

        def ensure_browser():
            nonlocal playwright, context, browser, store, order_watch_page
            if playwright is None:
                # 瀏覽器的位置要在 driver 起來之前決定好，它是靠環境變數傳下去的。
                configure_browsers_path()
                playwright = sync_playwright().start()
            if context is not None and not _browser_alive(context):
                context = browser = None      # 使用者自己把它關掉了，重開一個
            if context is None:
                context, browser = open_context(playwright)
                store = fetch_mod.new_store()
                order_watch_page = None       # 舊瀏覽器帶走的那一頁跟著作廢

        # 指令 -> (要呼叫誰, 回話時說自己是哪一種)
        jobs = {"login": (login_only, "logged_in"), "fetch": (collect, "fetched")}
        # 「借瀏覽器查一次、查完就結束」那一類，形狀跟 login/fetch 不一樣（參數
        # 不是 selected/path），但彼此之間一模一樣：ensure_browser() -> 呼叫 ->
        # 把回傳的 payload 送回主執行緒。「order」不在這裡——它要留一個 page 讓
        # 下面的閒置輪詢繼續盯著委託確認視窗，還有自己的 OrderMaybeSubmitted
        # 要分開處理，硬收進來只會讓這張表變成一堆例外。
        simple_jobs = {
            "order_quotes": (self._order_quotes_job, "order_quotes_fetched"),
            "pending": (self._pending_job, "pending_fetched"),
        }

        try:
            while True:
                try:
                    cmd, arg = self.browser_cmd_queue.get(timeout=0.15)
                except queue.Empty:
                    # 沒有指令要處理的時候也要繼續進出 Playwright，理由跟
                    # login.wait_until_finished 一樣：同步 API 只有在程式呼叫
                    # Playwright 時才會處理瀏覽器送來的事件，網站用 window.open
                    # 開出來的分頁（例如「即時看盤交易」「簡易看盤交易」）要等
                    # 事件被處理過才會開始載入。這裡沒有這一段的話，登入完、
                    # 使用者自己在瀏覽器裡點連結，新視窗會一直卡在 about:blank。
                    _pump_browser(context)
                    # 下單分頁「依序執行」在等這個：上一筆開出的委託確認視窗
                    # 有沒有真的關了。放在這個閒置分支裡輪詢，不是收到「order」
                    # 指令才檢查一次——「下一筆」按鈕要等這裡確認過才會解鎖
                    # （見 ui_order_exec.py「依序執行」那段對送錯帳戶風險的說明），
                    # 沒有人在這裡持續盯著的話，視窗關了畫面也不會知道。
                    if order_watch_page is not None and self._order_dialog_closed(order_watch_page):
                        order_watch_page = None
                        self.queue.put(("order_dialog_closed", {}))
                    continue
                if cmd == "stop":
                    break
                if cmd == "logout":
                    # clear_cookies 是這個網站定義的「登出」（見 login.do_login 換
                    # 帳號那段的說明）：JSESSIONID 這類 cookie 是伺服器唯一認人的
                    # 依據，清掉就等於把手上這批帳號一次全部登出。USER_DATA_DIR
                    # 指到的是持久化 profile 時尤其要做這一步——不清就直接關瀏覽器，
                    # cookie 還留在磁碟上，下次開起來網站可能直接把人當成還登入著。
                    payload = {}
                    try:
                        if context is not None:
                            try:
                                context.clear_cookies()
                            except PlaywrightError:
                                pass
                            context.close()
                        if browser is not None:
                            browser.close()
                    except Exception:
                        payload = {"error": traceback.format_exc()}
                    finally:
                        # 跟使用者自己把瀏覽器關掉走同一條路（見 ensure_browser 裡
                        # _browser_alive 判斷失敗那段）：context 收成 None，下一次
                        # 登入/讀取自己會重開一個乾淨的瀏覽器，不必特別處理這個狀態。
                        context = browser = None
                        store = fetch_mod.new_store()
                        order_watch_page = None
                    self.queue.put(("logged_out", payload))
                    continue

                if cmd == "order":
                    # 下單分頁的「開始下單／下一筆」，見 ui_order_exec.start_order_execution。
                    # 跟 login/fetch 那兩種不一樣：一次只做一筆委託、參數是
                    # (第幾組帳號, 帳號設定, 執行預覽裡的一列, 模式, 追價檔數,
                    # 是否自動送出, 買賣方向)，不是 selected/path 那種形狀，
                    # 所以不走下面 jobs[cmd] 那條共用路。模式／追價檔數／自動
                    # 與否／買賣方向都是這一輪按下「開始下單」那一刻凍結的值
                    # （見 start_order_execution），不是每筆下單前重讀畫面。
                    order_number, account, row, mode, ticks, auto, side = arg
                    payload = {}
                    try:
                        ensure_browser()
                        order_watch_page, extra = self._order_fill_job(
                            context, store, order_number, account, row, mode, ticks, auto, side)
                        payload.update(extra)
                    except order_fill.OrderMaybeSubmitted as exc:
                        # 「確認」已經真的按下去了，只是沒等到結果——這種不能讓
                        # 使用者以為跟一般失敗一樣「按下一筆重試就好」，重試會把
                        # 同一筆委託再送一次。maybe_submitted 這個旗標讓
                        # ui_order_exec._on_order_filled 用完全不同的文字警告使用者。
                        payload = {"error": str(exc), "maybe_submitted": True}
                    except RuntimeError as exc:
                        # _order_fill_job／fetch.ensure_logged_in／order_fill.select_stock
                        # 丟出來的其餘 RuntimeError 訊息本來就是寫給人看的（查不到
                        # 成交價、登入失敗、股票代號比對不符…），發生在真的按下
                        # 「確認」之前，重試安全，不是意外的臭蟲，不需要連
                        # traceback 一起丟到「這一筆下單失敗」的視窗上——那只會讓
                        # 真正的一句話被一大段檔案路徑、行號淹沒。
                        payload = {"error": str(exc)}
                    except Exception:
                        # 這裡才是真的沒預期到的狀況（Playwright 逾時、瀏覽器操作
                        # 失敗…），traceback 留著才查得出來是哪裡壞的。
                        payload = {"error": traceback.format_exc()}
                        if context is None:
                            payload["hint"] = BROWSER_HINT
                    self.queue.put(("order_filled", payload))
                    continue

                if cmd == "pending_cancel":
                    # 掛單分頁三顆取消按鈕的一則＝一個帳戶（見 ui_pending.
                    # _dispatch_next_cancel）。不收進下面 simple_jobs 那張表：
                    # 那張表是「一則指令跑完整批就結束」，而取消要一個帳戶一則
                    # 才停得下來（10.3 第六點），而且它跟下單一樣有
                    # OrderMaybeSubmitted 要分開處理。
                    order_number, account, sheet, committed, reservation = arg
                    payload = {"sheet": sheet}
                    try:
                        ensure_browser()
                        payload.update(self._pending_cancel_job(
                            context, store, order_number, account, sheet, committed, reservation))
                    except order_fill.OrderMaybeSubmitted as exc:
                        # 「確認」已經按下去了：那一批多半已經送到券商，絕對不能
                        # 被當成「這一則沒做，再按一次就好」。旗標讓
                        # _on_pending_cancelled 走完全不同的路（整批停下來）。
                        payload["error"] = str(exc)
                        payload["maybe_submitted"] = True
                    except RuntimeError as exc:
                        # order_cancel／order_cancel_reservation 與
                        # fetch.ensure_logged_in 丟的 RuntimeError 訊息本來就是
                        # 寫給人看的（核對不過、登入失敗、按鈕不見了…），而且都
                        # 發生在按下確認之前，不需要連 traceback 一起丟到畫面上。
                        payload["error"] = str(exc)
                    except Exception:
                        payload["error"] = traceback.format_exc()
                        if context is None:
                            payload["hint"] = BROWSER_HINT
                    self.queue.put(("pending_cancelled", payload))
                    continue

                if cmd == "order_odd_cancel":
                    # 出清零股跑完一輪、等 20 秒之後的自動撤單（見
                    # ui_order_exec._dispatch_next_odd_cancel）。形狀照
                    # pending_cancel 抄：一則指令＝一個帳戶（停止就是不派下一則）、
                    # 同樣有 OrderMaybeSubmitted 要分開處理，所以一樣不收進
                    # 下面 simple_jobs 那張表。
                    order_number, account, sheet, codes = arg
                    payload = {"sheet": sheet}
                    try:
                        ensure_browser()
                        payload.update(self._order_odd_cancel_job(
                            context, store, order_number, account, sheet, codes))
                    except order_fill.OrderMaybeSubmitted as exc:
                        # 刪單的「確認」已經按下去了：那一批多半已經送到券商，
                        # 不能被當成「這一則沒做，再派一次就好」。
                        payload["error"] = str(exc)
                        payload["maybe_submitted"] = True
                    except RuntimeError as exc:
                        # order_query／order_cancel／ensure_logged_in 丟的
                        # RuntimeError 訊息本來就是寫給人看的，而且都發生在按下
                        # 確認之前，不必連 traceback 一起丟到畫面上。
                        payload["error"] = str(exc)
                    except Exception:
                        payload["error"] = traceback.format_exc()
                        if context is None:
                            payload["hint"] = BROWSER_HINT
                    self.queue.put(("order_odd_cancelled", payload))
                    continue

                if cmd in simple_jobs:
                    # 「查完就結束」那一類：參數形狀各自不同（一批股票代號、一批
                    # 帳戶…），但錯誤處理與回話方式完全一樣，所以收成一張表，
                    # 之後再多一種查詢不必再複製一段 try/except。job 自己回傳
                    # 要送出去的 payload。
                    job, kind = simple_jobs[cmd]
                    payload = {}
                    try:
                        ensure_browser()
                        payload = job(context, store, *arg)
                    except RuntimeError as exc:
                        # 這幾支丟出來的 RuntimeError 訊息本來就是寫給人看的
                        # （登入失敗、查不到、身分對不上…），不需要連 traceback
                        # 一起丟到畫面上。
                        payload = {"error": str(exc)}
                    except Exception:
                        payload = {"error": traceback.format_exc()}
                        if context is None:
                            payload["hint"] = BROWSER_HINT
                    self.queue.put((kind, payload))
                    continue

                fetch_records, kind = jobs[cmd]
                # 「讀取」那條路多帶一個 need_bank（這次要不要查銀行餘額，見
                # fetch.collect）；「登入」沒有這個東西，所以用 *extra 收。
                selected, path, *extra = arg
                try:
                    ensure_browser()
                    records = fetch_records(context, selected, store, *extra)
                    if cmd == "login":
                        # 登入按鈕只做登入，不碰 Excel —— 現值要留給「讀取」去讀，
                        # 那邊本來就要讀一次，今日現金基準也是靠那次讀到的值設定（見 _on_fetched）。
                        payload = {"records": records}
                    else:
                        payload = _read_excel_after_fetch(records, path)
                except Exception as exc:
                    payload = {"error": traceback.format_exc()}
                    if excel_io.is_dead_object(exc):
                        payload["hint"] = excel_io.DEAD_EXCEL_HINT
                    elif context is None:
                        # 連瀏覽器都還沒開起來就失敗，那是這台電腦的環境問題，
                        # 不是網站或帳號的問題，講清楚才不會有人去重打密碼。
                        payload["hint"] = BROWSER_HINT
                self.queue.put((kind, payload))
        finally:
            try:
                if context is not None:
                    context.close()
                if browser is not None:
                    browser.close()
                if playwright is not None:
                    playwright.stop()
            # 收尾失敗沒有任何補救的餘地，而這裡通常是「介面正在關閉」——
            # 為了關不乾淨的瀏覽器再丟一個例外出去只會蓋掉真正的死因。
            except Exception:
                pass

    def _collect_writes(self):
        """
        把提案整理成「分頁 -> 要寫哪幾格」。

        要寫哪幾格完全由 planner 的 will_write 決定，介面不再插手。以前介面會把
        沒勾的就地改成不寫，那是紀錄檔跟 Excel 對不起來的唯一破口 —— 現在沒有
        這條路了，寫進 Excel 的跟記進紀錄檔的必定是同一批。

        只收 round_scope 裡那幾位。名單上其他人可能也有「要寫」的格子，但那是
        用上一輪的網頁資料算出來的 —— 按「讀取（王小明）帳戶」只會去查王小明，
        這時候順手把別人那幾格也寫進去，寫的是舊資料，而且沒有人要求過。
        """
        writes, total = {}, 0
        for name, items in self.proposals.items():
            if name not in self.round_scope:
                continue
            cells = [(item["row"], item["col"], item["formula"] or item["proposed"])
                     for item in items if item["will_write"]]
            if cells:
                writes[name] = cells
                total += len(cells)
        return writes, total

    def _begin_write(self, writes, total):
        self.write_count = total
        self._set_busy(True, f"寫入 {total} 格…")
        threading.Thread(target=self._write_worker, args=(self.path, writes), daemon=True).start()

    def _write_worker(self, path, writes):
        import pythoncom

        payload = {}
        excel = workbook = sheet = None
        pythoncom.CoInitialize()
        try:
            with excel_io.opened(path, True) as (excel, workbook, attached):
                for name, cells in writes.items():
                    sheet, error = excel_io.find_sheet(workbook, name)
                    if sheet is None:
                        raise RuntimeError(error)
                    excel_io.write_cells(sheet, cells)
                    if excel_io.marker_enabled():
                        excel_io.write_marker(sheet)
                # 一律存檔，接上使用者開著的 Excel 時也一樣。
                #
                # 原本接上時刻意不存、留給人自己按 Ctrl+S 當作多一道確認，但那道
                # 確認換來的是一個更糟的破口：紀錄檔在寫入「成功」之後就記成寫過了，
                # 人只要沒按 Ctrl+S（或關檔時選「不要儲存」），帳本就跟檔案分家，
                # 而且畫面上不會有任何徵兆。
                workbook.Save()
            payload = {"attached": attached}
        except Exception as exc:
            payload = {"error": traceback.format_exc()}
            if excel_io.is_dead_object(exc):
                payload["hint"] = excel_io.DEAD_EXCEL_HINT
        finally:
            sheet = excel = workbook = None
            pythoncom.CoUninitialize()

        self.queue.put(("written", payload))

    def _drain(self):
        """
        主執行緒每 120ms 看一次背景有沒有結果。widget 只在這裡被碰。

        每一筆各自包一層 try，而且「排下一次」放在 finally 裡。這個迴圈是背景與
        畫面之間唯一的通道：處理某一筆時丟出例外的話，Tk 只會把 traceback 印在
        console（打包成 exe 之後連 console 都沒有），然後最後那行排程就跳過了 ——
        取件迴圈從此停擺，之後每一次登入、讀取、寫入都會停在「進行中…」，
        而背景其實早就做完、結果就躺在 queue 裡沒有人取。與其這樣，寧可把出錯的
        那一筆報成失敗，讓迴圈活下去。
        """
        handlers = {"logged_in": self._on_logged_in, "fetched": self._on_fetched,
                    "written": self._on_written, "logged_out": self._on_logged_out,
                    "order_stock_list": self._on_order_stock_list,
                    "order_plans": self._on_order_plans_data, "order_filled": self._on_order_filled,
                    "order_dialog_closed": self._on_order_dialog_closed,
                    "order_price_refresh": self._on_order_price_refresh,
                    "order_odd_cancelled": self._on_order_odd_cancelled,
                    "order_rates": self._on_order_rates,
                    "excel_layout": self._on_excel_layout,
                    "order_quotes_fetched": self._on_order_quotes_fetched,
                    "pending_fetched": self._on_pending_fetched,
                    "pending_cancelled": self._on_pending_cancelled,
                    "macro_stuck": self._on_macro_stuck}
        try:
            while True:
                kind, payload = self.queue.get_nowait()
                try:
                    handlers[kind](payload)
                except Exception:
                    self._on_handler_error(kind)
        except queue.Empty:
            pass
        finally:
            self._check_browser_thread()
            self._sync_stop_all_button()
            self.root.after(120, self._drain)

    def _on_handler_error(self, kind):
        """
        收到背景結果之後，自己在處理時出錯（多半是程式的臭蟲）。

        一定要解除 busy：出錯的是「收尾」那一步，使用者按下的那顆按鈕還在等回音，
        不解除的話兩顆按鈕會一直是灰的，跟當掉沒有兩樣。

        也一定要說出來。這種錯發生在「已經做了一半」的時候 —— 網頁資料讀完了、
        Excel 可能也寫過了 —— 所以除了 traceback，還要請人去對一下歷程，
        那是唯一能看出實際做了什麼的地方。
        """
        detail = traceback.format_exc()
        step = STEP_NAMES.get(kind, kind)
        self._set_busy(False)
        self._say(f"「{step}」的後續處理出錯")
        show_error(self.root,
            "程式出錯",
            f"「{step}」的結果處理到一半出錯，這一批沒有做完。\n"
            f"Excel 與紀錄檔可能只完成了一部分，請切到「歷程」看實際做了哪幾格。\n\n"
            f"────────────────\n{detail[-1500:]}")

    def _check_browser_thread(self):
        """
        還在等背景回話、負責瀏覽器的那個執行緒卻已經不在了，就自己收尾。

        _browser_worker 已經把每一種失敗都包成一則結果回報了，這裡擋的是它自己
        也擋不住的那種（例如連 except 都還沒走完就整個結束）。沒有這道網的話，
        畫面會停在「登入中…」、登入與讀取都是灰的，使用者只能關掉重開 ——
        而關的時候還會被問一次「還在忙，確定要關嗎」。

        沒被取走的指令要一起倒掉。留著的話，下次按登入會先跑到那個舊指令
        （舊的檔案路徑、舊的帳號選擇），做了一件使用者這次沒有要求的事。

        下單「依序執行」在等委託確認視窗關閉的那段（order_exec_watching）另外
        算一種「還在等」：那個階段 order_filled 早就回過話了，browser_waiting
        已經歸零，只剩「等視窗關閉」這件事還沒有人跟畫面說完——執行緒要是剛好
        在這個節骨眼死掉，只看 browser_waiting 會判斷成「沒在等什麼」，「下一筆」
        就會永遠鎖死看不出原因。
        """
        if not self.browser_waiting and not self.order_exec_watching:
            return
        if self.browser_thread is not None and self.browser_thread.is_alive():
            return

        self.browser_waiting = 0
        while True:
            try:
                self.browser_cmd_queue.get_nowait()
            except queue.Empty:
                break

        # 下單「依序執行」正好卡在這裡的話，跟著一起收掉——瀏覽器都不在了，
        # 剩下的委託沒有任何辦法繼續，留著只會讓「下一筆」按鈕永遠是灰的。
        if self.order_exec_queue:
            self.order_exec_queue = []
            self.order_exec_pos = 0
            self.order_exec_busy = False
            self.order_exec_watching = False
            self._update_order_exec_ui()

        self._set_busy(False)
        self._say("背景作業中斷，這次的動作沒有做完")
        show_error(self.root,
            "背景作業中斷",
            "負責瀏覽器的背景作業結束了，這次的動作沒有做完。\n\n"
            f"{BROWSER_HINT}\n\n再按一次「登入」會重新開一個瀏覽器。")

    def _on_logged_in(self, payload):
        self.browser_waiting = max(0, self.browser_waiting - 1)
        self._set_busy(False)

        if "error" in payload:
            self._say("登入失敗")
            show_error(self.root, "登入失敗", _error_text(payload))
            return

        names, problems = [], []
        for record in payload["records"]:
            # 名字是登入才拿得到的東西，拿到就記著 —— 上面那個範圍選單、
            # 「讀取（某某）帳戶」那顆按鈕都靠這份對照。
            if record.get("sheet_name"):
                self.trader_of[record["order"]] = record["sheet_name"]
            if record["problems"]:
                problems.append(f"第 {record['order']} 組：" + "；".join(record["problems"]))
            elif record.get("sheet_name"):
                names.append(record["sheet_name"])
        self._refresh_account_choices()

        if problems:
            show_error(self.root, "登入失敗", "\n".join(problems))
            self._say("登入失敗")
            return

        self._say(f"已登入：{'、'.join(names)}。要更新資料時再按「更新」。")

    def _initialize(self, records):
        """
        把 Excel 上的現金餘額收成今日基準。回傳收了幾格。

        在「讀取」回呼裡呼叫，緊接在這一輪 Excel 現值讀回來（sheet_data
        更新）之後、真正拿網頁資料去算現金（replan）之前 —— 這一輪還沒寫過 Excel，
        所以這裡讀到的數字必定是「今天買賣之前」的狀態，現金基準可以直接取 B8、
        今天的流水先記 0，等 replan 再把這一輪查到的淨收付往上加。「B8 含不含
        今天的淨收付」這個最容易答錯的問題，在這個時間點根本不存在。

        判斷全在 planner.initialize()：一天只設一次現金基準，之後每一輪讀取都會
        呼叫這裡，但只有第一輪真的動得了 —— 介面不必自己記「今天設過了沒」。
        這裡只負責挑出「這一組這次讀到、而且 Excel 那一頁也真的讀到了」的分頁 ——
        這次沒讀到的一律跳過，沒讀到的東西不能拿來當起點。

        時間戳跟著 self.round_at 走，不再自己另外呼叫 datetime.now()（2026/08/24
        修正）——這裡跟 _commit_round 原本各叫各的 datetime.now()，登入到寫入中間
        隔著 Excel COM 寫入的好幾秒，兩邊蓋到的秒數幾乎不會一樣。ui_sync
        ._format_today_events 判斷「[今日初始餘額] 要不要跟這一輪的 [餘額更新]
        併在一起」靠的正是兩者 at 相不相等，at 對不起來就會兩行都印出來，銀行
        餘額推算那天畫面上多出一行講不出用途的「今日初始餘額」。同一輪本來就該
        算同一個時間點，改用 round_at 就是把這個假設落實。
        """
        if self.ledger is None:
            return 0

        events = []
        for record in records:
            name = record.get("sheet_name")
            data = self.sheet_data.get(name)
            if record["problems"] or not data:
                continue
            book = self.ledger.sheet(name)
            # 帳號代號在登入的當下就知道了，順手記進紀錄檔 ——
            # 使用者有可能登入完就沒有再按讀取。
            book["account_code"] = record.get("account_code", "")
            at = self.round_at.get(name) or datetime.datetime.now().isoformat(timespec="seconds")
            new_events, blocked = planner.initialize(data, book, name, self.today, at)
            events.extend(new_events)
            # B8 是空的，今日初始現金餘額設不成：讀幾次都一樣，不會自動好，
            # 掛一則 [異常] 讓人回頭去 Excel 補（見 ui_sync._fill_notes）。
            # 讀到非空的 B8 就代表補好了，把上一輪掛著的提醒收掉。
            if blocked:
                self.cash_baseline_errors[name] = {"text": "EXCEL 現金餘額是空白，無法設定今日初始餘額", "at": at}
            else:
                self.cash_baseline_errors.pop(name, None)

        if not events:
            return 0

        # 順序跟寫入那邊一樣：紀錄檔先落地，再追加歷程。
        self.ledger.save()
        self.ledger.append_history(events)
        self.refresh_history()
        return len(events)

    def _cash_item(self, name):
        """某個分頁的現金那一列提案，沒有就 None。"""
        return next((item for item in self.proposals.get(name, [])
                     if item["kind"] == "cash"), None)

    def _on_fetched(self, payload):
        self.browser_waiting = max(0, self.browser_waiting - 1)
        self._set_busy(False)

        if "error" in payload:
            self._say("讀取失敗")
            show_error(self.root, "讀取失敗", _error_text(payload))
            # 多輪出清叫來的那一輪同步（見 ui_order_exec._start_round_sync）：
            # 查不到網頁資料就判斷不出出清了沒，整批停下來，不拿舊持股硬跑。
            self._order_sync_finished(False, _error_text(payload))
            return

        # 這一輪讀到的一律是「補上去」，不是「整份換掉」：一次只更新一位的時候，
        # 名單上其他人手上那份是上一輪的資料，清掉他們等於整份名單只剩一個人。
        # 留著的代價是畫面上同時有好幾個時間點的資料——這件事原本靠訊息框第一行
        # 「讀取於 10:32:07」提醒，2026/08/22 使用者要求拿掉那一行（見
        # ui_sync._fill_notes），目前沒有別的畫面信號補這個位置。
        # 這一輪的時間戳，右邊常駐狀態列「✓ 與網頁一致」小提示、以及底下失敗
        # 原因要掛的時間都是它（見 ui.py 的 round_at 說明、ui_sync._fill_status／
        # _fill_notes）——同一輪裡大家一起讀，蓋同一個時間沒關係。要在下面這個
        # for 迴圈之前先算好，因為失敗原因當場就要記時間。
        now = datetime.datetime.now().isoformat(timespec="seconds")

        errors = payload["sheet_errors"]
        fresh = []
        for record in payload["records"]:
            order, name = record["order"], record.get("sheet_name")
            if name:
                self.trader_of[order] = name

            if record["problems"]:
                problem = "；".join(record["problems"])
            elif not name:
                problem = "讀不出這是哪一位交易人"
            elif name in errors:
                problem = errors[name]
            else:
                problem = None

            # 失敗原因跟著組別走：這一組這次成功就把上次的原因收掉，
            # 沒讀到的那幾組維持原樣 —— 別人的問題不會因為我這次讀成功就消失。
            # 文字不再自帶「第 N 組：」前綴（2026/08/22 使用者要求）：這則原因
            # 現在跟著這一組對應的交易人（見 ui_sync._fill_notes 用 trader_of
            # 反查名字）顯示在他自己的訊息框裡，人是誰已經在框的標題上了，前綴
            # 是重複資訊；真的還查不出是誰的那極少數情況，前綴改在顯示端現組
            # （見 ui_sync._fill_notes 的 fallback 那段）。
            if problem:
                self.problem_of[order] = {"text": problem, "at": now}
                continue
            self.problem_of.pop(order, None)
            self.records[name] = record
            fresh.append(name)

        # 這一輪只准碰這幾位。寫入、落帳全部照它 —— 別人手上那份是舊資料，
        # 拿舊資料去寫 Excel 是「一次只更新一位」最貴的一種錯。
        self.round_scope = set(fresh)
        for name in fresh:
            self.round_at[name] = now
        self.sheet_data.update(payload["sheets"])
        self._refresh_problems()
        self._refresh_account_choices()
        self.today = datetime.date.today()
        # 今日現金基準就在這裡定形（見 _initialize）：這一輪的 Excel 現值剛讀回來、
        # 還沒被這一輪的寫入動過，是「今天買賣之前」的最後一刻。
        self._initialize(payload["records"])
        self.replan()
        # 舊值要在任何寫入之前收好。寫入成功後 sheet_data 會被換成新數字，
        # 那時候再問「原本是多少」就沒有人記得了。只換這一輪讀到的那幾位：
        # 別人畫面上的「舊 → 新」是上一輪剛寫進去的結果，不該被這一輪抹掉。
        self.before = {key: value for key, value in self.before.items() if key[0] not in fresh}
        self.before.update({(name, item["cell"]): item["current"]
                            for name in fresh for item in self.proposals.get(name, [])})

        who = self.round_target
        note = self._problem_note()
        # 一位都沒讀成功的時候絕對不能說「已讀取」——「讀取（王小明）帳戶」按下去、
        # 他那一組登入逾時，畫面上其他人的數字全都還在，最像結論的那一句要是
        # 寫著「已讀取」，看的人不會知道自己看的是半小時前的東西。
        if not self.round_scope:
            self._say((f"{who} 這一次沒讀到，什麼都沒做。" if who
                       else "這一輪沒有一位對照得起來，什麼都沒做。") + note)
            # 一位都沒對照起來＝這一輪的持股完全沒更新，跟讀取失敗是同一種後果。
            self._order_sync_finished(False, "這一輪沒有一位對照得起來，持股沒有更新。" + note)
            return

        head = f"已讀取（{who}）。" if who else "已讀取。"
        if payload.get("attached"):
            self._say(head + "這個 Excel 正開著，程式會直接接上那個視窗寫入並存檔。" + note)
        else:
            self._say(head + note)

        # 讀完直接接著寫，中間不再問一次。按「讀取」本身就是意願的表達，
        # 再跳一個確認只是重複問同一件事。
        writes, total = self._collect_writes()
        if total:
            self._begin_write(writes, total)
        else:
            # 一格都不必寫，不代表沒事發生：今天的淨收付，是在這條路上落帳的。
            recorded = self._commit_round()
            # 順序不能反：refresh_history() 才會把剛落帳的這一筆讀回
            # self.history_rows，replan() 的 fill_sync_tree() 會用它畫訊息框
            # ——先 replan 再 refresh 的話，訊息框會晚一輪才看到這一筆
            # （2026/08/22 發現：模擬帳號改現金、按「更新全部帳戶」，餘額立刻
            # 更新但訊息框要按第二次才補上那一行）。
            self.refresh_history()
            self.replan()

            # 有人沒完成的時候，「一致」的範圍只到對照得起來的那幾位為止。
            # 主詞跟著縮小，這句話才不會替沒完成的那幾位背書。
            #
            # 看的是 proposals 不是 records：網頁讀到了、Excel 卻找不到那個
            # 分頁時 records 有東西、畫面上卻一位都沒有，這種時候說「一致」
            # 是把「沒得比」講成了「比過了」。
            scope = who if who else ("對照得起來的那幾位" if self.problems else "Excel 的數字")
            kept = f"紀錄檔更新了 {recorded} 筆（見歷程）。" if recorded else ""
            self._say(f"{head}{scope}跟網頁一致，沒有需要寫的格子。{kept}{note}")

            # 多輪出清那條路的另一個終點：一格都不必寫代表 Excel 本來就是最新的
            # （這一輪的委託一筆都沒成交），照樣可以回去判斷出清了沒——那個判斷
            # 本來就是「重讀一次 Excel 再算一遍」，不必真的寫過才算數。
            self._order_sync_finished(True)

    def _commit_round(self):
        """
        把這一輪的結果落實到紀錄檔：程式寫過的格子、現金的基準與流水。
        回傳追加了幾筆歷程。

        寫入成功之後一定要跑，而且只能在成功之後（見 _on_written）。但「一格都
        不必寫」的時候也一樣要跑 —— 那一輪照樣可能有話要記，最重要的一種是
        使用者剛在「重設現金餘額」填的答案：他填的開盤前金額加上今日淨收付，
        算出來剛好等於 B8 上的數字時（也就是 B8 早就含了今天的成交，正是那個
        對話框最該接住的情況），沒有任何一格需要寫。那個答案這時只存在提案裡，
        不落帳就會跟著整輪一起丟掉，下一次讀取拿沒被修正的基準再算一次，
        今天的淨收付就被加了第二次 —— 而畫面上不會有任何徵兆。

        時間戳跟 _initialize 一樣改用 self.round_at（2026/08/24 修正，理由見
        _initialize 的說明）——不再自己叫 datetime.now()，這一輪的 [今日初始
        餘額] 才會跟這裡寫出來的 [餘額更新] 落在同一個 at，訊息框的併行/去重
        判斷才對得起來。
        """
        if self.ledger is None:
            return 0

        events = []
        for name, items in self.proposals.items():
            # 跟 _collect_writes 同一個範圍。落帳記的是「這一輪發生了什麼」，
            # 沒去查的那幾位這一輪什麼也沒發生，尤其不能替他們記今天的淨收付。
            if name not in self.round_scope:
                continue
            book = self.ledger.sheet(name)
            at = self.round_at.get(name) or datetime.datetime.now().isoformat(timespec="seconds")
            events.extend(planner.commit(items, book, name, self.today, at))
        self.ledger.save()
        self.ledger.append_history(events)
        return len(events)

    def _on_written(self, payload):
        self._set_busy(False)

        if "error" in payload:
            self._say("寫入失敗")
            show_error(self.root, "寫入失敗", _error_text(payload))
            # 這一輪如果是多輪出清叫來的（見 ui_order_exec._start_round_sync），
            # 寫入失敗代表 Excel 上的持股還是舊的，下一輪不能跑。
            self._order_sync_finished(False, _error_text(payload))
            return

        # 紀錄檔一定在 Excel 寫成功之後才更新。順序反過來的話，寫入失敗會留下
        # 一份「以為自己寫過了」的帳本，現金基準會跟 Excel 上實際的數字對不起來。
        self._commit_round()

        # Excel 已經被改過了，手上的現值是舊的，重新讀一次才會準。
        # 只回填這一輪真的寫過的那幾位（範圍跟 _collect_writes 同一個）——
        # 別人那些「要寫」的格子這次沒寫進去，跟著改就會變成畫面說寫了、檔案沒有。
        for name, items in self.proposals.items():
            data = self.sheet_data.get(name)
            if not data or name not in self.round_scope:
                continue
            for item in items:
                if item["will_write"]:
                    if item["kind"] == "cash":
                        data["balance"] = item["proposed"]
                    else:
                        for line in data["rows"]:
                            if line["row"] == item["row"]:
                                line["qty" if item["which"] == "qty" else "cost"] = item["proposed"]

        # 順序不能反，理由同 _on_fetched 那個沒格子要寫的分支：refresh_history()
        # 要先把剛落帳的這一筆讀回 self.history_rows，replan() 畫訊息框才看得到。
        self.refresh_history()
        self.replan()

        # 刻意不跳視窗：一顆按鈕從頭做到尾，不要在最後又叫人按確定。
        #
        # 接上使用者開著的 Excel 時不再另外註明（2026/08/21 使用者要求）。
        # 那句話是寫給「以為自己開著的那份沒被改到」的人看的，而畫面上的數字
        # 本來就跟 Excel 一致、存檔也一律會做，兩種情況對人來說沒有差別。
        self._say(f"已自動寫入 {self.write_count} 格並存檔。{self._problem_note()}")

        # 多輪出清那條路的終點之一：Excel 的現金、股數、成本都已經是最新的了，
        # 可以回去判斷出清了沒（見 ui_order_exec._order_sync_finished）。平常按
        # 「更新」進來時這一支什麼都不做。
        self._order_sync_finished(True)

    def _on_logged_out(self, payload):
        self.browser_waiting = max(0, self.browser_waiting - 1)
        self._set_busy(False)

        if "error" in payload:
            self._say("登出失敗")
            show_error(self.root, "登出失敗", _error_text(payload))
            return

        self._say("已全部登出並關閉瀏覽器。下次按「登入」或「更新」會重新開一個瀏覽器。")

    def _refresh_problems(self):
        """
        把「第幾組 -> 失敗原因」攤平成畫面上那一串。

        原因記在組別上而不是每讀一次就整串重來：一次只更新一位的時候，別人上一輪
        沒完成的事並沒有因此解決 —— 那些 ⚠ 要留在畫面上，直到那一組自己再讀一次
        成功為止。照組別排序，畫面上的順序才不會每讀一次就跳一次。

        每一項帶著 `order`（不只 `text`／`at`）：ui_sync._fill_notes 要靠它
        反查 `self.trader_of` 才知道這則失敗屬於哪一位交易人，該顯示在誰的
        訊息框裡。
        """
        self.problems = [{"order": order, **self.problem_of[order]}
                         for order in sorted(self.problem_of)]

    def _problem_note(self):
        """
        這一輪有幾個帳戶沒完成。沒有就回空字串，直接串在狀態列後面。

        狀態列上任何一句「已讀取」「跟網頁一致」「已寫入」都要帶著它。20 組裡有
        一組登入逾時、或 Excel 少一個分頁時，那一位會直接從左邊名單上消失，而
        畫面上其他每一個數字看起來都正常 —— 提醒框裡雖然列著 ⚠，但狀態列同時
        說「跟網頁一致」的話，那句話會蓋過它：最像結論的那一句才是人會相信的。

        只報數量、不報細節。是哪幾組、為什麼失敗都在提醒框裡，狀態列只有一行，
        20 組全爛的時候塞不下，而且真正要看的細節本來就不該擠在那裡。
        """
        if not self.problems:
            return ""
        return f"⚠ 有 {len(self.problems)} 個帳戶沒完成，看下面的提醒。"

    def stop_all_operations(self):
        """
        底部狀態列那顆「停止」（見 docs/介面規劃.md 10.2）：跨分頁常駐，人按它的
        時候很可能正在掛單分頁看委託，而下單分頁裡那顆在別的分頁上按不到。

        停得掉的有兩件：下單分頁的依序執行，以及掛單分頁的整批取消。兩件的停法
        一模一樣——都是「不派下一則指令」，正在跑的那一則等它回話。登入／讀取／
        寫入／查詢那幾條沒有「停」這個動作可做，見 _browser_worker：每一個指令都
        一定要回一則結果，半路抽掉只會讓畫面永遠等不到回話。所以這顆按鈕只在真的
        有東西停得掉的時候才亮（見 _stop_all_available）。

        兩件不會同時發生（兩邊都借同一顆 self.busy），所以這裡不必決定先停誰。
        """
        if self.order_exec_active or self.order_exec_queue:
            self.stop_order_execution()
        elif self.pending_cancel_active:
            self.stop_pending_cancel()

    def _stop_all_available(self):
        """現在有沒有「停得掉」的東西。跟 order_exec_active 分開看的理由見那裡。"""
        return bool(self.order_exec_active or self.order_exec_queue
                    or self.pending_cancel_active)

    def _sync_stop_all_button(self):
        """
        由取件迴圈每 120ms 叫一次（見 _drain）——不是在每個狀態變動的地方各叫
        一次：那顆按鈕該亮不亮的條件散在依序執行那台狀態機的好幾個分支裡，
        漏掉一個就是「東西在跑但停不掉」。狀態沒變就不碰 widget，重畫成本是 0
        （見 docs/Tkinter ui設計原則.md 第十二節那份量測）。
        """
        want = "normal" if self._stop_all_available() else "disabled"
        if want != self._stop_all_state:
            self._stop_all_state = want
            self.stop_all_button.configure(state=want)

    def _excel_in_use(self):
        """
        現在有沒有任何一條路正在用 COM 動那份活頁簿。

        五個旗標各自誕生於不同的功能，本來各管各的：self.busy 是更新分頁的
        登入／讀取／寫入，order_busy 是下單分頁的「讀取試算」（2026/09/03 起
        觸發點是「新增」股票，見 ui_order.add_order_stock，這個旗標本身沒變），
        order_stock_list_busy 是「讀取ＯＯ持股」（只讀第一個分頁的候選清單），
        order_exec_price_busy 是多輪之間的重讀，order_rates_busy 是「執行帳戶」
        清單那趟只讀 B22 的補讀（見 ui_order.refresh_order_accounts）。問題是
        它們動的是**同一個 Excel 實例**——程式接上的是使用者眼前開著的那個
        （見 excel_io._open_once 的 GetObject 分支），不是各開各的一份。

        程式自己的讀寫都是限定寫法（sheet.Cells(...)），不受別人 Activate
        影響；但巨集用的是無限定的 Range()，只認 ActiveSheet。兩條執行緒交錯
        Activate 的話，巨集會跑在別人剛切過去的那一頁上——那一頁被更新兩次、
        自己這一頁從來沒更新過，而讀回來的是舊的 I4:I13，然後盤中追價就拿這個
        舊價當基準。不報錯、不缺欄位，只是靜靜地錯。

        所以五個旗標從這裡開始當成一個看。CLAUDE.md 那條「自動計算執行期間
        更新分頁的讀取／寫入要鎖住，反之亦然」就是靠這個述詞（畫面層）加上
        excel_io._EXCEL_LOCK（執行緒層）兩層一起實作的。
        """
        return bool(self.busy or self.order_busy or self.order_stock_list_busy
                    or self.order_exec_price_busy or self.order_rates_busy)

    def _apply_busy_state(self):
        """
        把「現在能不能碰 Excel」套到所有相關按鈕上。**上面 _excel_in_use() 那些
        旗標、以及 excel_open，任何一個變動都要叫一次**，不然畫面會停在上一個
        狀態：按鈕亮著、按下去卻被 guard 擋掉，看起來就是「按了沒反應」。
        """
        # 寫到一半換檔沒有意義，這一輪要動哪個檔早就決定了。
        self.excel_button.configure(state="disabled" if self._excel_in_use() else "normal")
        self._sync_buttons()
        self._order_excel_buttons()

    def _set_busy(self, busy, message=""):
        self.busy = busy
        self._sync_clear_button()
        if busy:
            self.progress.pack(side="right", padx=(8, 12), pady=6)
            # 100ms 一格，不是越快越好：ttkbootstrap 的進度條是用圖片畫的，每格
            # 都是一次重畫。2026/08/29 實測原本的 12ms（一秒 83 格，人眼分不出來）
            # 會讓動畫期間持續吃掉半顆核心（CPU 52%），而登入 20 組要跑好幾分鐘
            # —— 整個介面在最忙的時候反而最鈍。改成 100ms 之後降到 7%，看起來
            # 一樣在動。詳見 docs/Tkinter ui設計原則.md 的「ttkbootstrap 圖片元件的
            # 重畫成本」一節。
            self.progress.start(100)
            self._say(message)
        else:
            self.progress.stop()
            self.progress.pack_forget()
        self._apply_busy_state()

    def _say(self, message):
        self.status.configure(text=message)

    def _macro_stuck_notifier(self, macro, sheet_name):
        """
        給背景執行緒觸發巨集時當 `on_stuck` 傳（見 excel_io._run_macro_watched）。
        回傳的是一個不吃參數的 callable——真正呼叫它的是 `threading.Timer`
        自己的執行緒，那條執行緒不能碰 Tk widget，所以一樣繞 `self.queue`，
        跟其他背景結果收尾同一條路（見 `_drain`）。
        """
        return lambda: self.queue.put(("macro_stuck", {"macro": macro, "sheet": sheet_name}))

    def _on_macro_stuck(self, payload):
        """
        excel_io._run_macro_watched 的看門狗喊的：巨集卡在《payload['sheet']》
        跳出的對話框超過門檻秒數還沒回來（見 docs/介面規劃.md 9.6 第 2 點）。

        只是喊一聲，不做任何事——沒有辦法從這裡打斷還在跑的 Application.Run，
        真正卡住的話這句提示會一直留在常駐列上，直到那條背景執行緒自己
        因為使用者去 Excel 按了確定而回來，覆蓋掉這句話。
        """
        self._say(f"「{payload['macro']}」可能卡在《{payload['sheet']}》跳出的對話框，"
                   f"去 Excel 看一下。")

    def _path_text(self):
        """
        常駐列上「開啟EXCEL」旁邊那行字。只有兩句：沒開就是「EXCEL未開啟」，
        開著就是那份檔的路徑。

        2026/08/30 使用者要求收成這兩句。原本是三句（還沒選檔／選了沒開／開著），
        每一句都把路徑跟一段說明並排。拿掉的理由：
          - 「按「開啟EXCEL」」那句在講的事，左邊那顆按鈕自己就寫著了。
          - 還沒開起來的時候路徑不是使用者要看的東西——他要知道的只有「現在有沒有
            接上」。沒接上就別用一長串路徑去佔那一行。
          - 「還沒選檔」跟「選了但沒開」對使用者是同一件事：都要按同一顆按鈕。
            差別只在檔案對話框會不會幫他預選上次那一份，不值得多一種說法。
        路徑出現本身就等於「已經開著」，所以開著那句不必再補「已開在 Excel 裡」。
        """
        return str(self.path) if (self.excel_open and self.path) else "EXCEL未開啟"

    def open_excel(self):
        """
        選一份要同步的 Excel，用 Excel 把它開起來，路徑寫回 .env 下次直接用。

        為什麼是「開啟」不是「載入」
        --------------------------
        程式自己也開得起來（沒開著時 excel_io 會在背景 DispatchEx 一個隱形的），
        但 Office 沒啟用的機器上，那個背景 Excel 會在啟用檢查失敗後把自己收掉，
        同步是跑到一半才炸的 —— 那時搞不好已經寫了幾格。所以這裡改成
        由使用者親手把檔開起來，程式接上他那個視窗：看得見、讀到的是畫面上的
        即時內容、寫完他也馬上看得到。沒開起來就不給「更新」，把問題擋在最前面
        （2026/08/30 起「登入」不在這道關卡裡了，見 ui_sync._sync_buttons）。

        紀錄檔（現金基準）是跟著檔名走的，所以換一份檔等於換掉整個狀態來源
        —— 手上這批網頁資料與提案全是上一份檔算出來的，一律清掉重來，
        不能讓 A 檔的提案留在畫面上等著寫進 B 檔。
        """
        if self.busy:
            return

        start = self.path.parent if (self.path and self.path.parent.is_dir()) else app_dir()
        chosen = filedialog.askopenfilename(
            parent=self.root,
            title="選擇要同步的持股管理表",
            initialdir=str(start),
            initialfile=self.path.name if self.path else "",
            filetypes=[("Excel 活頁簿", "*.xls *.xlsx *.xlsm"), ("所有檔案", "*.*")],
        )
        if not chosen:
            return

        # 上一次錨點檢查的結論作廢：使用者按這顆按鈕，可能就是去改對了那份表、
        # 或改選另一份。留著的話，改對了也還是一直被擋住（見
        # _on_excel_layout／_poll_excel）。
        self.excel_layout_problem = None

        path = Path(chosen)
        # 選同一份檔不必重來一遍，但還是要確認它開著 —— 使用者按這顆按鈕，
        # 想要的就是「把它打開」，不是「什麼都沒發生」。
        if path == self.path:
            self._open_in_excel(path)
            return

        try:
            ledger = ledger_mod.Ledger(path)
        except RuntimeError as exc:
            show_error(self.root, "紀錄檔有問題", str(exc))
            return

        try:
            excel_io.remember_excel_path(path)
        except OSError as exc:
            show_warning(self.root, "沒寫進 .env",
                         f"這次可以用，但路徑沒記起來，下次要再選一次：\n{exc}")
        self.path = path
        self.ledger = ledger
        self.ledger_error = None
        self.ledger_fresh = not ledger.existed
        # excel_open 記的還是上一份檔的答案（通常是 True），而輪詢要三秒後才會發現
        # 換人了。中間那幾秒畫面在說謊：路徑列寫著新檔的路徑、後面卻接「已開在
        # Excel 裡」，登入也還亮著。那時候按下去，程式會自己開一個看不見的 Excel
        # 來讀這個檔，而 os.startfile 正在把同一個檔開進使用者看得見的那個 Excel
        # —— 兩邊搶同一個檔案，總有一邊拿到唯讀。
        # 亮回來是 _open_in_excel 的事：它每半秒確認一次，真的看到開起來才放行。
        self._set_excel_open(False)
        self.path_label.configure(text=self._path_text())

        # 手上這一輪全是上一份檔算出來的，跟著整批作廢（見 _forget_round）。
        self._forget_round()
        # 換檔就回到預設的「初始餘額累加」（算法不再跟著紀錄檔跨天沿用，
        # 見 _default_method）；「這次執行問過了」是按路徑記的，所以換到這次執行
        # 還沒問過的檔，下次讀取會再問一次。
        self.cash_method.set(self._default_method())
        self._refresh_method_label()

        self.fill_sync_tree()
        self.refresh_history()
        if not ledger.existed:
            self._say(f"已換成 {path.name}（這份檔還沒有紀錄檔，下次登入會以 Excel "
                      f"現在的數字設定今日現金基準）")
        self._open_in_excel(path)

    def _open_in_excel(self, path, tries=0):
        """
        把檔案交給 Excel 打開，然後等它真的開起來。

        用 os.startfile 而不是 COM：這等於使用者自己雙擊那個檔，Excel 是他的、
        視窗是看得見的，程式手上不留任何 COM 物件。真要同步時 excel_io 會用
        GetObject 接上這個視窗。

        不能開下去就當作成功 —— startfile 只是把檔案丟給系統，Excel 起來要幾秒，
        中間還可能卡在受保護的檢視或啟用巨集的提示上。所以這裡每半秒問一次檔案鎖，
        真的看到它被鎖住才把登入放行。等待期間不能用 sleep 卡住主執行緒，
        不然整個介面會凍在那裡，連「正在開啟」那行字都畫不出來。
        """
        if path != self.path:
            return                        # 等待期間又換了檔，這條等待就作廢
        if excel_io.is_open_in_excel(path):
            self._set_excel_open(True)
            self._say(f"{path.name} 已經開在 Excel 裡，可以按「登入」了。")
            return

        if tries == 0:
            try:
                os.startfile(str(path))
            except OSError as exc:
                self._set_excel_open(False)
                show_error(self.root, "開不起來", f"沒辦法用 Excel 打開這個檔：\n{path}\n\n{exc}")
                return
            self._say(f"正在用 Excel 開啟 {path.name}…")
        if tries >= 60:                   # 30 秒
            self._set_excel_open(False)
            self._say(f"還沒看到 {path.name} 開起來。")
            show_warning(self.root,
                "沒看到 Excel 開起來",
                f"{path.name} 交給 Excel 了，但 30 秒過去還沒看到它被開啟。\n\n"
                f"Excel 可能卡在「受保護的檢視」或啟用提示上，也可能還在啟動。\n"
                f"請看一下 Excel 那邊；等它真的開好，登入就會自己亮起來。")
            return
        self.root.after(500, lambda: self._open_in_excel(path, tries + 1))

    def _poll_excel(self):
        """
        每三秒確認一次「這份 Excel 現在開著沒」，然後排下一次。

        為什麼要一直問：使用者隨時可能把 Excel 關掉，而按鈕的亮暗說的是「現在」
        能不能做，不是「按開啟那一刻」的狀態。檔案鎖很便宜（開一個檔案代號就關掉），
        也不碰 COM，所以固定輪詢比在十幾個地方各補一次檢查可靠。
        """
        try:
            here = self.path is not None and self.path.is_file() and excel_io.is_open_in_excel(self.path)
        except OSError:
            here = self.excel_open        # 判斷不出來就沿用上次，不要亂閃
        # 版面對不上的那一份，就算真的開在 Excel 裡也一律當成「不能用」——這道
        # 旗標要在這裡看，不是只在 _on_excel_layout 設一次 excel_open：三秒後
        # 這支就會照「檔案開著」把它扳回 True（見 _on_excel_layout 的說明）。
        self._set_excel_open(here and not self.excel_layout_problem)

        # Excel 一關，畫面上那些數字就地作廢。它們是「這份 Excel 開著的時候，
        # 從它讀出來的現值跟網頁比出來」的結果 —— 檔一關就沒有東西替它們背書了：
        # 人可能是在 Excel 裡自己改完數字才關的，也可能關掉之後換一份檔再開回來。
        # 按鈕變灰只擋得住下一步，擋不住有人照著螢幕上那個舊餘額去下單。
        #
        # 忙的時候先不動：那一輪還會把結果填回畫面（讀取完就緊接著寫入），
        # 這裡清掉只是被它蓋回去。輪詢三秒一次，忙完的下一次自然會補上。
        if not here and not self.busy and self._has_round_data():
            self._forget_round()
            self.fill_sync_tree()
            self._say(f"{self.path.name} 關掉了 —— 畫面上的數字是它開著的時候算的，"
                      f"已經清空。按「開啟EXCEL」重新開起來，再讀一次。")
        self.root.after(3000, self._poll_excel)

    def _set_excel_open(self, value):
        """狀態有變才動畫面 —— 每三秒重畫一次按鈕會讓游標懸停的效果一直閃。"""
        if value == self.excel_open:
            return
        self.excel_open = value
        self.path_label.configure(text=self._path_text())
        # 走 _apply_busy_state() 而不是只叫 _sync_buttons()：下單分頁的「讀取ＯＯ
        # 持股」「讀取試算」也卡在 excel_open 上（見 ui_order._order_excel_buttons），
        # 只更新更新分頁那幾顆的話，Excel 一關它們會一直亮著。
        self._apply_busy_state()
        # 接上的那一刻才讀得到那份表裡有誰——下單分頁的「執行帳戶」整格都是
        # Excel 那一頭的答案（分頁名＋B22，見 refresh_order_accounts 與
        # _forget_round）。這裡是換檔那條路唯一「新路徑 ＋ 真的開著」同時成立
        # 的地方，也是「開啟EXCEL 之後帳戶就自己出現」靠的那一下。
        if value:
            # 接上的那一刻先驗錨點（2026/09/02 使用者要求：「開啟EXCEL的時候
            # 可以先檢查錨點」）。順序上它跟 refresh_order_accounts 是同時出發
            # 的兩條背景執行緒，靠 excel_io.opened() 那把鎖排隊——版面對不上的
            # 話，帳戶清單那一趟本來就會讀回空清單（沒有一個分頁的錨點對得上，
            # 見 excel_io.list_account_sheets），不會有「清單有人但版面是錯的」
            # 這種半調子狀態。
            self.check_excel_layout()
            self.refresh_order_accounts()

    def check_excel_layout(self):
        """
        開啟 EXCEL（或使用者自己把檔開起來）之後的**錨點檢查**：這份活頁簿是不是
        程式認得的那一版持股管理表（A22 ＝「今年報酬率」，見 excel_io.layout_problem）。
        2026/09/02 使用者要求。

        對不上就把 excel_open 扳回 False，等於整支程式的 Excel 功能都停在那裡
        （更新、讀取ＯＯ持股、讀取試算、寫入的按鈕全跟著 excel_open 走）——這不是「提醒一下
        還是可以用」的等級：版面對不上代表程式要寫的 E/F 欄、要讀的 M/N 欄
        全部落在別人的格子上，而且不會報錯。檔案本身不動（不會去關掉使用者的
        Excel 視窗），人自己去開對的那一份。

        不算進 `_excel_in_use()`：它只讀幾格、不 Activate、不跑巨集，跟別條
        COM 路的互斥交給 `excel_io.opened()` 那把鎖就夠了。算進去反而會讓
        「開啟EXCEL 之後那一兩秒」整排按鈕閃一下灰。
        """
        if self.path is None or self.excel_layout_busy:
            return
        self.excel_layout_busy = True
        threading.Thread(target=self._excel_layout_worker,
                         args=(self.path,), daemon=True).start()

    def _excel_layout_worker(self, path):
        """背景執行緒：接上活頁簿、問一次 layout_problem（見 check_excel_layout）。"""
        import pythoncom

        pythoncom.CoInitialize()
        workbook = None
        payload = {}
        try:
            with excel_io.opened(path, False) as (_excel, workbook, _attached):
                payload = {"path": str(path), "problem": excel_io.layout_problem(workbook)}
        except Exception as exc:
            payload = {"error": str(exc)}
        finally:
            workbook = None
            pythoncom.CoUninitialize()
        self.queue.put(("excel_layout", payload))

    def _on_excel_layout(self, payload):
        """
        錨點檢查的回話。

        讀不到（"error"）**不擋**：那是別的問題——Excel 剛好又被關掉、正在被別的
        程式鎖著之類，輪詢那條路自己會發現。這裡只處理「真的讀到了、而且對不上」
        這一種，因為那是唯一「檔案開得好好的、看起來一切正常、寫下去卻會寫錯格」
        的情況。

        回來時路徑已經換掉的話整份丟掉：這一趟驗的是上一份檔。
        """
        self.excel_layout_busy = False
        if "error" in payload or payload.get("path") != str(self.path):
            return
        problem = payload["problem"]
        if problem is None:
            return
        # 記在旗標上而不是只把 excel_open 設 False：三秒一次的輪詢看到檔案還開著
        # 就會把它扳回 True（見 _poll_excel），只設一次擋不住。旗標在使用者下次
        # 按「開啟EXCEL」時清掉，那時會重驗一遍。
        self.excel_layout_problem = problem
        self._set_excel_open(False)
        self._say("開到的持股管理表版面對不上，已經擋住——請開 10 家那一版。")
        show_error(self.root, "開錯持股管理表", problem)

    def _has_round_data(self):
        """畫面上（左邊名單、右邊明細、訊息框）現在有沒有東西。"""
        return bool(self.records or self.sheet_data or self.proposals or self.problems)

    def _forget_round(self):
        """
        把「這一份 Excel 的這一輪」整批忘掉：網頁資料、Excel 現值、提案、提醒、
        舊值、這一輪的範圍。換檔（open_excel）與 Excel 被關掉（_poll_excel）都走這裡。

        畫面不在這裡重畫 —— 呼叫的人接下來還有別的事要先做（換檔要先把路徑與
        紀錄檔換好），由他自己挑時機叫 fill_sync_tree()。

        trader_of 不清：那是「第幾組帳號是哪一位交易人」的對照，跟開哪一份 Excel、
        Excel 開著沒都無關，瀏覽器那邊也還登著，下次讀取不必重登。

        範圍退回「全部」：手上一位的資料都不剩了，下一步一定是重讀一輪，
        停在某一位身上的話，「讀取（某某）帳戶」按下去只會讀回半份名單。
        """
        self.records, self.sheet_data, self.proposals = {}, {}, {}
        self.warnings, self.problems = {}, []
        self.before = {}
        self.round_at = {}
        self.problem_of, self.round_scope = {}, set()
        self.round_target = None
        self.current_sheet = None
        self.account_choice.current(0)
        self._refresh_fetch_button()
        # 下單分頁那幾份也是「這一份 Excel 的」：持股、試算、股價、報酬率全部
        # 來自分頁上的格子，換一份檔案就全數作廢。報酬率還多一層——它決定
        # 「執行帳戶」清單本身（有誰、排序、號碼），留著舊檔的名單會讓人對著
        # 另一份表裡的人下單。下一次 Excel 接上時會重讀（見 _set_excel_open）。
        self.order_return_rates = {}
        # 只重畫選單，不在這裡去讀新檔的 B22：這支被呼叫的時候 Excel 一定是
        # 「還沒接上」的狀態（換檔那條路剛把 excel_open 設成 False，Excel 被關掉
        # 那條路更不用說）。補讀交給 _set_excel_open——接上的那一刻才是能讀的
        # 那一刻。
        self._fill_order_accounts()
        # 持股、試算、股價那幾份跟「改勾選」要清的東西一模一樣，走同一支，不在
        # 這裡再列一次（列兩份遲早會有一邊漏掉一項）。差別是這裡連「指定股票」
        # 也要清掉（keep_stocks=False）：換了一份 Excel，候選清單本身就換了一份，
        # 留著等於拿舊檔的股票去對新檔的分頁。勾選的那幾位由 _fill_order_accounts
        # 順手清掉——名單都換了，舊名字不會出現在新清單上。
        self._clear_order_round(keep_stocks=False)

    def _require_excel(self):
        """
        讀取（更新分頁）與「讀取ＯＯ持股」「讀取試算」（下單分頁）之前的最後
        一道關卡。按鈕平常是灰的，這裡擋的是鍵盤觸發那種漏網。

        「Excel 沒開著」那一種刻意不跳視窗，只是擋下來（2026/08/31 使用者要求）：
        這三顆按鈕現在都跟著 excel_open 變灰（見 ui_sync._sync_buttons、
        ui_order._order_excel_buttons），常駐列上又寫著「EXCEL未開啟」、旁邊就是
        「開啟EXCEL」那顆按鈕——畫面已經把這件事講完了，再彈一個視窗要人按掉是
        在重複同一句話。剩下兩種（沒選檔、檔案不見了）還是要說：那是畫面上看不
        出來的事。
        """
        if self.path is None:
            show_error(self.root, "還沒選檔案", "請先按左上角「開啟EXCEL」選一份持股管理表。")
            return False
        if not self.path.is_file():
            show_error(self.root, "找不到 Excel", f"{self.path}\n\n可以在 .env 用 EXCEL_PATH 指定位置。")
            return False
        return bool(self.excel_open)

    def _default_method(self):
        """
        現金算法的起始值：每次程式開起來（換一份 Excel 也一樣）都是「初始餘額累加」。

        2026/08/21 使用者要求，取代原本「選了哪一種就寫進紀錄檔跨天沿用」。理由是
        兩種算法錯的方式不對稱：全額交割當天留在銀行餘額推算會把同一筆錢扣兩次，
        而且是直接寫進 Excel、隔天才看得出來；初始餘額累加最壞是漏掉匯撥與股利，
        數字對不起來當下就看得到、也補得回來。所以「忘了換」要往安全的那一邊倒，
        不能讓某一天的特例自己延續到之後每一天。
        """
        return planner.METHOD_OPENING

    def _refresh_method_label(self):
        """
        算法變了（或換了 Excel）就把右邊常駐狀態列重畫一次 —— 「現金算法」寫的
        就是它，而「今日初始現金餘額」與底下那顆「修改」在不在也跟著它走
        （哪一項要畫、按鈕收不收，全在 ui_sync._fill_status / _show_opening_row 判）。

        名字這次執行還沒問過就不畫，但基準與「修改」相反，還沒問過也要照沿用
        下來的算法決定 —— 名字可以先不講（那只是還沒有答案），但一顆按下去會
        蓋掉 B8 的東西不能等到問完才收起來。兩件事都在 _fill_status 裡。
        """
        self._fill_right()

    def _toggle_cash_method(self):
        """
        點一下「現金算法」那個名字：換成另一種。只有已經問過、名字露臉之後才點得到。

        先跳一個確認視窗 —— 這顆管全部交易人、換錯會直接蓋掉現金那一格的算法，
        誤觸的代價比多一次點擊高。預設答案是「取消」，跟 ui_history 清除歷程
        那顆同一個規矩：有風險的操作，手滑按下 Enter 也不該真的動到。

        confirm_style="primary"：焦點鎖「否」、Enter 觸發取消的安全機制不變，
        只是「是」的顏色 2026/08/22 使用者要求跟同一天新增的「今天的現金餘額
        怎麼算」那顆藍色「確定」統一，不用預設的警示橘色（見 ask_confirm）。
        """
        other = (planner.METHOD_OPENING if self.cash_method.get() == planner.METHOD_BANK
                  else planner.METHOD_BANK)
        if not ask_confirm(
                self.root,
                "切換現金算法",
                f"現金算法要從「{planner.METHOD_NAMES[self.cash_method.get()]}」"
                f"換成「{planner.METHOD_NAMES[other]}」嗎？\n\n"
                f"全額交割當天要留在「{planner.METHOD_NAMES[planner.METHOD_OPENING]}」。",
                confirm_style="primary"):
            return
        self._set_method(other, asked=True)

    def _set_method(self, method, asked=False):
        """
        換現金算法：記進紀錄檔、重算、重畫。asked=True 代表這次是使用者自己選的
        （對話框的「確定」或工具列點一下切換都算），不必等下次讀取再跳一次視窗問。

        選了哪一種只在這次執行內有效，不寫進紀錄檔 —— 下次開程式一律回到
        「初始餘額累加」（見 _default_method）。「這次問過了」同樣只放記憶體，
        那是為了不要同一次執行裡跳第二次。

        訊息框不必在這裡另外清——它現在直接讀歷程檔（見 ui_sync._fill_notes），
        換算法不會讓任何舊資料留在畫面上，_refresh_method_label／replan 各自
        重畫一次就已經是最新狀態。
        """
        self.cash_method.set(method)
        if asked:
            self.cash_method_asked.add(self.path)
        self._refresh_method_label()
        if self.proposals:
            self.replan()

    def _maybe_ask_cash_method(self):
        """
        這次程式開起來、第一次按「更新全部帳戶」時，
        先問今天要用哪一種算法。回傳 True 才可以繼續讀取，False 表示使用者取消，
        呼叫端（start_fetch）要整個收手，不能當作已經問過。

        問在抓資料之前：算法決定要不要去查銀行餘額那兩支，選在後面就得再讀一次。
        它也不需要任何資料才問得出口 —— 判斷依據是「今天有沒有買全額交割股」，
        那是人自己知道的事。

        故意不記進紀錄檔、只放記憶體：問過一次是為了同一次執行裡不要每讀一次就跳
        一次（變成一個要人閉著眼睛按掉的東西），但關掉程式重開就當作沒問過 ——
        使用者可能在中間發現今天有全額交割、想換答案。

        2026/08/26 之前這個視窗沒有取消：X 跟 Escape 都關不掉。改回可以取消是因為
        使用者只是想按「登入」，卻按到旁邊那一顆——當時「登入」跟這一顆疊在同一直行，
        而且第一次讀取前它的標籤是「登入+讀取全部帳戶」，兩顆都有「登入」兩個字。
        視窗跳出來又走不掉，只能把整支程式關掉重開。那個版面 2026/08/30 拆開了
        （「登入」搬到分頁列上面的常駐列，這一顆只叫「更新全部帳戶」），誤按的機會
        小很多，但「跳出來的時機不是使用者要的」還是需要退路，所以取消留著。取消
        不會套用任何算法（answer 是 None），也不記進 cash_method_asked——下次再按
        「更新全部帳戶」都會重新問一次，不會卡在「問過但沒答案」的中間狀態。
        """
        if self.ledger is None or self.path in self.cash_method_asked:
            return True

        picked = ask_cash_method(self.root, self.cash_method.get())
        if picked is None:
            return False
        self._set_method(picked, asked=True)
        return True
