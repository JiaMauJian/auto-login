"""
下單分頁的行為：「盤前」（股票／比重／價格設定）跟「盤中」（股票／比重／
追價檔數設定，價格用 Excel 成交價＋下單前查對手方第一檔算出來的）共用
同一套勾帳戶、算執行預覽、依序半自動填單的機制，只有「股票設定要填什麼」
「怎麼組出執行清單」不一樣（見 `_on_order_mode_changed`／`start_order_execution`）。

半自動填單只做到「開出委託確認視窗」為止，不會按裡面的「確認」——那一步
要送出真實委託，留給人自己決定，見 order_fill.py／`_order_fill_job`。

比重→張數、帳戶依 B17 報酬率排序、組出預覽清單、追價檔數換算價格，全部是
orders.py 的純函式，這裡只負責收使用者的輸入、讀 Excel（含成交價，盤中
新增股票／重新整理時順便觸發「更新股價」巨集）、查即時對手方第一檔、
操作瀏覽器、把結果畫出來。
"""

import threading
import tkinter as tk
from tkinter import messagebox

import ttkbootstrap as ttk
from playwright.sync_api import Error as PlaywrightError

import excel_io
import fastquote
import fetch as fetch_mod
import order_fill
import orders
from ui_common import FONT_SIZE, PRICE_PENDING_TEXT, ask_confirm, col_width, wide
from util import show

# 「多輪直到出清」的安全上限：不管有沒有真的出清，跑滿這個輪數就一定停下來
# 等人看過再決定要不要繼續，不無限跑下去——2026/08/28 使用者確認要做這個
# 開關的時候沒特別要求上限數字，這裡抓一個「明顯比一般情境需要的輪數多，
# 但出問題也不會跑太久」的值，不是照文件或使用者指定的數字。
ORDER_MULTI_ROUND_CAP = 10


class UiOrderMixin:
    # ---------- 下單分頁：盤前模式 ----------

    def _order_init_state(self):
        """SyncApp.__init__ 呼叫一次。"""
        self.order_rows = []              # 這一輪加進來的股票設定列（見 add_order_stock）
        self.order_holdings = {}          # (分頁名, 股票代號) -> 股數，「重新整理」才會更新
        self.order_names = {}             # 股票代號 -> 名稱，畫面顯示用
        self.order_prices = {}            # 股票代號 -> Excel I 欄讀回來的股價；盤中模式這份就是
                                           # chase_price 的 pricenow 來源（見 start_order_execution），
                                           # 不只是畫面顯示用（跟 order_names 平行）
        self.order_return_rates = {}      # 分頁名 -> B17 報酬率或 None（讀不到）
        self.order_account_vars = {}      # 分頁名 -> 勾選狀態的 BooleanVar（見 _fill_order_accounts）
        self.order_busy = False
        # 股票代號 -> {"bid","ask","last"}，「查詢委買賣」按鈕整批查回來的即時
        # 委買賣一（見 fetch_order_quotes／fastquote.FastQuoteStream.latest()
        # 的形狀）。有這份資料時 orders.plan_intraday_orders 會直接算出實際
        # 會送出的價格，不是只留一句「下單前會再查」的說明文字（2026/08/29
        # 使用者要求）。切模式會清空重來，理由跟 order_rows 一樣。
        self.order_quotes = {}
        self.order_quotes_busy = False
        self._order_quotes_requested = []  # 上一次按「查詢委買賣」實際問了哪幾檔，回話時算漏了誰用
        self.order_stock_price_busy = False  # 盤中「新增」股票附帶觸發的股價重讀還在跑（見 _refresh_added_stock_price）

        # 「盤前」／「盤中」模式（見 ui_layout._build_order_tab）。追價檔數
        # 是盤中模式整批共用的一個值（使用者 2026/08/28 確認過，不是像比重
        # 那樣每檔股票各自設定），trace 讓改了立刻反映在執行預覽上，跟股票
        # 設定列的 weight/price 是同一個做法。
        self.order_mode = tk.StringVar(value="pre")
        self._order_mode_last = "pre"
        self.order_ticks = tk.StringVar(value="2")
        self.order_ticks.trace_add("write", lambda *_a: self._recompute_order_preview())

        # 買／賣：整批共用一個方向（見 orders.SIDE_SELL 的說明），跟「盤前」
        # ／「盤中」同一顆開關並排。切換的時候比照切模式，把股票清單清空
        # 重選（2026/08/28 使用者確認），不是想辦法把舊的列轉成新方向——
        # 「比重」在買賣兩個方向意義不完全一樣（買方向一樣是「比重×目前
        # 持股」，只是完全沒持有的股票還是會被判定沒有這檔），沿用舊的比重
        # 數字換方向繼續用不會出錯，但清空重選比較不會讓人誤會兩邊比重是
        # 同一件事。
        self.order_side = tk.StringVar(value=orders.SIDE_SELL)
        self._order_side_last = orders.SIDE_SELL

        # 自動送出開關：關（預設）＝半自動，跟原本一樣停在委託確認視窗給人看、
        # 給人按；開＝程式自己按下確認視窗裡的「確認」，委託真的送出去，不會
        # 停下來等人。2026/08/28 使用者要求加的，AskUserQuestion 確認過節奏
        # 不變——不管開不開，還是「下一筆」按一次只處理一筆，差別只在「這
        # 一筆處理完」是靠人看過按確認，還是程式自己按。
        self.order_auto_confirm = tk.BooleanVar(value=False)
        self._order_auto_last = False

        # 半自動送單（見 start_order_execution）：這一輪凍結起來要依序送出的
        # 委託清單（orders.executable_orders／executable_intraday_orders 過濾過
        # 的執行預覽），跟 order_rows／order_holdings 這些會一直變動的畫面狀態
        # 分開——按下「開始下單」那一刻就定案，之後就算使用者又改了比重或
        # 勾選，也不會半路影響正在跑的這一輪。order_exec_mode／order_exec_ticks
        # ／order_exec_auto 也是同一輪按下去那一刻凍結的（不是每筆下單前重讀
        # 畫面上的即時值），理由跟凍結 queue 一樣。
        self.order_exec_queue = []
        self.order_exec_pos = 0        # 下一筆（或正在處理的這一筆）在 queue 裡的位置
        self.order_exec_busy = False   # 背景正在登入／填單，還沒回話
        self.order_exec_watching = False  # 委託確認視窗開著，等它關閉才能按「下一筆」
        self.order_exec_mode = "pre"
        self.order_exec_ticks = None
        self.order_exec_auto = False
        self.order_exec_side = orders.SIDE_SELL
        self.order_exec_last_note = ""    # 上一筆自動送出的結果，_update_order_exec_ui 顯示用

        # 多輪直到出清（見 _order_round_finished／_prepare_next_round）：跟
        # order_exec_mode 這些一樣，是按下「開始下單」那一刻凍結的值，不是
        # 每輪重讀畫面。order_exec_stock_settings／order_exec_accounts 也要
        # 凍結——第 2 輪以後要拿同一組股票設定、同一批帳戶去對「重讀之後的
        # 新持股」重新組隊列，不是每輪重新看畫面上現在勾了誰、填了什麼。
        self.order_multi_round = tk.BooleanVar(value=False)
        self.order_auto_price = tk.BooleanVar(value=False)
        self.order_exec_multi_round = False
        self.order_exec_auto_price = False
        self.order_exec_round = 1
        self.order_exec_stock_settings = []
        self.order_exec_accounts = []
        # 代號 -> Excel 讀回來的股價，盤中模式 _order_fill_job 算追價時一律
        # 拿這裡的值當 pricenow，不現查網頁成交價（2026/08/29 使用者確認：
        # 成交價已經在新增股票／重新整理時讀進 Excel，不必等下單前才查）。
        # 開始下單那一刻先用 self.order_prices（見那裡的說明）當第 1 輪的
        # 起始值（start_order_execution），之後每輪重讀就整份換掉，不是
        # 累加（同一檔股票這一輪的價格只有一個版本）——order_exec_auto_price
        # 這個開關現在只決定重讀前要不要先觸發「更新股價」巨集，不再決定
        # pricenow 走哪條路（見 _on_order_price_refresh）。
        self.order_exec_prices = {}
        # 開始下單那一刻凍結的 self.order_quotes 快照，只給第 1 輪用——第 2
        # 輪以後在 _on_order_price_refresh 會清成空字典，逼所有列都退回
        # 「下單前才查」那條舊路，不讓第 1 輪查到的即時報價被沿用到之後幾輪
        # （市場已經過了一段時間，繼續當最新報價用是在猜數字）。
        self.order_exec_quotes = {}
        self.order_exec_price_busy = False  # 輪與輪之間正在重讀 Excel／觸發巨集，還沒回話
        # 「這一整批多輪出清作業還在不在跑」，跟 order_exec_queue 分開——
        # 輪與輪之間 queue 會是空的（下一輪還沒組出來），但整批作業並沒有
        # 結束，「停止」要能在這個空檔也按得下去（見 stop_order_execution）。
        self.order_exec_active = False

    def refresh_order_data(self):
        """
        重新讀 Excel：把每個已知名字的帳戶分頁的持股（E/F 欄）、B17、股價
        （I 欄）都讀一次。

        盤中模式每個分頁在讀它之前先各觸發一次「更新股價」巨集再讀
        （2026/08/29 使用者確認：盤中追價用的成交價來自這裡讀到的
        order_prices，新增股票／重新整理這一步就要盡量拿到新的價格，不能
        留著上次殘留的舊數字）；盤前模式的價格是人手動填的，不需要 Excel
        股價，維持原本只讀不觸發巨集。

        帳戶多的時候這一步會明顯變慢：巨集是一檔股票打一次 Yahoo 的 HTTP
        請求，5 檔 × N 個分頁全部逐一發出去，這是慢不是當掉。

        帳戶名單只能從 self.trader_of 來——那是「登入過才知道名字」的既有
        規則（見 ui.py），還沒登入過的帳戶這裡也看不到，跟同步分頁的範圍
        選單是同一個限制，不是這裡另外加的。
        """
        # 看 _excel_in_use() 而不是只看 order_busy：同步分頁的寫入、「新增」股票
        # 附帶的股價重讀、多輪之間的重讀，動的都是同一份活頁簿（見那個述詞）。
        if self._excel_in_use() or not self._require_excel():
            return
        names = sorted(set(self.trader_of.values()))
        if not names:
            messagebox.showinfo(
                "還沒有帳戶名字",
                "還沒有任何帳戶登入過，名字都還不知道。\n請先到「同步」分頁按「登入」。",
                parent=self.root)
            return

        run_macro = self.order_mode.get() == "intraday"
        self.order_busy = True
        self._apply_busy_state()
        self.order_status.configure(text="更新股價、讀取中…" if run_macro else "讀取中…")
        threading.Thread(target=self._order_read_worker, args=(self.path, names, run_macro), daemon=True).start()

    def _order_read_worker(self, path, names, run_macro):
        """
        背景執行緒：用 COM 讀 E/F 欄、B17、I 欄。run_macro 為真的話，每個
        分頁在讀它之前先各觸發一次使用者既有的「更新股價」巨集——是「每個
        分頁各一次」不是「整批一次」，那個巨集只認 ActiveSheet（見
        excel_io.run_update_price_macro）。跟 _order_price_refresh_worker
        同一個做法，見那邊 write=run_macro 為什麼要一起傳的說明。
        """
        import pythoncom

        pythoncom.CoInitialize()
        excel = workbook = sheet = None
        payload = {}
        try:
            with excel_io.opened(path, run_macro) as (excel, workbook, _attached):
                sheets, errors = {}, {}
                with excel_io.keep_active_sheet(workbook):
                    for name in names:
                        sheet, error = excel_io.find_sheet(workbook, name)
                        if sheet is None:
                            errors[name] = error
                            continue
                        # 巨集只認 ActiveSheet，所以每一頁都要各 Activate 一次、
                        # 各跑一次，不能整批只呼叫一次（見
                        # excel_io.run_update_price_macro 說明的那個 bug）。
                        # 同一個理由，這一整段也不能跟別條執行緒同時跑——
                        # excel_io.opened 那把鎖擋的就是這件事。
                        if run_macro:
                            excel_io.run_update_price_macro(excel, sheet)
                        data = excel_io.read_sheet(sheet)
                        data["return_rate"] = excel_io.read_return_rate(sheet)
                        sheets[name] = data
                payload = {"sheets": sheets, "errors": errors}
        except Exception as exc:
            payload = {"error": str(exc)}
        finally:
            sheet = excel = workbook = None
            pythoncom.CoUninitialize()
        self.queue.put(("order_data", payload))

    def _on_order_data(self, payload):
        self.order_busy = False
        self._apply_busy_state()

        if "error" in payload:
            self.order_status.configure(text="讀取失敗")
            messagebox.showerror("讀取失敗", payload["error"])
            return

        self.order_holdings, self.order_return_rates, self.order_names, self.order_prices = {}, {}, {}, {}
        for name, data in payload["sheets"].items():
            self.order_return_rates[name] = data["return_rate"]
            for row in data["rows"]:
                self.order_holdings[(name, row["code"])] = row["qty"]
                self.order_names.setdefault(row["code"], row["label"].split("(")[0].split("（")[0].strip())
                # 哪個帳戶先讀到就先用哪個，跟 _on_order_price_refresh 彙整
                # order_exec_prices 同一個態度——同一檔股票的 Excel 股價不會
                # 因為帳戶不同而不同，不比對多帳戶是否一致。讀不到（None）
                # 就不佔位，讓 add_order_stock 那邊看到「沒有」而不是猜一個值。
                if row["price"] is not None:
                    self.order_prices.setdefault(row["code"], row["price"])

        choices = sorted(f"{code} {name}" for code, name in self.order_names.items())
        self.order_stock_pick.configure(values=choices)
        self._fill_order_accounts()

        errors = payload["errors"]
        note = f"　（{len(errors)} 個帳戶讀不到：{'、'.join(errors)}）" if errors else ""
        self.order_status.configure(text=f"已讀取 {len(payload['sheets'])} 個帳戶的持股與報酬率。{note}")
        self._recompute_order_preview()

    def _fill_order_accounts(self):
        """
        重畫帳戶勾選區：整份銷毀重建，不是想辦法保留舊的勾選狀態——目前只有
        「重新整理」會呼叫這裡，一輪通常只按一次，不值得為了這個情境先做。

        名字後面直接接百分比，不再重複寫「今年報酬率」——上面標題已經講過
        「依今年報酬率由低到高排序」，每一列再寫一次是多餘的字。
        """
        for child in self.order_account_inner.winfo_children():
            child.destroy()
        self.order_account_vars = {}
        for name in sorted(self.order_return_rates):
            rate = self.order_return_rates[name]
            # B17 存的是小數（0.185222... 代表 18.5%），畫面要顯示的是百分比，
            # 這裡要乘 100——漏了這一步會把 18.5% 顯示成 0.2%，跟現金查詢
            # Amount 除以 100 是同一種「單位不對但不會報錯」的坑（CLAUDE.md）。
            text = f"{name}　{rate * 100:.1f}%" if rate is not None else f"{name}　報酬率讀不到"
            var = tk.BooleanVar(value=False)
            self.order_account_vars[name] = var
            ttk.Checkbutton(self.order_account_inner, text=text, variable=var,
                           command=self._recompute_order_preview).pack(anchor="w", pady=1)
        self._resize_order_sheet_column()

    # ---------- 模式切換 ----------

    def _on_order_mode_changed(self):
        """
        切「盤前」／「盤中」。兩邊股票設定列的欄位形狀不一樣（盤中沒有價格，
        價格是整批共用的「追價檔數」，見 orders.plan_intraday_orders），與其
        想辦法把舊的列轉成新形狀，不如整批清掉重來——這跟「追價檔數整批共用
        一個值」是使用者同一次確認過的決定，兩種模式的股票清單本來就不該
        混用。
        """
        if self.busy:
            self.order_mode.set(self._order_mode_last)
            messagebox.showinfo("忙碌中", "現在有背景工作在跑，先等它結束才能切換模式。",
                                parent=self.root)
            return

        self._order_mode_last = self.order_mode.get()
        intraday = self._order_mode_last == "intraday"

        self.order_ticks_entry.configure(state="normal" if intraday else "disabled")
        # 「查詢委買賣」是盤中限定功能，跟 order_ticks_entry 同一個道理
        # （disabled 不整個藏起來）。切模式代表股票清單整批清掉重來（下面），
        # 舊查到的報價沒有對象可用，一併清空——不留著一份查不到任何一列在
        # 用的舊資料。
        self.order_quotes = {}
        self._update_order_quotes_ui()

        # 多輪直到出清／自動更新股價是盤中限定的功能（規劃文件「是否要跑
        # 多輪」只列在盤中設定底下，2026/08/28 使用者更正）。切到盤前就強制
        # 關掉、鎖住兩個勾選框；切回盤中才解鎖多輪那顆——自動更新股價那顆
        # 還是要等使用者自己勾多輪才會跟著解鎖（見 _on_order_multi_round_changed），
        # 不是切模式就自動打開。
        self.order_multi_round_check.configure(state="normal" if intraday else "disabled")
        if not intraday:
            self.order_multi_round.set(False)
            self.order_auto_price.set(False)
            self.order_auto_price_check.configure(state="disabled")

        for row in list(self.order_rows):
            row["frame"].destroy()
        self.order_rows = []
        self._resize_order_stock_column()
        self._recompute_order_preview()
        # 「新增」能不能按跟模式有關（見 _order_excel_buttons），切模式要重算一次。
        self._apply_busy_state()

    def _on_order_side_changed(self):
        """
        切「買」／「賣」。跟 _on_order_mode_changed 同一個做法：整批清掉股票
        清單重選，不是想辦法沿用舊的比重（2026/08/28 使用者確認，見
        _order_init_state 開頭 order_side 的說明）。
        """
        if self.busy:
            self.order_side.set(self._order_side_last)
            messagebox.showinfo("忙碌中", "現在有背景工作在跑，先等它結束才能切換買賣。",
                                parent=self.root)
            return

        self._order_side_last = self.order_side.get()

        for row in list(self.order_rows):
            row["frame"].destroy()
        self.order_rows = []
        self._resize_order_stock_column()
        self._recompute_order_preview()

    def _on_order_multi_round_changed(self):
        """
        「多輪直到出清」勾／不勾。「自動更新股價」是它的子選項——沒勾多輪，
        自動更新股價這件事根本不會發生（只跑一輪，沒有「下一輪開始前」這個
        時間點），所以子選項要跟著鎖起來、順便清掉，不留一個勾了但沒作用
        的狀態讓人誤會。
        """
        if not self.order_multi_round.get():
            self.order_auto_price.set(False)
        self.order_auto_price_check.configure(
            state="normal" if self.order_multi_round.get() else "disabled")

    def _on_order_auto_changed(self):
        """
        切「半自動」／「自動送出」。半自動一直是這裡的預設、也是目前唯一
        實測過整條路能通的模式（見記憶 order-exec-sequential-wired-up）；
        切到自動那一刻要跳一個夠重的警告——這不是畫面選項，是「程式會自己
        按下真的會送出委託的按鈕」，跟其他「按錯了大不了重選」的設定不是
        同一個等級的風險，要讓使用者確認過才生效，而不是勾了就算。
        """
        if self.busy:
            self.order_auto_confirm.set(self._order_auto_last)
            messagebox.showinfo("忙碌中", "現在有背景工作在跑，先等它結束才能切換。",
                                parent=self.root)
            return

        if self.order_auto_confirm.get() and not ask_confirm(
                self.root, "切換成自動送出",
                "切換為自動送出後，按下「下一筆」時，程式會自動按下委託確認視窗"
                "中的「確認」，委託將直接送出，不再等待人工確認。\n\n"
                "執行節奏不變，仍為逐筆處理（每按一次「下一筆」僅送出一筆），"
                "但委託一經送出即無法收回——如需取消或修改，須自行至「委託查詢」"
                "／「預約查詢」頁面操作，程式不會代為處理。\n\n"
                "確定要切換為自動送出嗎？",
                confirm_style="primary"):
            self.order_auto_confirm.set(False)

        self._order_auto_last = self.order_auto_confirm.get()

    def _order_excel_buttons(self):
        """
        下單分頁裡「按下去會用 COM 動 Excel」的按鈕：只要有任何一條路正在動那份
        活頁簿就一起變灰（見 ui_background._apply_busy_state，它負責在四個旗標
        變動時呼叫這裡）。

        原本這兩顆各自只看自己那一個旗標——「持股與報酬率」跑著的時候「新增」
        還是亮的，而兩條路都會一頁一頁 Activate 再跑巨集，交錯之後巨集會跑在
        別人剛切過去的那一頁上（見 excel_io._EXCEL_LOCK 的說明）。

        擋住而不是排隊：跟 _refresh_added_stock_price 對自己重複點擊的態度一致
        （那裡的註解有寫理由——下一次「新增」或「重新整理」還會再有機會補上）。
        """
        busy = self._excel_in_use()
        self.order_refresh_button.configure(state="disabled" if busy else "normal")
        # 「新增」只有盤中那條路會附帶跑巨集（見 add_order_stock 的說明），盤前
        # 完全不碰 COM，沒有理由跟著變灰——讀取 20 組帳戶要跑好幾分鐘，那段時間
        # 還是該能把股票加進清單。所以這一顆多看一個模式。
        add_busy = busy and self.order_mode.get() == "intraday"
        self.order_add_button.configure(state="disabled" if add_busy else "normal")

    def _order_ticks_setting(self):
        """
        盤中模式的「追價檔數」，讀不懂（不是 0 以上的整數）回 None——不猜、
        不偷偷代成規劃文件講的預設值 2，打錯了要讓使用者自己看到、自己改，
        不能被程式默默帶過去（跟 fetch.settle_problem「讀不懂就整格擋住」
        同一種態度）。
        """
        try:
            ticks = int(self.order_ticks.get().strip())
        except ValueError:
            return None
        return ticks if ticks >= 0 else None

    # ---------- 股票設定 ----------

    def add_order_stock(self):
        """
        把下拉選單（或手動輸入）裡的股票加進設定清單，一檔一列。盤前模式
        比重／價格各自獨立輸入（使用者確認過：不同股票想賣的比例、價格
        通常不一樣，共用一個值沒意義）；盤中模式只有比重，價格由整批共用
        的「追價檔數」在下單當下算出來，這一列不需要 price 這個欄位。

        盤中模式額外觸發一次背景的「更新股價」（見 _refresh_added_stock_price）
        ——這裡顯示的 Excel 股價是加進清單那一刻的快照（見 _build_order_stock_row
        的說明），剛加的這一檔如果原本沒被最近一次「重新整理」涵蓋到（例如
        本來沒持股），不補這一步就會一直停在讀不到／舊的數字。
        """
        raw = self.order_stock_pick.get().strip()
        if not raw:
            return
        code = raw.split(" ")[0].strip().upper()
        if any(row["code"] == code for row in self.order_rows):
            messagebox.showinfo("已經加過了", f"{code} 已經在清單裡了。", parent=self.root)
            return
        name = self.order_names.get(code) or (raw.split(" ", 1)[1].strip() if " " in raw else code)

        row = {"code": code, "name": name, "weight": tk.StringVar()}
        row["weight"].trace_add("write", lambda *_a: self._recompute_order_preview())
        if self.order_mode.get() == "pre":
            row["price"] = tk.StringVar()
            row["price"].trace_add("write", lambda *_a: self._recompute_order_preview())
        self._build_order_stock_row(row)
        self.order_rows.append(row)
        self.order_stock_pick.set("")
        self._resize_order_stock_column()
        self._recompute_order_preview()
        if self.order_mode.get() == "intraday":
            self._refresh_added_stock_price()

    def _refresh_added_stock_price(self):
        """
        盤中模式「新增」股票時附帶觸發一次「更新股價」巨集、重讀 Excel I 欄
        （2026/08/29 使用者要求）。

        刻意不共用 refresh_order_data／_on_order_data 那條路——那邊會整個
        重建帳戶勾選框（見 _fill_order_accounts 的說明：「目前只有『重新
        整理』會呼叫這裡，一輪通常只按一次」），如果「新增」一檔股票也走
        同一條路，使用者每加一檔股票，已經勾好的帳戶就會被清空重建一次。
        這裡只更新 self.order_prices、刷新畫面上已加入股票的價格文字，不碰
        帳戶勾選、持股、報酬率。

        order_stock_price_busy 是這條路自己的忙碌旗標，跟 order_busy（重新
        整理）分開——短時間連續按好幾次「新增」，這裡選擇跳過而不是排隊，
        反正下一次「新增」或「重新整理」還會再有機會補上。
        """
        # 這裡改看 _excel_in_use()：原本只看自己那一個旗標，所以「持股與報酬率」
        # 正在跑（5 檔 × N 個分頁的 HTTP，很慢）的時候按「新增」就會起第二條
        # 執行緒，兩邊都在 Activate → 跑巨集 → Activate → 跑巨集。
        if self._excel_in_use() or not self.excel_open:
            return
        names = sorted(set(self.trader_of.values()))
        if not names:
            return
        self.order_stock_price_busy = True
        self._apply_busy_state()
        threading.Thread(target=self._order_stock_price_worker, args=(self.path, names), daemon=True).start()

    def _order_stock_price_worker(self, path, names):
        """
        背景執行緒：每個分頁各觸發一次「更新股價」巨集、重讀 I 欄，跟
        _order_price_refresh_worker 同一個做法（一頁一次的理由見
        excel_io.run_update_price_macro）。
        """
        import pythoncom

        pythoncom.CoInitialize()
        excel = workbook = sheet = None
        payload = {}
        try:
            with excel_io.opened(path, True) as (excel, workbook, _attached):
                sheets = {}
                with excel_io.keep_active_sheet(workbook):
                    for name in names:
                        sheet, error = excel_io.find_sheet(workbook, name)
                        if sheet is not None:
                            # 一頁一次，理由同 _order_read_worker。
                            excel_io.run_update_price_macro(excel, sheet)
                            sheets[name] = excel_io.read_sheet(sheet)
                payload = {"sheets": sheets}
        except Exception as exc:
            payload = {"error": str(exc)}
        finally:
            sheet = excel = workbook = None
            pythoncom.CoUninitialize()
        self.queue.put(("order_stock_price", payload))

    def _on_order_stock_price(self, payload):
        """
        _refresh_added_stock_price 的背景回話。讀不到／出錯就默默放棄、維持
        畫面上原本的股價——這只是「新增」附帶的加值，不是使用者當下在等的
        主要操作，不值得為了它彈錯誤視窗（真的要查，「重新整理」還在）。
        """
        self.order_stock_price_busy = False
        self._apply_busy_state()
        if "error" in payload:
            return
        for data in payload["sheets"].values():
            for row in data["rows"]:
                if row["price"] is not None:
                    self.order_prices[row["code"]] = row["price"]
        for row in self.order_rows:
            label = row.get("price_label")
            if label is None:
                continue
            excel_price = self.order_prices.get(row["code"])
            price_text = f"Excel股價 {show(excel_price)} 元" if excel_price is not None else "Excel股價：讀不到"
            label.configure(text=price_text)

    def _build_order_stock_row(self, row):
        """
        一檔股票一區塊，分兩行：股票名稱＋買賣別在上面，比重（／價格）在下面。
        分兩行是因為左邊這個窄面板一行塞不下「名稱＋比重＋價格＋移除」，
        會把價格輸入框擠到剩沒幾個像素、打不進去字——用 pack 而不是 grid
        編號，移除中間一列時不會留下空位，不必自己重新排列剩下的列。

        買賣別跟著畫面上目前的「買」／「賣」設定走（見 order_side），一整批
        股票共用同一個方向，加進來那一刻就定案——切買賣的時候整批清單會被
        清空重選（見 _on_order_side_changed），不會出現舊列還留著舊方向的
        情況。底色跟網站本身買紅賣綠的配色一致（Sell.TLabel／Buy.TLabel 在
        ui_layout._build() 裡註冊），不必看文字就認得出方向。

        `"price" in row` 決定要不要畫價格輸入框——盤中模式的 row 沒有這個
        key（見 add_order_stock），不是留白也不是畫一個不會被讀的欄位。

        盤中模式沒有價格輸入框的位置改顯示 Excel 讀回來的股價（self.
        order_prices，跟「重新整理」讀回來的 order_names 同一批資料）——
        這不只是給人參考，開始下單那一刻會拿 order_prices 當第一輪
        chase_price 的 pricenow（見 start_order_execution），追價檔數還是
        要在下單前用這個基準再算一次邊界、查一次對手方第一檔（見
        orders.chase_price），這裡顯示的數字就是實際會拿去算價的那一個
        （不是另一條網頁現查的路，2026/08/29 使用者確認拿掉了）。這裡存了
        Label 物件本體（row["price_label"]）而不是只畫一次文字，因為
        _refresh_added_stock_price 每次有人按「新增」都會觸發一次背景重讀
        （2026/08/29 使用者要求），回來要能就地更新這一列的文字，不是只有
        新加的那一列，而是畫面上全部盤中列一起刷新——單按「重新整理」不會
        觸發這個更新，只有「新增」股票才會。
        """
        side_text = "買" if self.order_side.get() == orders.SIDE_BUY else "賣"
        side_style = "Buy.TLabel" if self.order_side.get() == orders.SIDE_BUY else "Sell.TLabel"

        block = ttk.Frame(self.order_stock_frame)
        block.pack(fill="x", pady=(0, 8))

        head = ttk.Frame(block)
        head.pack(fill="x")
        ttk.Label(head, text=side_text, style=side_style, width=2, anchor="center").pack(side="left")
        ttk.Label(head, text=f" {row['code']} {row['name']} ", style=side_style).pack(side="left")
        ttk.Button(head, text="移除", bootstyle="danger-outline",
                  command=lambda: self.remove_order_stock(row)).pack(side="left", padx=(6, 0))

        fields = ttk.Frame(block)
        fields.pack(fill="x", pady=(4, 0))
        ttk.Label(fields, text="比重").pack(side="left")
        ttk.Entry(fields, textvariable=row["weight"], width=6,
                 font=(self.family, FONT_SIZE)).pack(side="left", padx=(4, 0))
        ttk.Label(fields, text="%").pack(side="left", padx=(2, 12))
        if "price" in row:
            ttk.Label(fields, text="價格").pack(side="left")
            ttk.Entry(fields, textvariable=row["price"], width=8,
                     font=(self.family, FONT_SIZE)).pack(side="left", padx=(4, 0))
            ttk.Label(fields, text="元").pack(side="left", padx=(2, 0))
        else:
            excel_price = self.order_prices.get(row["code"])
            price_text = f"Excel股價 {show(excel_price)} 元" if excel_price is not None else "Excel股價：讀不到"
            label = ttk.Label(fields, text=price_text, style="Hint.TLabel")
            label.pack(side="left", padx=(4, 0))
            row["price_label"] = label
        row["frame"] = block

    def remove_order_stock(self, row):
        row["frame"].destroy()
        self.order_rows.remove(row)
        self._resize_order_stock_column()
        self._recompute_order_preview()

    # ---------- 執行預覽 ----------

    def _selected_order_accounts(self):
        return [
            {"sheet": name, "return_rate": self.order_return_rates.get(name)}
            for name, var in self.order_account_vars.items() if var.get()
        ]

    def _order_stock_settings(self):
        """
        把畫面上股票清單的目前輸入值（比重，盤前模式再加價格）整理成
        orders.plan_stock_orders／plan_intraday_orders 吃的格式。給
        _recompute_order_preview 跟 start_order_execution 共用——後者要用
        「跟畫面上一模一樣」的設定去組執行清單，不是自己另外算一次。
        """
        stock_settings = []
        for row in self.order_rows:
            try:
                weight = float(row["weight"].get())
            except ValueError:
                weight = 0
            setting = {"code": row["code"], "name": row["name"], "weight_pct": weight}
            if "price" in row:
                setting["price"] = row["price"].get()
            stock_settings.append(setting)
        return stock_settings

    def _recompute_order_preview(self):
        """
        比重／價格（或追價檔數）／勾選的帳戶，任何一個變了就整份重算重畫——
        跟同步分頁 fill_sync_tree() 同一個做法，整份重建比自己追蹤哪一列該
        更新可靠。
        """
        ordered, skipped = orders.order_accounts(self._selected_order_accounts())
        stock_settings = self._order_stock_settings()
        side = self.order_side.get()
        hints = []

        if self.order_mode.get() == "intraday":
            ticks = self._order_ticks_setting()
            if ticks is None:
                preview = []
                hints.append("⚠ 追價檔數要填 0 以上的整數。")
            else:
                preview = orders.plan_intraday_orders(
                    stock_settings, ordered, self.order_holdings, ticks, side,
                    prices=self.order_prices, quotes=self.order_quotes)
        else:
            preview = orders.plan_stock_orders(stock_settings, ordered, self.order_holdings, side)

        if skipped:
            hints.append(f"⚠ {'、'.join(a['sheet'] for a in skipped)} 讀不到報酬率，沒有排進執行順序。")
        self._render_order_preview(preview, hints)

    def _render_order_preview(self, preview, hints):
        """
        把一份執行預覽（orders.plan_stock_orders／plan_intraday_orders 的回傳值）
        畫進 Treeview，跟 hints 一起蓋掉現在畫面上的內容。

        跟 _recompute_order_preview 拆開是因為多輪出清（見 _on_order_price_refresh）
        每跑完一輪重新讀了持股之後，也要用同一套畫法把「下一輪還剩什麼」畫出來
        給人看——那個 preview 是用凍結的 order_exec_stock_settings／
        order_exec_accounts 算出來的，不是重新讀畫面上現在的設定，不能共用
        _recompute_order_preview 整支（那支一開頭就會去讀畫面上的即時設定）。
        """
        for item in self.order_preview.get_children():
            self.order_preview.delete(item)

        side_names = {"B": "買", "S": "賣"}
        for item in preview:
            # 跳過的列不上買賣底色，只淡化文字——已經跳過了，不該看起來像
            # 真的會發生的一筆交易（見 ui_layout._build_order_right 的 tag_configure）。
            tag = "skip" if item["skip"] else {"B": "buy", "S": "sell"}.get(item["side"], "")
            # 盤中模式的 price 通常是 None（還沒按「查詢委買賣」，或那檔沒查
            # 到），畫面上顯示文字說明，不是空白也不是猜一個數字；查到了才是
            # orders.chase_price 算出來的數字，用 show() 補千分位，跟 Excel
            # 股價那句「Excel股價 {show(excel_price)} 元」同一個格式。盤前
            # 模式的 price 是使用者自己打的字串，原樣顯示，不套 show()。
            if item["price"] is None:
                price_text = PRICE_PENDING_TEXT
            elif isinstance(item["price"], str):
                price_text = item["price"]
            else:
                price_text = show(item["price"])
            self.order_preview.insert("", "end", values=(
                item["order"], item["sheet"], side_names.get(item["side"], item["side"]),
                f"{item['code']} {item['name']}",
                # 持股最小單位是 1 股，不需要小數點；股數本來就可能上看百萬，
                # 千分位才看得出位數（util.show 是全專案統一用的數字顯示格式）。
                show(item["held_qty"]), item["lots"], price_text, item["note"],
            ), tags=(tag,) if tag else ())

        self.order_preview_hint.configure(text="　".join(hints))

    def _resize_order_stock_column(self):
        """
        「股票」欄寬跟著目前 order_rows 的股票名稱重量一次。只在股票清單
        結構真的變了（新增/移除一檔、或切模式/買賣整批清空重選）的時候
        呼叫，不掛在比重/價格輸入框的 trace 上——那些每個按鍵都會觸發
        _recompute_order_preview，要是連欄寬也跟著每個按鍵重算，欄位會在
        使用者打字的時候一直跳動，比原本切到看不全還難用（見 ui_common.
        col_width 的說明）。
        """
        texts = [f"{row['code']} {row['name']}" for row in self.order_rows]
        self.order_preview.column("stock", width=col_width(self.family, texts, minimum=wide(90)))

    def _resize_order_sheet_column(self):
        """
        「帳戶」欄寬跟著這次「重新整理」讀回來的帳戶名單重量一次——只在
        _fill_order_accounts 換了一批名單時呼叫，理由跟 _resize_order_stock_column
        一樣：名單只在讀取的當下換一批，不會因為使用者操作畫面上其他東西
        （勾帳戶、改比重）而變動。
        """
        self.order_preview.column(
            "sheet", width=col_width(self.family, list(self.order_return_rates), minimum=wide(90)))

    # ---------- 查詢委買賣（盤中限定） ----------
    #
    # 「查詢委買賣」按鈕：先幫目前清單裡的股票整批查一次即時委買賣一，讓
    # 執行預覽直接顯示 orders.chase_price 算出來的實際價格，不用等「開始
    # 下單」依序跑到那一筆才臨時查（2026/08/29 使用者要求：出清股票時想在
    # 按下去之前就看到會發生什麼事）。跟 start_order_execution 借同一組
    # self.busy／瀏覽器背景執行緒，理由一樣：這一步也要登入／換 cookie，
    # 不能跟同步分頁或下單依序執行同時搶同一顆瀏覽器。

    def fetch_order_quotes(self):
        """
        觸發背景查詢；結果回來見 _on_order_quotes_fetched。

        報價是公開資料，不因帳戶而不同（跟 order_exec_prices 那份 Excel
        股價「哪個帳戶先讀到就先用哪個」同一種態度）——這裡挑排在執行順序
        第一位的帳戶去登入，純粹是借「已經登入」這件事開 FastQuote 彈出
        視窗，不代表這批報價只給那個帳戶用。
        """
        if self.busy or self.order_quotes_busy:
            return
        if self.order_mode.get() != "intraday":
            return

        codes = sorted({row["code"] for row in self.order_rows})
        if not codes:
            messagebox.showinfo("還沒有股票", "請先加入至少一檔股票。", parent=self.root)
            return

        ordered, _skipped = orders.order_accounts(self._selected_order_accounts())
        if not ordered:
            messagebox.showinfo("還沒有帳戶", "請先勾選至少一個帳戶——查詢委買賣要借用一組帳戶登入。",
                                parent=self.root)
            return
        order_number = self._order_number_for_sheet(ordered[0]["sheet"])
        if order_number is None:
            messagebox.showerror("找不到帳戶", f"{ordered[0]['sheet']} 對不到任何一組帳號。", parent=self.root)
            return

        self._order_quotes_requested = codes
        self.order_quotes_busy = True
        self._set_busy(True, "查詢即時委買賣中…")
        self._update_order_quotes_ui()
        self._ensure_browser_thread()
        self.browser_waiting += 1
        self.browser_cmd_queue.put(
            ("order_quotes", (order_number, self.accounts[order_number - 1], codes)))

    def _order_quotes_job(self, context, store, order_number, account, codes):
        """
        背景執行緒用（只能在 ui_background._browser_worker 裡呼叫）：借這組
        已登入的帳戶開一個 fastquote.FastQuoteStream，一次訂閱這一批股票
        代號，查回目前的委買一／委賣一／成交價。

        跟 _order_fill_job 裡那個「每筆單各自開各自關」的一次性用法是同一招，
        差別只在這裡一次訂閱一整批代號、不是一檔——FastQuoteStream.subscribe
        本來就吃一個代號清單，不需要另外寫批次版本。查不到的代號（逾時、
        不在自選清單…）就不會出現在回傳的字典裡，不是塞一個 None 佔位，
        呼叫端（_on_order_quotes_fetched）自己比對哪些代號漏了。
        """
        page, _, problems = fetch_mod.ensure_logged_in(context, [(order_number, account)], store)[order_number]
        if problems:
            raise RuntimeError("；".join(problems))

        quotes = {}
        stream = fastquote.FastQuoteStream(page)
        try:
            stream.subscribe(codes)
            for code in codes:
                quote = stream.wait_for(code)
                if quote:
                    quotes[code] = quote
        finally:
            stream.close()
        return quotes

    def _on_order_quotes_fetched(self, payload):
        """
        fetch_order_quotes 的背景回話。查到的併進 self.order_quotes（不是
        整份換掉——重複按「查詢委買賣」，這次沒查到的代號還留著上次查到的
        舊值，比整份清空更安全，見下面 missing 那段的說明），再重算一次
        執行預覽讓畫面反映最新算出來的價格。
        """
        self.browser_waiting = max(0, self.browser_waiting - 1)
        self.order_quotes_busy = False
        self._set_busy(False)
        self._update_order_quotes_ui()

        if "error" in payload:
            detail = payload["error"][-1500:]
            hint = payload.get("hint")
            text = f"{hint}\n\n────────────────\n{detail}" if hint else detail
            messagebox.showerror("查詢委買賣失敗", text, parent=self.root)
            return

        quotes = payload["quotes"]
        self.order_quotes.update(quotes)
        missing = [code for code in self._order_quotes_requested if code not in quotes]
        self._recompute_order_preview()
        if missing:
            self._say(f"查詢委買賣：{len(quotes)} 檔查到、{len(missing)} 檔沒查到"
                      f"（{'、'.join(missing)}），這幾檔下單前還是會照原本方式即時查一次。")
        else:
            self._say(f"查詢委買賣：{len(quotes)} 檔都查到了，執行預覽已經是下單會用的價格。")

    def _update_order_quotes_ui(self):
        if self.order_quotes_busy:
            self.order_quotes_button.configure(text="查詢中…", state="disabled")
            return
        intraday = self.order_mode.get() == "intraday"
        self.order_quotes_button.configure(
            text="查詢委買賣", state="normal" if intraday and not self.busy else "disabled")

    # ---------- 依序執行（半自動下單） ----------
    #
    # 跟 order_fill.py 同一個界線：程式只填到「開出委託確認視窗」為止，不按
    # 裡面的「確認」，那一步留給人。這裡多做的是把執行預覽裡一整批委託接起來，
    # 一筆一筆換帳戶（換 cookie，不重登，見 fetch.new_store）自動填單——但兩筆
    # 之間一定要等使用者在瀏覽器裡把上一筆的確認視窗處理掉（按確認或取消）才
    # 能繼續，不能自動往下接：整個瀏覽器只有一組 cookie，要是視窗還開著就換了
    # 下一個帳戶的 cookie，使用者事後才回去按那個視窗的「確認」，送出的會是
    # 「現在瀏覽器帶著的身分」而不是視窗原本對應的那個帳戶——這是真的會送錯
    # 帳戶的委託，不是好看不好看的問題。「下一筆」按鈕只在背景執行緒親眼確認
    # 視窗真的關了（_order_dialog_closed）之後才解鎖，就是為了擋住這條路。
    #
    # 這一輪按下「開始下單」凍結的 self.order_exec_queue 用的是 self.busy
    # （見 ui_background._set_busy）——跟同步分頁共用同一個總開關，是因為
    # 兩邊操作的是同一個瀏覽器 context、同一組 cookie：這一輪還沒跑完時，
    # 「登入」「讀取」「全部登出」都會換手上這組 cookie，一樣會製造上面
    # 那種送錯帳戶的風險，所以整輪期間直接借用同一顆busy鎖把那幾顆按鈕鎖住。

    def _order_number_for_sheet(self, name):
        """分頁名 -> 第幾組帳號。找不到（理論上不會，執行預覽的名字都從 trader_of 長出來）就回 None。"""
        return next((order for order, sheet in self.trader_of.items() if sheet == name), None)

    def start_order_execution(self):
        """
        「開始下單／下一筆」共用同一顆按鈕：queue 是空的就是「開始」（重新算一次
        執行預覽、跳確認視窗、凍結成 queue），queue 還有東西就是「下一筆」，
        直接送出目前這一筆（可能是全新的一筆，也可能是上一筆失敗後的重試，
        見 _on_order_filled）。

        模式（盤前／盤中）、盤中的追價檔數、買賣方向，都在這裡凍結成
        self.order_exec_mode／self.order_exec_ticks／self.order_exec_side，
        之後每一筆都用凍結的值，不是每次都重讀畫面——理由跟凍結 queue 一樣：
        按下「開始下單」之後使用者還是可以去改上面的設定，那些改動只影響
        「下一輪」，不能半路插進正在跑的這一輪。
        """
        if self.order_exec_queue:
            self._dispatch_next_order()
            return

        # 看 _excel_in_use()：勾了「自動更新股價」的話，按下去第一件事就是
        # _prepare_next_round → 用 COM 跑巨集、重讀 Excel（見那條路），不能在
        # 別人正在動同一份活頁簿的時候開始。
        if self._excel_in_use():
            return

        mode = self.order_mode.get()
        side = self.order_side.get()
        multi_round = self.order_multi_round.get()
        auto_price = self.order_auto_price.get()
        ordered, _skipped = orders.order_accounts(self._selected_order_accounts())
        stock_settings = self._order_stock_settings()

        if mode == "intraday":
            ticks = self._order_ticks_setting()
            if ticks is None:
                messagebox.showerror("追價檔數不對", "追價檔數要填 0 以上的整數。", parent=self.root)
                return
            preview = orders.plan_intraday_orders(
                stock_settings, ordered, self.order_holdings, ticks, side,
                prices=self.order_prices, quotes=self.order_quotes)
            queue_rows = orders.executable_intraday_orders(preview)
        else:
            ticks = None
            preview = orders.plan_stock_orders(stock_settings, ordered, self.order_holdings, side)
            queue_rows = orders.executable_orders(preview)

        if not queue_rows:
            reason = ("沒有持股，或比重算出來不到 1 張" if mode == "intraday"
                      else "沒有持股、比重算出來不到 1 張，或者還沒填價格")
            messagebox.showinfo("沒有可以執行的委託",
                f"目前的執行預覽裡，沒有一列是真的可以送出委託的（可能是{reason}）。",
                parent=self.root)
            return

        total_lots = sum(row["lots"] for row in queue_rows)
        auto = self.order_auto_confirm.get()
        side_word = "買進" if side == orders.SIDE_BUY else "賣出"

        if mode == "intraday":
            # 按過「查詢委買賣」的那幾筆 row["price"] 已經是算好的數字（見
            # orders.plan_intraday_orders），開始下單時會直接照用、不再重查
            # （2026/08/29 使用者要求）；沒查過的那幾筆還是老路，下單前才
            # 臨時查一次——三種措辭分開講，不能讓人以為全部都是「現在看到
            # 的數字不會變」或全部都是「還會再變」，兩種情況可能同時存在。
            frozen = sum(1 for row in queue_rows if row["price"] is not None)
            if frozen == len(queue_rows):
                price_note = (f"每一筆的價格都已經用「查詢委買賣」查到的即時報價算好，"
                              f"下單就是直接用執行預覽上看到的數字，不會再重查。\n")
            elif frozen == 0:
                price_note = (f"每一筆的價格以 Excel 讀到的成交價為基準，下單前再追 {ticks} 檔、"
                              f"跟對手方第一檔比價算出來，不是現在看到的數字。\n")
            else:
                price_note = (f"其中 {frozen} 筆已經用「查詢委買賣」查到的即時報價算好，"
                              f"直接用執行預覽上看到的數字；其餘 {len(queue_rows) - frozen} 筆"
                              f"還沒查過，下單前才會即時查一次算出來。\n")
            head = (
                f"即將依序處理 {len(queue_rows)} 筆「{side_word}」委託（共 {total_lots} 張），用 IOC。\n\n"
                f"{price_note}沒成交的部位 IOC 會自動取消，不會掛著。\n"
            )
        else:
            head = f"即將依序處理 {len(queue_rows)} 筆「{side_word}」委託（共 {total_lots} 張），用 ROD-當日有效。\n\n"

        if multi_round:
            if auto_price:
                price_note = "每一輪開始前會先觸發 Excel 的「更新股價」巨集、重讀最新股價。\n"
            elif mode == "intraday":
                price_note = ("沒勾「自動更新股價」：後面幾輪的追價基準價還是會重讀一次 Excel，"
                              "但不會先觸發「更新股價」巨集，數字通常跟這一輪相同。\n")
            else:
                price_note = "盤前的價格不會跟著更新，後面幾輪會沿用你現在填的這個價格。\n"
            head += (
                f"已勾選「多輪直到出清」：這一輪的委託處理完之後，會自動重讀持股，"
                f"還有沒出清的部位就自動接下一輪，最多跑 {ORDER_MULTI_ROUND_CAP} 輪，"
                f"跑滿還沒出清會停下來等你決定。\n{price_note}"
                f"每一筆委託是否要停下來等你確認，規則不受影響。\n\n"
            )

        if auto:
            tail = (
                f"本輪為「自動送出」模式：每一筆委託皆會自動按下委託確認視窗中的"
                f"「確認」，不再等待人工確認，委託將直接送出。\n"
                f"執行節奏不變，仍為逐筆處理——視窗自動關閉後「下一筆」才會轉為"
                f"可按。\n\n"
                f"確定要開始嗎？"
            )
        else:
            tail = (
                f"每一筆都會停在委託確認視窗，不會自動送出，請自己確認或取消。\n"
                f"視窗關閉後「下一筆」才會亮起——換下一筆之前，瀏覽器裡不要留著沒處理的確認視窗。\n\n"
                f"確定要開始嗎？"
            )
        if not ask_confirm(self.root, "開始下單", head + tail, confirm_style="primary"):
            return

        self.order_exec_mode = mode
        self.order_exec_ticks = ticks
        self.order_exec_auto = auto
        self.order_exec_side = side
        self.order_exec_multi_round = multi_round
        self.order_exec_auto_price = auto_price
        self.order_exec_stock_settings = stock_settings
        self.order_exec_accounts = ordered
        # 第 1 輪（沒勾自動更新股價時就是唯一一輪）直接拿 self.order_prices
        # 當起點——那是新增股票／上次「重新整理」讀進來的 Excel 成交價，
        # 不用再另外查一次。勾了自動更新股價的話，這份值一送進
        # _prepare_next_round 馬上就會被剛重讀（含觸發巨集）的結果整份蓋掉
        # （見 _on_order_price_refresh），不是兩份資料混用。
        self.order_exec_prices = dict(self.order_prices)
        # 凍結這一刻查到的即時委買賣（見 order_exec_quotes 開頭的說明：只給
        # 第 1 輪用，第 2 輪以後 _on_order_price_refresh 會清空）。
        self.order_exec_quotes = dict(self.order_quotes)
        self.order_exec_active = True
        self._set_busy(True, "下單：準備第 1 輪…" if multi_round else "下單：準備第 1 筆…")

        if auto_price:
            # 「每一輪開始前都先觸發更新股價，包含第一輪」（規劃文件）——第
            # 一輪不能例外用這裡先算好的 queue_rows 直接送，要跟第 2 輪以後
            # 走同一條路（_prepare_next_round）：先觸發巨集、重讀 Excel，
            # 用讀回來的持股／股價重新組隊列再開始送。round 從 0 開始，
            # _on_order_price_refresh 判斷「還有東西可以送」時會 +1 變成
            # 第 1 輪——跟後面每一輪 +1 的邏輯是同一條，不必另外寫一次。
            self.order_exec_round = 0
            self._prepare_next_round()
        else:
            self.order_exec_queue = queue_rows
            self.order_exec_pos = 0
            self.order_exec_round = 1
            self._dispatch_next_order()

    def _dispatch_next_order(self):
        row = self.order_exec_queue[self.order_exec_pos]
        order_number = self._order_number_for_sheet(row["sheet"])
        if order_number is None:
            messagebox.showerror("找不到帳戶", f"{row['sheet']} 對不到任何一組帳號，這筆沒辦法執行，這一輪停止。")
            self.order_exec_active = False
            self.order_exec_queue = []
            self.order_exec_pos = 0
            self._set_busy(False)
            self._update_order_exec_ui()
            return

        account = self.accounts[order_number - 1]
        self.order_exec_busy = True
        self._update_order_exec_ui()
        self._say(f"下單：第 {self.order_exec_pos + 1}/{len(self.order_exec_queue)} 筆"
                  f"（{row['sheet']} {row['code']}）登入／填單中…")
        self._ensure_browser_thread()
        self.browser_waiting += 1
        self.browser_cmd_queue.put(
            ("order", (order_number, account, row, self.order_exec_mode,
                      self.order_exec_ticks, self.order_exec_auto, self.order_exec_side)))

    def stop_order_execution(self):
        """
        放棄整批作業（可能不只一輪）。瀏覽器裡如果還留著一個沒處理的委託
        確認視窗，程式不會再幫忙追蹤它——那一頁還在，使用者自己回去按
        「確認」或「取消」就好，只是「下一筆」的自動流程到這裡為止。

        order_exec_active 而不是 order_exec_queue 當守門——多輪出清輪與輪
        之間 queue 是空的（正在背景重讀 Excel），這個空檔也要按得下「停止」，
        不能因為 queue 剛好是空的就當作沒什麼在跑。這個空檔按下去沒辦法
        真的中斷背景那條讀 Excel 的執行緒，只能讓它讀完之後的結果被
        _on_order_price_refresh 忽略掉（見那裡的 order_exec_active 檢查）。
        """
        if not self.order_exec_active:
            return
        if self.order_exec_price_busy:
            note = ("\n\n目前正在背景重讀 Excel／觸發「更新股價」巨集，這個動作沒辦法"
                    "中途中斷，會等它跑完，但跑完的結果不會再接下一輪。")
        else:
            note = ""
        if not ask_confirm(
                self.root, "停止下單",
                f"確定要停止整批「多輪直到出清」作業嗎？{note}\n\n"
                f"如果瀏覽器裡還留著一個沒處理的委託確認視窗，"
                f"程式不會再幫你追蹤它，請自己到瀏覽器裡按「確認」或「取消」。"
                if self.order_exec_multi_round else
                "確定要停止這一輪嗎？\n\n如果瀏覽器裡還留著一個沒處理的委託確認視窗，"
                "程式不會再幫你追蹤它，請自己到瀏覽器裡按「確認」或「取消」。",
                confirm_style="primary"):
            return
        self.order_exec_active = False
        self.order_exec_queue = []
        self.order_exec_pos = 0
        self.order_exec_busy = False
        self.order_exec_watching = False
        self.order_exec_last_note = ""
        self._set_busy(False)
        self._update_order_exec_ui()
        self._say("下單：已停止。")

    def _on_order_filled(self, payload):
        """
        背景回話：這一筆的登入／填單有沒有成功。

        半自動：成功不代表這一輪結束——委託確認視窗還開著，要等
        _on_order_dialog_closed 才會往下一筆走（見本節開頭「依序執行」的
        說明）。自動送出：這時候委託已經真的送出去了（見 order_fill.
        confirm_order），視窗也已經關掉，_on_order_dialog_closed 幾乎會
        立刻跟著發生，不用另外處理。

        maybe_submitted（見 order_fill.OrderMaybeSubmitted）是完全不同等級
        的錯誤：「確認」已經按下去、委託多半已經送出去了，只是沒等到結果，
        絕對不能沿用一般失敗那句「按下一筆會重試同一筆」——重試會把同一筆
        委託再送一次，這裡要用不一樣的標題跟文字整個攔下來。
        """
        self.browser_waiting = max(0, self.browser_waiting - 1)
        self.order_exec_busy = False
        row = self.order_exec_queue[self.order_exec_pos]

        if "error" in payload:
            self._update_order_exec_ui()
            detail = payload["error"][-1500:]
            hint = payload.get("hint")
            text = f"{hint}\n\n────────────────\n{detail}" if hint else detail
            if payload.get("maybe_submitted"):
                messagebox.showerror(
                    "委託結果不確定——請先去網站查證",
                    f"{row['sheet']} {row['code']} 這一筆已經按下「確認」，但程式沒辦法"
                    f"確定送出去的結果。\n請先到瀏覽器裡「委託查詢」或「預約查詢」頁自己"
                    f"確認這筆的狀態，確認清楚之前不要按「下一筆」——重試可能會把同一筆"
                    f"委託再送一次。\n\n{text}")
                self._say(f"下單：第 {self.order_exec_pos + 1}/{len(self.order_exec_queue)} 筆"
                          f"結果不確定，先去網站查證，不要按「下一筆」。")
            else:
                messagebox.showerror(
                    "這一筆下單失敗",
                    f"{row['sheet']} {row['code']} 這一筆失敗，這一輪先停在這裡。\n"
                    f"排除問題後按「下一筆」會重試同一筆，或按「停止」放棄這一輪。\n\n{text}")
                self._say(f"下單：第 {self.order_exec_pos + 1}/{len(self.order_exec_queue)} 筆失敗，等你處理。")
            return

        auto_result = payload.get("auto_result")
        self.order_exec_last_note = (f"已自動送出，結果：{auto_result['message']}"
                                     if auto_result else "")
        self.order_exec_watching = True
        self._update_order_exec_ui()
        if auto_result:
            self._say(f"下單：{row['sheet']} {row['code']} 已自動送出，結果：{auto_result['message']}")
        else:
            self._say(f"下單：{row['sheet']} {row['code']} 委託確認視窗已開啟，"
                      f"請到瀏覽器裡確認或取消。視窗關閉後「下一筆」才會亮起來。")

    def _on_order_dialog_closed(self, _payload):
        """背景執行緒親眼確認委託確認視窗真的關了（見 ui_background._browser_worker 的閒置輪詢）。"""
        self.order_exec_watching = False
        self.order_exec_pos += 1
        if self.order_exec_pos >= len(self.order_exec_queue):
            self.order_exec_queue = []
            self.order_exec_pos = 0
            self._order_round_finished()
            return
        self._update_order_exec_ui()

    def _order_round_finished(self):
        """
        這一輪的委託全部處理完了。沒勾「多輪直到出清」就跟原本行為一樣直接
        收尾；勾了的話要重讀 Excel（見 _prepare_next_round）才知道還有沒有
        沒出清的部位，不能在這裡就地判斷完成——判斷完成需要的是「這一輪
        委託真的成交之後」的持股，不是下單前那個舊的 self.order_holdings。
        """
        if not self.order_exec_multi_round:
            self.order_exec_active = False
            self._set_busy(False)
            self._say("下單：這一輪已經跑完。")
            self._update_order_exec_ui()
            return

        if self.order_exec_round >= ORDER_MULTI_ROUND_CAP:
            self.order_exec_active = False
            self._set_busy(False)
            self._update_order_exec_ui()
            self._say(f"下單：已經連續跑了 {ORDER_MULTI_ROUND_CAP} 輪，先停在這裡等你確認。")
            messagebox.showwarning(
                "已到多輪上限",
                f"「多輪直到出清」已經連續跑了 {ORDER_MULTI_ROUND_CAP} 輪，為了安全先停下來，"
                f"不會無限跑下去。\n請自己檢查目前持股，如果還沒出清，可以再按一次"
                f"「開始下單」重新跑一輪。", parent=self.root)
            return

        self._say(f"下單：第 {self.order_exec_round} 輪已跑完，重新讀取持股中…")
        self._prepare_next_round()

    def _prepare_next_round(self):
        """
        背景重讀 Excel（見 _order_price_refresh_worker），依 order_exec_auto_price
        決定要不要先觸發「更新股價」巨集。結果回來由 _on_order_price_refresh
        接手判斷是否已經出清、要不要繼續下一輪。

        第 1 輪也會走這裡（見 start_order_execution）——「每一輪開始前都先
        更新股價，包含第一輪」是規劃文件明講的，第一輪不能因為 queue 已經
        算好了就跳過這一步，兩者共用同一條路，不是各自維護一份類似的邏輯。
        """
        self.order_exec_price_busy = True
        self._apply_busy_state()
        self._update_order_exec_ui()
        sheets = sorted({account["sheet"] for account in self.order_exec_accounts})
        threading.Thread(
            target=self._order_price_refresh_worker,
            args=(self.path, sheets, self.order_exec_auto_price),
            daemon=True,
        ).start()

    def _order_price_refresh_worker(self, path, sheets, run_macro):
        """
        背景執行緒：多輪出清用。run_macro 為真就先觸發使用者既有的「更新
        股價」巨集、等它跑完，再重讀這幾個分頁的持股（E/F 欄）跟股價
        （I 欄）——巨集也好、重讀也好，都跟 refresh_order_data 一樣只在
        COM 層面動，不碰瀏覽器，所以不必透過 browser_cmd_queue，直接開一條
        執行緒做（同 _order_read_worker 的做法）。
        """
        import pythoncom

        pythoncom.CoInitialize()
        excel = workbook = sheet = None
        payload = {}
        try:
            # write=run_macro：巨集要往 I4:I8 寫回股價，開檔如果是唯讀的話
            # 巨集這一步可能直接被 Excel 擋下來或跑到一半出錯。實務上這裡
            # 幾乎一定會走「檔案已經開在使用者的 Excel 裡」那個接上既有
            # 視窗的分支（見 excel_io.open_workbook），write 參數對那個分支
            # 沒有影響，只有在真的還沒開過檔、要另外生一個隱形 Excel 實例
            # 那個少見的情況才有差——但既然要觸發會寫入的巨集，這裡就不要
            # 省這個 True，錯的方向（該可寫卻開成唯讀）比多寫一個參數危險。
            with excel_io.opened(path, run_macro) as (excel, workbook, _attached):
                data, errors = {}, {}
                with excel_io.keep_active_sheet(workbook):
                    for name in sheets:
                        sheet, error = excel_io.find_sheet(workbook, name)
                        if sheet is None:
                            errors[name] = error
                            continue
                        # 一頁一次，理由同 _order_read_worker。
                        if run_macro:
                            excel_io.run_update_price_macro(excel, sheet)
                        data[name] = excel_io.read_sheet(sheet)
                payload = {"sheets": data, "errors": errors}
        except Exception as exc:
            payload = {"error": str(exc)}
        finally:
            sheet = excel = workbook = None
            pythoncom.CoUninitialize()
        self.queue.put(("order_price_refresh", payload))

    def _on_order_price_refresh(self, payload):
        """
        _prepare_next_round 的背景讀取回話。重新組一次隊列，還有東西就是
        還沒出清，凍結成下一輪的 queue 繼續跑；空了就是這一批全部出清了。
        """
        self.order_exec_price_busy = False
        self._apply_busy_state()

        if not self.order_exec_active:
            # 讀到一半使用者按了「停止」——這個結果已經跟不上了，不要
            # 回頭把畫面/忙碌狀態又動一次（stop_order_execution 已經處理過）。
            return

        if "error" in payload:
            self.order_exec_active = False
            self._set_busy(False)
            self._update_order_exec_ui()
            messagebox.showerror(
                "重讀持股失敗",
                f"第 {self.order_exec_round} 輪跑完後想重新讀取持股，但失敗了，"
                f"「多輪直到出清」先停在這裡：\n\n{payload['error']}", parent=self.root)
            self._say("下單：重讀持股失敗，多輪出清已停止。")
            return

        errors = payload["errors"]
        if errors:
            # 部分帳戶讀不到就整批停下來，不要拿舊資料繼續算下一輪——這幾位
            # 沒讀到的話沒辦法判斷他們到底出清了沒，硬跑下去可能漏單也可能
            # 對著已經出清的部位重複下單。
            self.order_exec_active = False
            self._set_busy(False)
            self._update_order_exec_ui()
            messagebox.showerror(
                "重讀持股失敗",
                f"這幾個帳戶讀不到，「多輪直到出清」先停在這裡：\n"
                f"{'、'.join(f'{name}：{msg}' for name, msg in errors.items())}", parent=self.root)
            self._say("下單：部分帳戶讀不到持股，多輪出清已停止。")
            return

        for name, data in payload["sheets"].items():
            for row in data["rows"]:
                self.order_holdings[(name, row["code"])] = row["qty"]
            # 這個帳戶原本有、這次沒讀到的股票代表已經出清歸零（read_sheet
            # 空白列直接跳過，不會出現在 rows 裡）——不是「讀不到所以不管」，
            # 是這一列真的空了，_order_stock_settings 裡的每一檔都要能反映
            # 這件事，所以先把這個帳戶名下、這一輪處理過的股票代碼清成 0，
            # 再用這次讀到的蓋回去。
            codes_this_round = {s["code"] for s in self.order_exec_stock_settings}
            for code in codes_this_round:
                self.order_holdings.setdefault((name, code), 0)
            for code in codes_this_round - {r["code"] for r in data["rows"]}:
                self.order_holdings[(name, code)] = 0

        # I 欄跟 D 欄同一列對應同一檔股票（見 excel_io.py 開頭說明），這裡把
        # 讀到的股價彙整成 代號->價格，哪個帳戶先讀到就先用哪個——同一檔
        # 股票的市價理論上不會因為帳戶不同而不同，這裡不刻意比對多個帳戶
        # 讀到的數字是否一致，跟現有 fetch 那幾支查詢的態度不同，是因為這裡
        # 本來就只是「參考同一份 Excel 既有機制」，不是本程式自己認定的
        # 權威資料源。整份換掉不是累加——上一輪的舊價格這一輪不該還留著
        # （見 order_exec_prices 開頭的說明）。不管 order_exec_auto_price 這
        # 一輪有沒有先觸發「更新股價」巨集都覆蓋——這裡的 sheets 資料本來就
        # 已經讀進來了，沒理由只挑巨集有跑的那幾輪才拿來用（2026/08/29 使用
        # 者確認：盤中一律用 Excel 成交價，這個開關只決定重讀前要不要先觸發
        # 巨集，不再決定 pricenow 走 Excel 還是現查網頁）。
        # _order_fill_job 盤中算追價時會拿這份當 pricenow，跟 FastQuote 現查
        # 委買賣一一起餵給 chase_price 當邊界。
        price_by_code = {}
        for data in payload["sheets"].values():
            for row in data["rows"]:
                if row["code"] not in price_by_code and row["price"] is not None:
                    price_by_code[row["code"]] = row["price"]
        self.order_exec_prices = price_by_code
        # 第 1 輪凍結的即時委買賣只給第 1 輪用（見 order_exec_quotes 開頭的
        # 說明）——這裡是第 2 輪以後才會走到的路，清空逼這一輪的每一列都
        # 退回「下單前才查」，不把已經過時的報價繼續當最新的用。
        self.order_exec_quotes = {}

        side = self.order_exec_side
        if self.order_exec_mode == "intraday":
            preview = orders.plan_intraday_orders(
                self.order_exec_stock_settings, self.order_exec_accounts,
                self.order_holdings, self.order_exec_ticks, side,
                prices=self.order_exec_prices, quotes=self.order_exec_quotes)
            queue_rows = orders.executable_intraday_orders(preview)
        else:
            preview = orders.plan_stock_orders(
                self.order_exec_stock_settings, self.order_exec_accounts,
                self.order_holdings, side)
            queue_rows = orders.executable_orders(preview)

        self._render_order_preview(preview, [])

        if not queue_rows:
            self.order_exec_active = False
            self._set_busy(False)
            self._update_order_exec_ui()
            # queue 空了不一定是「真的出清歸零」——比重四捨五入到 1 張，
            # 剩不到半張的零股永遠湊不出下一張整張委託，會一直卡在這裡，
            # 不能因為 queue 空了就跟人講「已經出清」，那不是真的（見
            # docs 規劃裡「出清股票(零股)」是另一個還沒做的功能，這裡的
            # 多輪只處理整張）。
            leftover = sum(item["held_qty"] for item in preview if item["held_qty"] > 0)
            if self.order_exec_round == 0:
                # 第一輪重讀完就發現沒東西可送——可能是持股在按下「開始
                # 下單」之後、巨集跑完之前這個空檔剛好變了，一輪都還沒真的
                # 跑，跟「跑了幾輪之後出清」是不同的事，訊息不能講「跑了
                # 0 輪」，那不是人話。
                self._say("下單：重讀 Excel 之後，沒有需要處理的委託了（可能持股剛好在這個空檔變了）。")
            elif leftover > 0:
                self._say(f"下單：跑了 {self.order_exec_round} 輪，整張的部分已經出清，"
                          f"但還剩下不到 1 張的零股（比重換算不出下一張整張委託），"
                          f"多輪出清到這裡為止，零股要另外處理。")
            else:
                self._say(f"下單：跑了 {self.order_exec_round} 輪，已經全部出清。")
            return

        self.order_exec_round += 1
        self.order_exec_queue = queue_rows
        self.order_exec_pos = 0
        self._say(f"下單：開始第 {self.order_exec_round} 輪…")
        self._update_order_exec_ui()
        self._dispatch_next_order()

    def _update_order_exec_ui(self):
        if self.order_exec_price_busy:
            # 多輪出清輪與輪之間：queue 是空的，但整批作業還在跑（見
            # order_exec_active 的說明），「停止」要維持可以按。
            self.order_exec_button.configure(text="處理中…", state="disabled")
            self.order_exec_stop_button.configure(state="normal")
            self.order_exec_status.configure(
                text=f"第 {self.order_exec_round} 輪已跑完，重新讀取持股中…")
            return

        total = len(self.order_exec_queue)
        if total == 0:
            self.order_exec_button.configure(text="開始下單（依序執行）",
                                             state="disabled" if self.busy else "normal")
            self.order_exec_stop_button.configure(state="disabled")
            self.order_exec_status.configure(text="")
            return

        self.order_exec_stop_button.configure(state="normal")
        pos = self.order_exec_pos
        row = self.order_exec_queue[pos]
        label = f"{row['sheet']}　{row['code']} {row['name']}"
        # 多輪出清才標「第 N 輪」——單輪模式一直都只有一輪，標了反而多餘。
        prefix = f"第 {self.order_exec_round} 輪 " if self.order_exec_multi_round else ""
        if self.order_exec_watching:
            self.order_exec_button.configure(text="下一筆", state="disabled")
            note = self.order_exec_last_note or (
                "委託確認視窗已開啟，去瀏覽器確認或取消，視窗關閉後才能按「下一筆」。")
            self.order_exec_status.configure(text=f"{prefix}({pos + 1}/{total}) {label}：{note}")
        elif self.order_exec_busy:
            self.order_exec_button.configure(text="處理中…", state="disabled")
            self.order_exec_status.configure(text=f"{prefix}({pos + 1}/{total}) {label}：登入／填單中…")
        else:
            # 不在等視窗、也不在忙，卡在這裡代表上一筆失敗了（剛開始執行的當下
            # 一定馬上進 order_exec_busy，不會停留在這個分支，見 start_order_execution）。
            self.order_exec_button.configure(text=f"下一筆（{pos + 1}/{total}）", state="normal")
            self.order_exec_status.configure(text=f"{prefix}({pos + 1}/{total}) {label}：上一筆失敗，可以重試或停止。")

    def _order_fill_job(self, context, store, order_number, account, row, mode, ticks, auto, side):
        """
        背景執行緒用（只能在 ui_background._browser_worker 裡呼叫）：登入或換
        cookie 到 row 對應的帳戶、開下單頁、選股票、填單、開出委託確認視窗；
        auto 為真的話再多按下視窗裡的「確認」，真的把委託送出去（見
        order_fill.confirm_order）。

        回傳 (page, extra)：page 只能留在呼叫端（_browser_worker）手上、記著
        要盯哪一頁的確認視窗，絕對不能塞進 self.queue 送回主執行緒——
        Playwright 的物件被綁死在建立它的那個執行緒，主執行緒去碰會直接壞掉
        （見 ui_background.py 開頭的執行緒說明）。extra 是給主執行緒看的純
        資料，auto 送出成功時帶 "auto_result"，半自動或還沒送出前失敗時是
        空字典。

        半自動（auto=False）不按確認視窗裡的「確認」，理由跟 order_fill.
        fill_order 一樣：那一步留給人。

        mode=="intraday" 時 row["price"] 通常是 None（見 orders.plan_intraday_orders），
        要在這裡算出真正要送出的價格——除非按過「查詢委買賣」把這一檔的即時
        報價整批先查回來過（見 fetch_order_quotes），那樣 row["price"] 在
        建執行清單那一刻就已經是算好的數字，這裡直接照用，不再重查一次
        （2026/08/29 使用者要求：按下「開始下單」就是執行預覽上看到的價位，
        不要下單前又變成另一個數字）。

        pricenow（chase_price 的基準價）一律來自 self.order_exec_prices——
        Excel I 欄讀回來的成交價，不現查網頁（2026/08/29 使用者確認：沒必要
        等下單前才另外查一次）。第 1 輪這份值是開始下單那一刻拿 self.
        order_prices 當快照（見 start_order_execution），也就是上一次
        「重新整理」讀到的數字——盤中模式的「重新整理」會先觸發「更新股價」
        巨集才讀（見 refresh_order_data），所以只要開始下單前有按過一次
        「重新整理」，這個基準價就是新的，不是放到過期的舊資料。
        order_exec_auto_price 這個開關現在只決定多輪出清時，輪與輪之間重讀
        Excel 前要不要先觸發巨集，不再決定 pricenow 的資料來源。
        """
        page, _, problems = fetch_mod.ensure_logged_in(
            context, [(order_number, account)], store)[order_number]
        if problems:
            raise RuntimeError("；".join(problems))

        price = row["price"]
        if mode == "intraday" and price is None:
            pricenow = self.order_exec_prices.get(row["code"])
            if pricenow is None:
                raise RuntimeError(
                    f"沒有讀到 {row['code']} 的股價（Excel I 欄），這一筆沒辦法算追價。"
                    f"請先按「重新整理」讓 Excel 更新股價。")

            # 對手方第一檔：借這個已登入的 page 開一個 FastQuote 彈出視窗，
            # 只為了這一檔股票訂閱、等一下、拿到就關掉（見 fastquote.py
            # 「另開分頁不是一律不行」）——刻意保持簡單，不常駐訂閱、不接
            # ui_background.py 的即時表格，每筆單各自開各自關。查哪一檔跟
            # side 是反的：買方向查委賣一（ask）、賣方向查委買一（bid），見
            # orders.chase_price 的說明。收不到（逾時、或 WebSocket 這輪
            # 剛好不穩）就讓 best_opposite 維持 None，chase_price 自己會退回
            # 邊界價，不擋單。
            best_opposite = None
            stream = fastquote.FastQuoteStream(page)
            try:
                if stream.subscribe([row["code"]]):
                    quote = stream.wait_for(row["code"])
                    if quote:
                        best_opposite = quote["ask"] if side == orders.SIDE_BUY else quote["bid"]
            finally:
                stream.close()

            price = orders.chase_price(pricenow, ticks, side, best_opposite)

        # 委託別跟著模式走，不是兩邊共用同一個值（見 orders.BS_FLAG_PRE 的
        # 說明）：盤前開盤前還沒有連續交易，只能用 ROD；盤中規劃文件明講
        # 用 IOC。2026/08/28 使用者更正過，之前這裡兩種模式都寫死 IOC 是錯的。
        bs_flag = orders.BS_FLAG_INTRADAY if mode == "intraday" else orders.BS_FLAG_PRE

        page.goto(order_fill.ORDER_ENTRY_PAGE, wait_until="domcontentloaded")
        order_fill.open_order_form(page)
        order_fill.select_stock(page, row["code"])
        order_fill.fill_order(page, side=side, qty=row["lots"], price=price, bs_flag=bs_flag)

        if not auto:
            return page, {}

        ok, message = order_fill.confirm_order(page)
        return page, {"auto_result": {"ok": ok, "message": message}}

    def _order_dialog_closed(self, page):
        """
        委託確認視窗還開著沒？只能在背景瀏覽器執行緒裡呼叫（同上，Playwright
        物件的執行緒限制）。

        頁面／瀏覽器已經不在了（使用者自己把分頁或整個瀏覽器關掉）就當作「已經
        關閉」——不能讓一個已經摸不到的頁面卡住整個「下一筆」流程，最壞的後果
        頂多是使用者接下來自己再重登一次，比程式在這裡卡死好處理。
        """
        try:
            return not page.locator(".layui-layer-title", has_text="委託確認").first.is_visible()
        except PlaywrightError:
            return True
