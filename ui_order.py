"""
下單分頁「按下開始下單之前」的那一半：收使用者的輸入（選股票、填比重、勾
帳戶）、讀 Excel（持股、B17 報酬率、成交價）、查即時委買賣一，算出執行預覽。

按下去之後的另一半在 ui_order_exec.py——那裡從凍結這一輪的設定開始，一筆
一筆送出、跑完一輪再決定要不要接下一輪。界線是「這一輪要送什麼」（這裡）
跟「怎麼把它送出去」（那裡），見那個檔案開頭的說明。

「盤前」（股票／比重／價格設定）跟「盤中」（股票／比重／追價檔數設定，價格
用 Excel 成交價＋下單前查對手方第一檔算出來的）共用同一套選帳戶、算執行
預覽、依序執行的機制，只有「股票設定要填什麼」「怎麼組出執行清單」不一樣
（見 `_on_order_mode_changed`／ui_order_exec 的 `start_order_execution`）。

比重→張數、帳戶依 B17 報酬率排序、組出預覽清單、追價檔數換算價格，全部是
orders.py 的純函式，這裡只負責收輸入、讀 Excel（含成交價，盤中新增股票／
讀取持股時順便觸發「更新股價」巨集）、查即時對手方第一檔、把結果畫出來。
"""

import threading
import tkinter as tk

import ttkbootstrap as ttk
from playwright.sync_api import Error as PlaywrightError

import excel_io
import fastquote
import fetch as fetch_mod
import order_fill
import orders
from ui_common import (
    FONT_SIZE, ORDER_STOCK_ROW_H, ORDER_STOCK_ROWS_SHOWN, PRICE_PENDING_TEXT,
    ask_confirm, col_width, show_error, show_info, wide,
)
from util import show


# 「指定股票」每一列的欄位起點（照 10 級字的像素，實際用 wide() 換算）。
# 每一列各自是一個 Frame，但欄的 minsize 全部用這同一組，所以三檔的「移除」
# 「比重」「價格」會上下對齊成一張小表——不是靠股票名稱剛好一樣長。
#
# 股票那一欄 110 夠放「2330 台積電」這種四碼＋三到五個中文；真的更長也不會被
# 切掉（minsize 是下限不是上限），只是那一列的後面幾欄會往右推、跟別列對不齊。
ORDER_STOCK_COL_W = (26, 110, 60, 130)


class UiOrderMixin:
    # ---------- 下單分頁：盤前模式 ----------

    def _order_init_state(self):
        """SyncApp.__init__ 呼叫一次。"""
        self.order_rows = []              # 這一輪加進來的股票設定列（見 add_order_stock）
        self.order_holdings = {}          # (分頁名, 股票代號) -> 股數，按「讀取持股」才會更新
        self.order_names = {}             # 股票代號 -> 名稱，畫面顯示用
        self.order_prices = {}            # 股票代號 -> Excel I 欄讀回來的股價；盤中模式這份就是
                                           # chase_price 的 pricenow 來源（見 start_order_execution），
                                           # 不只是畫面顯示用（跟 order_names 平行）
        self.order_return_rates = {}      # 分頁名 -> B17 報酬率或 None（讀不到）
        # (分頁名, 股票代號) -> {"name", "qty"(股數，正買負賣), "price"}：買賣股票作業
        # 用的下單試算 M14:N18（見 excel_io.read_order_plan）。**每個帳戶各有一份**
        # ——它是那一頁自己的試算結果，不像股價那樣全帳戶共用一個值，所以 key 要
        # 帶分頁名，跟 order_holdings 同一種形狀，不是 order_prices 那種。
        self.order_plans = {}
        # 「執行帳戶」表（左欄的兩欄 Treeview，一列一位）：一次一位，沒有
        # 「全部帳戶」（2026/09/01 使用者決定，見 docs/介面規劃.md 9.3 第 5 點）。
        # order_account 存的是**分頁名**，不是畫面上那一列的文字——那一列還帶著
        # 順位號碼與報酬率，兩者都會隨著補讀 B17 而改變，拿它當識別的話，每重畫
        # 一次清單就對不回同一個人。
        self.order_account = tk.StringVar()
        self.order_account_rank = {}      # 分頁名 -> 報酬率由低到高的順位（讀不到是「－」）
        # 「正在去 Excel 讀帳戶清單」（見 refresh_order_accounts）。那一趟只讀
        # B17 與 D4:D8，不跑巨集，但動的是同一份活頁簿，所以照樣要算進
        # _excel_in_use()。
        self.order_rates_busy = False
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

        # 作業（見 docs/介面規劃.md 9.2／9.3）：買賣股票／出清股票／全持股交易
        # 三選一，是這個分頁最上層的選擇。第二列跟著整列換成「那個作業自己的
        # 設定」，右半邊三個作業共用同一份 widget。
        self.order_job = tk.StringVar(value=orders.JOB_TRADE)
        self._order_job_last = orders.JOB_TRADE

        # 單位：整張／零股。買賣股票與出清股票兩張第二列各畫一組 Radiobutton，
        # 但綁的是這同一個變數（Tk 會自己讓兩組保持一致，不必手動同步）。
        self.order_unit = tk.StringVar(value=orders.UNIT_LOT)

        # 全持股交易的兩個追價檔數（整張、零股各一個，見 9.3 的表）。行為還沒
        # 接上，先存著讓第二列畫得出來、看得到版面。
        self.order_full_lot_ticks = tk.StringVar(value="2")
        self.order_full_odd_ticks = tk.StringVar(value="3")

        # 買／賣方向：9.3 第 3 點定案把「賣／買」那組單選鈕整個拿掉——三個作業
        # 沒有一個需要人選方向（出清永遠是賣；買賣股票由 M14:M18 的正負逐檔
        # 決定；全持股交易也是算出來的）。方向只該出現在執行預覽的「買賣」欄
        # 跟紅綠底色。這個變數留著是因為 orders.py 那幾支 plan_* 還是收一個
        # side 參數——差別在它現在由作業決定，不是由人選。
        self.order_side = tk.StringVar(value=orders.SIDE_SELL)

        # 自動送出開關：關（預設）＝半自動，跟原本一樣停在委託確認視窗給人看、
        # 給人按；開＝程式自己按下確認視窗裡的「確認」，委託真的送出去，不會
        # 停下來等人。2026/08/28 使用者要求加的，AskUserQuestion 確認過節奏
        # 不變——不管開不開，還是「下一筆」按一次只處理一筆，差別只在「這
        # 一筆處理完」是靠人看過按確認，還是程式自己按。
        self.order_auto_confirm = tk.BooleanVar(value=False)
        self._order_auto_last = False

        # 「開始下單」之後那一整套凍結狀態（queue、這一輪的模式／追價檔數／
        # 買賣方向、多輪的第幾輪…）在 ui_order_exec.py，跟這裡「畫面上現在填了
        # 什麼」的狀態刻意分開——見那個檔案開頭的說明。
        self._order_exec_init_state()

    def refresh_order_data(self):
        """
        重新讀 Excel：把**選中那一位**的持股（E/F 欄）、B17、股價（I 欄）讀一次，
        買賣股票作業再多讀下單試算 M14:N18。

        2026/09/01 之前這裡讀的是「每個已知名字的帳戶」，一次 20 個分頁——
        使用者的操作本來就是一次處理一位（選帳戶 →「更新→自動計算」→ 加股票），
        另外 19 個分頁讀回來的數字這一輪一個都用不到。盤中模式更貴：那 19 個
        分頁每一頁都要各跑一次「更新股價」巨集，一頁 5 檔就是 5 次 Yahoo HTTP
        （todo.txt C2 記的那 100 次就是這麼來的）。

        盤中模式讀之前先觸發一次「更新股價」巨集再讀（2026/08/29 使用者確認：
        盤中追價用的成交價來自這裡讀到的 order_prices，新增股票／讀取持股與
        報酬率這一步就要盡量拿到新的價格，不能留著上次殘留的舊數字）；盤前
        模式的價格是人手動填的，不需要 Excel 股價，維持原本只讀不觸發巨集。

        帳戶名單只能從 self.trader_of 來——那是「登入過才知道名字」的既有
        規則（見 ui.py），還沒登入過的帳戶這裡也看不到，跟更新分頁的範圍
        選單是同一個限制，不是這裡另外加的。
        """
        # 看 _excel_in_use() 而不是只看 order_busy：更新分頁的寫入、「新增」股票
        # 附帶的股價重讀、多輪之間的重讀，動的都是同一份活頁簿（見那個述詞）。
        if self._excel_in_use() or not self._require_excel():
            return
        sheet = self._order_sheet()
        if sheet is None:
            show_info(self.root, "還沒選帳戶", self._order_no_account_text())
            return
        names = [sheet]

        run_macro = self._order_intraday()
        # 買賣股票的張數與價格來自下單試算 M14:N18，那是另外幾格，只有這個作業
        # 要讀——其他作業讀它只是多 10 格 COM 往返。
        read_plan = self.order_job.get() == orders.JOB_TRADE
        self.order_busy = True
        self._apply_busy_state()
        self.order_status.configure(text="更新股價、讀取中…" if run_macro else "讀取中…")
        threading.Thread(target=self._order_read_worker,
                         args=(self.path, names, run_macro, read_plan), daemon=True).start()

    def _order_read_worker(self, path, names, run_macro, read_plan=False):
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
                            excel_io.run_update_price_macro(
                                excel, sheet, on_stuck=self._macro_stuck_notifier("更新股價", name))
                        data = excel_io.read_sheet(sheet)
                        data["return_rate"] = excel_io.read_return_rate(sheet)
                        if read_plan:
                            data["plan"] = excel_io.read_order_plan(sheet)
                        sheets[name] = data
                # 巨集寫過 I4:I8 就要存檔，不然沒接上使用者既有視窗時
                # close_workbook 會 Close(False) 把這次更新的股價丟掉。
                if run_macro:
                    workbook.Save()
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
            show_error(self.root, "讀取失敗", payload["error"])
            return

        # 持股／股票名稱／股價／試算整份換掉：這一趟讀的就是選中那一位的全部，
        # 留著上一位的只會讓「這一位到底有沒有這檔」變成看運氣（換帳戶那一刻
        # 其實已經清過一次，見 _on_order_account_changed，這裡是第二道）。
        #
        # order_return_rates 例外，不清：它是**所有**帳戶的（那份清單本身，見
        # refresh_order_accounts），清掉的話讀一位就把左邊其他 19 位整個弄丟。
        # 這一位的 B17 順手更新（下面那一行）——這一趟本來就讀到了，不用再跑
        # 一次帳戶清單那條路。
        self.order_holdings, self.order_names, self.order_prices = {}, {}, {}
        self.order_plans = {}
        for name, data in payload["sheets"].items():
            self.order_return_rates[name] = data["return_rate"]
            for code, plan in (data.get("plan") or {}).items():
                self.order_plans[(name, code)] = plan
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

        # 一次只有一位，所以名字寫得出來（以前是「已讀取 N 個帳戶」）。讀不到
        # 那一位的時候整句換掉，不是把名字的位置填一句「沒有分頁」——那會變成
        # 「已讀取 沒有分頁 的持股」這種讀起來卡住的句子。
        errors = payload["errors"]
        if payload["sheets"]:
            done = f"已讀取 {'、'.join(payload['sheets'])} 的持股與報酬率。"
        else:
            done = "沒有讀到任何分頁。"
        note = f"　（讀不到：{'、'.join(errors)}）" if errors else ""
        self.order_status.configure(text=done + note)
        self._recompute_order_preview()

    # ---------- 執行帳戶（一次一位） ----------

    def refresh_order_accounts(self):
        """
        去 Excel 把「這份表裡有誰、各自的今年報酬率」讀回來，重畫「執行帳戶」。

        **開檔就有，不必等登入**（2026/09/01 使用者要求）：一份持股管理表的分頁
        本來就是一位交易人一頁，開檔當下就知道有誰了（見
        excel_io.list_account_sheets）。登入只跟「送得出委託」有關——真的按下
        「開始買賣」時才需要那個人的 cookie，那時對不到帳號會擋下來
        （見 _order_number_for_sheet 的呼叫端）。

        呼叫點兩個，都是「Excel 那一頭可能變了」的時刻：Excel 接上的那一刻
        （ui_background._set_excel_open）與切到下單分頁（ui._on_tab_changed）。
        每次都真的重讀一遍，不記「掃過了沒」——它只讀 B17 跟 D4:D8，一位最多
        六格 COM 往返，20 位也只是幾十毫秒（真正貴的是巨集那 5 次 Yahoo HTTP），
        換來的是「人在 Excel 裡加了一頁、改了報酬率，切個分頁回來就對了」。
        """
        if self._excel_in_use() or not self.excel_open or self.path is None:
            self._fill_order_accounts()
            return
        self.order_rates_busy = True
        self._apply_busy_state()
        threading.Thread(target=self._order_rates_worker,
                         args=(self.path,), daemon=True).start()

    def _order_rates_worker(self, path):
        """背景執行緒：列出交易人分頁與 B17（見 refresh_order_accounts）。"""
        import pythoncom

        pythoncom.CoInitialize()
        excel = workbook = None
        payload = {}
        try:
            with excel_io.opened(path, False) as (excel, workbook, _attached):
                payload = {"accounts": excel_io.list_account_sheets(workbook)}
        except Exception as exc:
            payload = {"error": str(exc)}
        finally:
            excel = workbook = None
            pythoncom.CoUninitialize()
        self.queue.put(("order_rates", payload))

    def _on_order_rates(self, payload):
        """
        帳戶清單回來了。**讀不到不跳視窗**：這一趟是背景自己去的，人沒有按任何
        按鈕，為了它彈視窗等於自己去做事再自己來抱怨。清單空著、或那一位寫著
        「讀不到」，看得到就夠了。

        整份換掉不是逐位更新：這就是「這份活頁簿現在有誰」的答案，人在 Excel 裡
        刪掉一頁的話，逐位更新會讓那一位永遠留在畫面上。
        """
        self.order_rates_busy = False
        self._apply_busy_state()
        if "error" in payload:
            return
        self.order_return_rates = {name: rate for name, rate in payload["accounts"]}
        self._fill_order_accounts()

    def _order_locked(self):
        """
        現在能不能換帳戶。有任何一條路在動那份活頁簿（_excel_in_use）、或這一輪
        委託還沒跑完（order_exec_queue）都不行——理由見 _order_excel_buttons 與
        _on_order_account_changed。
        """
        return self._excel_in_use() or bool(self.order_exec_queue)

    def _order_known_sheets(self):
        """
        清單上會出現哪幾位：**上一次讀 Excel 讀到的那些分頁**（見
        refresh_order_accounts）。2026/09/01 之前是「登入過、知道名字的那些」
        （self.trader_of），改成看 Excel 是因為要開檔就有，不必等登入。
        """
        return list(self.order_return_rates)

    def _fill_order_accounts(self):
        """
        重畫左欄的「執行帳戶」表：整份清掉重建，選中的那一位留著。

        排序用 orders.order_accounts——那支本來是排「一次跑多個帳戶時誰先執行」，
        2026/09/01 改成一次一位之後，它排的是**清單順序**，也就是建議的處理
        順序（規格「報酬率低的先執行」，見 docs/介面規劃.md 9.3 第 5 點）。
        「帳戶」欄名字前面那個號碼就是它給的 order。

        B17 讀不到的那幾位排在最後、報酬率欄寫「讀不到」，但**照樣選得到**：
        報酬率在這裡只決定排序，不決定能不能下單，而現在一次只有一位，根本
        沒有順序可言。（一次跑多位的年代它是硬性條件——見 orders.order_accounts
        的說明：讀不到就當最低硬排進去，等於用猜的決定誰先執行。）

        每一列的 iid 就是分頁名，所以重建之後把選取設回去只要 selection_set(名字)
        ——那一列的文字（號碼、百分比）每次補讀 B17 都可能變，拿文字當識別的話，
        選好的帳戶會在每次重畫時被清掉一次。

        這裡的 delete／insert／selection_set 全都會讓 Tk 送出 `<<TreeviewSelect>>`，
        而 _on_order_account_changed 的工作是「換人了，把上一位的股票清單、持股、
        試算全部清掉」——擋這幾次假換人的機制在那一支，用的是「比對選取與
        order_account」而不是這裡設個旗標，理由見那邊（那些事件是排進佇列的，
        這一支返回之後才送到，旗標早就收掉了）。
        """
        ordered, skipped = orders.order_accounts(
            [{"sheet": name, "return_rate": self.order_return_rates.get(name)}
             for name in self._order_known_sheets()])

        rows = []
        for account in ordered:
            # B17 存的是小數（0.185222... 代表 18.5%），畫面要顯示的是百分比，
            # 這裡要乘 100——漏了這一步會把 18.5% 顯示成 0.2%，跟現金查詢
            # Amount 除以 100 是同一種「單位不對但不會報錯」的坑（CLAUDE.md）。
            rows.append((account["sheet"], account["order"],
                         f"{account['order']}　{account['sheet']}",
                         f"{account['return_rate'] * 100:.1f}%"))
        for account in skipped:
            # 沒有號碼可給——排序這件事對他不成立，用「－」佔位，不要給一個
            # 看起來像順位的數字（執行預覽的「順序」欄顯示的也是這個值）。
            rows.append((account["sheet"], "－", f"－　{account['sheet']}", "讀不到"))

        keep = self.order_account.get()
        self.order_accounts.delete(*self.order_accounts.get_children())
        self.order_account_rank = {}
        for sheet, rank, name_text, rate_text in rows:
            self.order_account_rank[sheet] = rank
            self.order_accounts.insert("", "end", iid=sheet, text=name_text,
                                       values=(rate_text,))
        # 那個人整個從名單上消失了（換 Excel 檔、還沒登入到他）就把選擇清掉，
        # 不留一個對不到任何一列的名字——畫面會一路退回「請先選一位」，不會
        # 拿著一個查不到的人繼續算。
        if keep in self.order_account_rank:
            self.order_accounts.selection_set(keep)
            self.order_accounts.see(keep)
        else:
            self.order_account.set("")
        self._resize_order_sheet_column()
        self._order_excel_buttons()

    def _order_sheet(self):
        """現在選的是哪一位（分頁名），沒選就是 None。"""
        name = self.order_account.get()
        return name if name in self.order_account_rank else None

    def _order_no_account_text(self):
        """
        沒選帳戶時要說的話。清單根本是空的（Excel 還沒開）跟只是還沒點一位，
        要人做的事不一樣。

        「還沒登入」不在這裡講：帳戶清單來自 Excel 分頁，跟登入無關（見
        refresh_order_accounts）；沒登入的話會在真的要送委託那一刻才擋下來。
        """
        if not self._order_known_sheets():
            return "還沒讀到任何帳戶——請先按左上角的「開啟EXCEL」。"
        return "請先在左邊的「執行帳戶」點一位。"

    def _on_order_account_changed(self):
        """
        左邊那張表的選取變了。**「選取變了」不等於「使用者換人了」**，所以這一支
        開頭那兩道判斷是整個機制的重點，不是防禦性冗詞：

        1. `picked == order_account` 就直接回頭。Tk 的 `<<TreeviewSelect>>` 是
           **排進事件佇列、之後才送達**的（不是同步呼叫），而 _fill_order_accounts
           的 delete／insert／selection_set 每一步都會送一個。也就是說「補讀完
           報酬率重畫清單」跟「使用者點了另一個人」長得一模一樣，只有比對最後的
           選取跟 order_account 分得出來。分不出來的代價是每次重畫清單都會把人
           剛加好的股票清單清掉——不報錯，只是東西不見了。
           （用旗標擋不住：這一支返回之後那些事件才送到，旗標早就收掉了。改用
           radiobutton 就沒有這個問題，那是為了跟更新分頁一致付的代價。）
        2. 鎖住的時候把選取扳回去。ttk 的 disabled 擋得住鍵盤滑鼠，擋不住已經
           在路上的那一次選取，而換人會把背景那趟正在讀的資料對應到的人換掉。
           扳回去自己也會再送一次事件，那一次會被第 1 道判斷擋掉。

        真的換人了，股票清單、執行預覽、查回來的報價全部清掉重來——跟切作業
        （_on_order_job_changed）同一條規矩，理由也一樣而且更硬：畫面上留著的
        持股、試算股數、價格全是**上一位**的，只有帳戶那一格換了名字。這種錯
        不會報錯，只會用 A 的數字掛在 B 的帳上（CLAUDE.md 講 _revisit 那四道
        身分核對時擔心的就是同一件事，只是那邊是機器搞錯，這裡是畫面騙人）。
        """
        picked = (self.order_accounts.selection() or ("",))[0]
        if picked == self.order_account.get():
            return
        if self._order_locked():
            keep = self.order_account.get()
            if keep in self.order_account_rank:
                self.order_accounts.selection_set(keep)
            else:
                self.order_accounts.selection_remove(*self.order_accounts.selection())
            return

        self.order_account.set(picked)
        self._clear_order_round()

    def _clear_order_round(self):
        """
        把「這一位手上的東西」整批清掉：股票清單、持股、試算、股價、查回來的
        報價。換人（_on_order_account_changed）與換 Excel 檔
        （ui_background._forget_round）都走這裡——兩邊要清的是同一份東西，
        列兩份遲早會有一邊漏掉一項。

        清掉之後不自動重讀：讀 Excel 是使用者自己按的（「讀取持股」或
        「更新→自動計算」），點一下名字就自動跑一趟 COM，等於在人還在挑人的
        時候把 Excel 鎖起來。
        """
        self.order_holdings, self.order_names, self.order_prices = {}, {}, {}
        self.order_plans, self.order_quotes = {}, {}
        self.order_stock_pick.configure(values=[])
        self.order_stock_pick.set("")
        for row in list(self.order_rows):
            row["frame"].destroy()
        self.order_rows = []

        sheet = self._order_sheet()
        self.order_status.configure(
            text=(f"已選 {sheet}，接著按「讀取持股」。" if sheet else ""))
        self._resize_order_stock_column()
        self._update_order_quotes_ui()
        self._recompute_order_preview()
        self._update_order_exec_ui()
        self._apply_busy_state()

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
            show_info(self.root, "忙碌中", "現在有背景工作在跑，先等它結束才能切換模式。")
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
        # 執行按鈕上寫著「盤前」還是「盤中」（見 _order_exec_label），要跟著換。
        self._update_order_exec_ui()
        # 「新增」能不能按跟模式有關（見 _order_excel_buttons），切模式要重算一次。
        self._apply_busy_state()

    def _on_order_job_changed(self):
        """
        切作業。跟切模式同一條規矩：**股票清單整批清掉重選**（9.3 第 1 點）
        ——「比重」在出清股票是人填的設定，在買賣股票根本不存在（張數與價格
        來自 Excel 的下單試算 M14:N18），沿用舊的列會讓人以為兩邊是同一個
        數字。
        """
        if self.busy:
            self.order_job.set(self._order_job_last)
            show_info(self.root, "忙碌中", "現在有背景工作在跑，先等它結束才能切換作業。")
            return

        job = self.order_job.get()
        self._order_job_last = job

        # 第二列整列換掉。用 grid_remove()／grid() 而不是 pack_forget()：
        # grid_remove 記得住格子位置，再 grid() 回來會回到原位，pack_forget
        # 放回來會跑到這一列最後面（9.3 第 2 點，也是 order_ticks_entry 當初
        # 選擇「留著 disabled」而不是藏起來的原因）。
        for key, box in self.order_job_frames.items():
            if key == job:
                box.grid()
            else:
                box.grid_remove()

        # 「執行 更新→自動計算」住在「指定股票」那一格的「新增」右邊（見
        # ui_layout._build_order_stocks），而那一格三個作業共用，所以它得自己
        # 跟著作業收放——它算的是 M14:N18，只有買賣股票用得到那兩格。
        if job == orders.JOB_TRADE:
            self.order_auto_calc_button.grid()
        else:
            self.order_auto_calc_button.grid_remove()

        # 查到的即時報價跟著作廢，理由同切模式：清單都清空了，那份報價沒有
        # 任何一列在用。
        self.order_quotes = {}
        self._update_order_quotes_ui()

        # 「多輪直到出清」「自動更新股價」是出清・盤中限定的（見
        # _on_order_mode_changed）。切到別的作業要強制關掉並鎖住，不然從
        # 出清・盤中切過來時它們還留著勾選狀態，看起來像買賣股票也會跑多輪。
        if job != orders.JOB_CLEAR:
            self.order_multi_round.set(False)
            self.order_auto_price.set(False)
            self.order_multi_round_check.configure(state="disabled")
            self.order_auto_price_check.configure(state="disabled")
        else:
            self._on_order_mode_changed()

        # 「單位」兩個作業各畫一組單選鈕，綁的卻是同一個變數（見
        # ui_layout._build_order_unit）——切過來之後那個值在新作業可能是還沒接
        # 上的那一段（例如買賣股票的零股接上了、出清的零股還沒），單選鈕是灰
        # 的，變數卻還停在上面：按鈕上的字、送出去的量都會照一個按不到的選項
        # 走。跟 _order_intraday 要連作業一起問 order_mode 是同一類的錯——不會
        # 報錯，只會做了不該做的事。
        if not orders.unit_ready(job, self.order_unit.get()):
            self.order_unit.set(orders.UNIT_LOT)

        for row in list(self.order_rows):
            row["frame"].destroy()
        self.order_rows = []
        self._resize_order_stock_column()
        self._recompute_order_preview()
        self._update_order_exec_ui()
        self._apply_busy_state()

    def _on_order_unit_changed(self):
        """
        切整張／零股。買賣股票的「張數」「備註」兩欄是照單位算出來的（同一個
        試算股數拆兩段，見 orders.plan_trade_orders），所以預覽要跟著重算，
        不是只換執行按鈕上的字——只更新按鈕的話，畫面會停在另一半的數字上。
        """
        self._recompute_order_preview()
        self._update_order_exec_ui()

    def _order_intraday(self):
        """
        現在是不是「出清股票・盤中」。

        盤前／盤中是**出清作業自己的設定**（9.3 第 4 點），所以問「是不是盤中」
        一定要連作業一起問：從出清・盤中切到買賣股票的時候 order_mode 還留著
        "intraday"，只看它的話，買賣股票會莫名其妙跑去追價、跑去觸發更新股價
        巨集。這種錯不會報錯，只會做了一堆不該做的事。
        """
        return (self.order_job.get() == orders.JOB_CLEAR
                and self.order_mode.get() == "intraday")

    def _order_job_ready(self):
        """
        這個作業的行為接上了沒。買賣股票與全持股交易要到 9.7 第 4 步才接——
        在那之前選得到、第二列也看得到，但執行按鈕是灰的、預覽區會講一句
        為什麼（見 _recompute_order_preview）。刻意不是整個 disabled：這一步
        的重點就是把版面做出來給人看、把互換機制驗起來。
        """
        return self.order_job.get() in orders.JOBS_READY

    def _order_exec_label(self):
        """
        執行按鈕上的字跟著作業走——「按下去會動到什麼」要寫在按鈕上，不是一句
        通用的「開始下單」（9.3 最後一段，跟第四節那條原則同一個道理）。
        """
        job = self.order_job.get()
        unit = orders.UNIT_NAMES[self.order_unit.get()]
        if job == orders.JOB_CLEAR:
            when = "盤中" if self.order_mode.get() == "intraday" else "盤前"
            return f"開始出清（{unit}・{when}）"
        if job == orders.JOB_TRADE:
            return f"開始買賣（{unit}）"
        return "開始全持股交易"

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
        切到自動那一刻要跳確認——這不是畫面選項，是「程式會自己按下真的會送出
        委託的按鈕」，跟其他「按錯了大不了重選」的設定不是同一個等級的風險，要
        讓使用者確認過才生效，而不是勾了就算。

        **確認框留著，只是把說明縮成一句**（2026/08/31 使用者要求）。原本那三段
        （會自動按確認、節奏不變、送出後收不回要自己去委託查詢取消）是寫給第一次
        看到這個開關的人看的，而按它的人就是每天在用的那一位；擋一次的效果來自
        「要多按一下」，不是來自那三段字。
        """
        if self.busy:
            self.order_auto_confirm.set(self._order_auto_last)
            show_info(self.root, "忙碌中", "現在有背景工作在跑，先等它結束才能切換。")
            return

        if self.order_auto_confirm.get() and not ask_confirm(
                self.root, "切換為自動送出", "確定切換為自動送出委託單",
                confirm_style="primary"):
            self.order_auto_confirm.set(False)

        self._order_auto_last = self.order_auto_confirm.get()

    def _order_excel_buttons(self):
        """
        下單分頁裡「按下去會用 COM 動 Excel」的按鈕：只要有任何一條路正在動那份
        活頁簿（或那份活頁簿根本沒開著）就變灰（見 ui_background._apply_busy_state，
        它負責在四個旗標或 excel_open 變動時呼叫這裡）。

        原本「讀取持股」與「新增」各自只看自己那一個旗標——前者跑著的
        時候後者還是亮的，而兩條路都會一頁一頁 Activate 再跑巨集，交錯之後巨集會
        跑在別人剛切過去的那一頁上（見 excel_io._EXCEL_LOCK 的說明）。

        擋住而不是排隊：跟 _refresh_added_stock_price 對自己重複點擊的態度一致
        （那裡的註解有寫理由——下一次「新增」或「讀取持股」還會再有
        機會補上）。
        """
        busy = self._excel_in_use()
        # 「讀取持股」還要 Excel 真的開著才亮：這一顆做的事整個就是讀
        # 那份活頁簿，沒開著根本無事可做——跟更新分頁的「更新全部帳戶」同一個
        # 規矩（見 ui_sync._sync_buttons）。2026/08/31 之前這裡只看忙碌旗標，
        # 所以 Excel 沒開的時候「更新」那顆是灰的、這顆卻亮著，按下去換來一個
        # 「Excel 沒開著」的視窗；現在兩顆一起灰，那個視窗也跟著拿掉了（見
        # ui_background._require_excel）。
        self.order_refresh_button.configure(
            state="normal" if self.excel_open and not busy else "disabled")
        # 「更新→自動計算」跟上面那顆同一條規矩：它整個就是「開 Excel、跑兩支
        # 巨集、把結果讀回來」，Excel 沒開著或別人正在動同一份活頁簿都無事可做
        # （run_auto_calc 開頭那兩道 guard 擋的就是這兩件事）。2026/08/31 之前
        # 它不在這份清單裡，所以那兩種情況下按鈕還亮著，按下去被 _require_excel
        # 靜靜擋掉、什麼都不跳——就是這個述詞開頭講的「按了沒反應」。
        self.order_auto_calc_button.configure(
            state="normal" if self.excel_open and not busy else "disabled")
        # 「新增」只有盤中那條路會附帶跑巨集（見 add_order_stock 的說明），盤前
        # 完全不碰 COM，沒有理由跟著變灰——讀取 20 組帳戶要跑好幾分鐘，那段時間
        # 還是該能把股票加進清單。所以這一顆多看一個模式。它也不跟著 excel_open
        # 走：這顆真正在做的是「把這一檔加進清單」，那是純畫面的事，Excel 沒開
        # 只是附帶那次股價重讀會跳過（見 _refresh_added_stock_price 的守門）。
        add_busy = busy and self._order_intraday()
        self.order_add_button.configure(state="disabled" if add_busy else "normal")
        # 「執行帳戶」在有事情在跑的時候鎖住。換人會把股票清單、持股、試算整批
        # 清掉（見 _on_order_account_changed），而背景那一趟讀回來的是**上一位**
        # 的資料，兩件事撞在一起的結果是清單上點著 B、手上的試算卻是 A 的——
        # 不會報錯，只會拿 A 的數字算 B 的張數。依序執行中一樣鎖住：queue 在按下
        # 「開始下單」那一刻就凍結了，換人動不到它，但畫面會在跑到一半的時候被
        # 清空，看起來像壞掉。
        #
        # Treeview 的 disabled 是用 state() 給的，不是 configure(state=...)。
        # 它只是這道鎖的一半（看得出來、點不動），另一半在
        # _on_order_account_changed：真的收到選取變動時把它扳回去。
        self.order_accounts.state(["disabled"] if self._order_locked() else ["!disabled"])

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
        的說明），剛加的這一檔如果原本沒被最近一次「讀取持股」涵蓋到
        （例如本來沒持股），不補這一步就會一直停在讀不到／舊的數字。
        """
        raw = self.order_stock_pick.get().strip()
        if not raw:
            return
        code = raw.split(" ")[0].strip().upper()
        if any(row["code"] == code for row in self.order_rows):
            show_info(self.root, "已經加過了", f"{code} 已經在清單裡了。")
            return
        name = self.order_names.get(code) or (raw.split(" ", 1)[1].strip() if " " in raw else code)

        row = {"code": code, "name": name, "weight": tk.StringVar()}
        row["weight"].trace_add("write", lambda *_a: self._recompute_order_preview())
        if self.order_job.get() == orders.JOB_TRADE:
            # 買賣股票不填任何數字：張數與價格都來自各帳戶自己的下單試算
            # （規劃文件「一、買賣股票」只有「指定股票」跟「選帳戶」兩項設定）。
            row.pop("weight", None)
        elif self.order_mode.get() == "pre":
            row["price"] = tk.StringVar()
            row["price"].trace_add("write", lambda *_a: self._recompute_order_preview())
        self._build_order_stock_row(row)
        self.order_rows.append(row)
        self.order_stock_pick.set("")
        self._resize_order_stock_column()
        self._recompute_order_preview()
        if self._order_intraday():
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
        反正下一次「新增」或「讀取持股」還會再有機會補上。
        """
        # 這裡改看 _excel_in_use()：原本只看自己那一個旗標，所以那顆「讀取持股
        # 與報酬率」正在跑（5 檔 × N 個分頁的 HTTP，很慢）的時候按「新增」就會
        # 起第二條執行緒，兩邊都在 Activate → 跑巨集 → Activate → 跑巨集。
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
                            excel_io.run_update_price_macro(
                                excel, sheet, on_stuck=self._macro_stuck_notifier("更新股價", name))
                            sheets[name] = excel_io.read_sheet(sheet)
                # 巨集寫過 I4:I8，理由同 _order_read_worker。
                workbook.Save()
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
        主要操作，不值得為了它彈錯誤視窗（真的要查，「讀取持股」還在）。
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
        一檔股票一列，**一行排完**：買賣別、股票、移除、比重、價格（或試算／
        Excel 股價）。

        2026/08/31 之前是分兩行的，那是為了「指定股票」還在 300 像素寬的左欄時
        設計的——一行塞不下「名稱＋比重＋價格＋移除」，價格輸入框會被擠到剩沒
        幾個像素、打不進去字。那一格搬到右欄之後有一千多像素寬，前提不成立了，
        而且分兩行等於把橫向那一大片空白換成縱向的浪費：一檔 70 像素降到 39，
        省下來的高度全部流給下面的執行預覽（它是「帳戶 × 股票」的交叉，永遠
        不夠——見 ui_layout._build_order_tab）。

        每一列自己是一個 Frame（用 pack 疊起來，不是整個清單一張大 grid）：
        移除中間一列時不會留下空位，不必自己重新排列剩下的列。列**內部**才用
        grid，而且每一列的欄 minsize 都用同一組 ORDER_STOCK_COL_W，所以三檔的
        「移除」「比重」「價格」會上下對齊成一張小表——不是靠股票名稱剛好一樣長。

        買賣別跟著這一輪的方向走（見 _order_init_state 的 order_side：9.3 之後
        方向是作業算出來的，不是人選的），一整批
        股票共用同一個方向，加進來那一刻就定案——切買賣的時候整批清單會被
        清空重選（見 _on_order_job_changed），不會出現舊列還留著舊方向的
        情況。底色跟網站本身買紅賣綠的配色一致（Sell.TLabel／Buy.TLabel 在
        ui_layout._build() 裡註冊），不必看文字就認得出方向。

        `"price" in row` 決定要不要畫價格輸入框——盤中模式的 row 沒有這個
        key（見 add_order_stock），不是留白也不是畫一個不會被讀的欄位。

        盤中模式沒有價格輸入框的位置改顯示 Excel 讀回來的股價（self.
        order_prices，跟「讀取持股」讀回來的 order_names 同一批資料）——
        這不只是給人參考，開始下單那一刻會拿 order_prices 當第一輪
        chase_price 的 pricenow（見 start_order_execution），追價檔數還是
        要在下單前用這個基準再算一次邊界、查一次對手方第一檔（見
        orders.chase_price），這裡顯示的數字就是實際會拿去算價的那一個
        （不是另一條網頁現查的路，2026/08/29 使用者確認拿掉了）。這裡存了
        Label 物件本體（row["price_label"]）而不是只畫一次文字，因為
        _refresh_added_stock_price 每次有人按「新增」都會觸發一次背景重讀
        （2026/08/29 使用者要求），回來要能就地更新這一列的文字，不是只有
        新加的那一列，而是畫面上全部盤中列一起刷新——單按「讀取持股」不會
        觸發這個更新，只有「新增」股票才會。
        """
        # 買賣股票的方向是**逐檔逐帳戶**由試算的正負決定（規劃文件：正數為買、
        # 負數為賣），同一檔股票在甲帳戶是買、在乙帳戶可能是賣。所以這一列不掛
        # 買賣底色——掛了就是給一個對某些帳戶剛好相反的顏色。方向畫在右邊執行
        # 預覽的「買賣」欄，那裡才是一列一個帳戶。
        trade = self.order_job.get() == orders.JOB_TRADE
        buy = self.order_side.get() == orders.SIDE_BUY
        side_text = "" if trade else ("買" if buy else "賣")
        side_style = "TLabel" if trade else ("Buy.TLabel" if buy else "Sell.TLabel")

        block = ttk.Frame(self.order_stock_frame)
        block.pack(fill="x", pady=(0, 4))
        for index, width in enumerate(ORDER_STOCK_COL_W):
            block.columnconfigure(index, minsize=wide(width))

        if side_text:
            ttk.Label(block, text=side_text, style=side_style, width=2,
                     anchor="center").grid(row=0, column=0, sticky="w")
        ttk.Label(block, text=f" {row['code']} {row['name']} ", style=side_style).grid(
            row=0, column=1, sticky="w")
        ttk.Button(block, text="移除", bootstyle="danger-outline",
                  command=lambda: self.remove_order_stock(row)).grid(
            row=0, column=2, sticky="w", padx=(6, 0))

        if trade:
            # 這一格顯示勾選帳戶的試算股數（9.3：左邊那格顯示 M14:N18 試算值，
            # 唯讀）。試算是逐帳戶的，所以只有全部一致時才報一個數字，不一致就
            # 講範圍、叫人看右邊——報一個數字卻其實每個帳戶不同，比不報還糟。
            # 由 _update_trade_row_labels 在每次重算預覽時更新（勾的人會變）。
            #
            # 不設 wraplength：一列就是一行，換行會讓這一列比別列高、整張表的
            # 對齊就散了。「試算各帳戶不同（… ～ … 股），見右邊」那句最長，在
            # 這一格一千多像素的寬度下一行放得完。
            label = ttk.Label(block, text="", style="Hint.TLabel")
            label.grid(row=0, column=3, columnspan=2, sticky="w", padx=(12, 0))
            row["plan_label"] = label
            row["frame"] = block
            return

        # 比重、價格各自包一個小 Frame 再放進格子裡：label＋entry＋單位是一組
        # 三件套，讓它們在組內用 pack 貼在一起，組跟組之間才靠 grid 的欄對齊。
        weight_box = ttk.Frame(block)
        weight_box.grid(row=0, column=3, sticky="w", padx=(12, 0))
        ttk.Label(weight_box, text="比重").pack(side="left")
        ttk.Entry(weight_box, textvariable=row["weight"], width=6,
                 font=(self.family, FONT_SIZE)).pack(side="left", padx=(4, 0))
        ttk.Label(weight_box, text="%").pack(side="left", padx=(2, 0))

        price_box = ttk.Frame(block)
        price_box.grid(row=0, column=4, sticky="w", padx=(12, 0))
        if "price" in row:
            ttk.Label(price_box, text="價格").pack(side="left")
            ttk.Entry(price_box, textvariable=row["price"], width=8,
                     font=(self.family, FONT_SIZE)).pack(side="left", padx=(4, 0))
            ttk.Label(price_box, text="元").pack(side="left", padx=(2, 0))
        else:
            excel_price = self.order_prices.get(row["code"])
            price_text = f"Excel股價 {show(excel_price)} 元" if excel_price is not None else "Excel股價：讀不到"
            label = ttk.Label(price_box, text=price_text, style="Hint.TLabel")
            label.pack(side="left")
            row["price_label"] = label
        row["frame"] = block

    def remove_order_stock(self, row):
        row["frame"].destroy()
        self.order_rows.remove(row)
        self._resize_order_stock_column()
        self._recompute_order_preview()

    # ---------- 執行預覽 ----------

    def _order_execution_accounts(self):
        """
        這一輪要動誰——選中的那一位，沒選就是空清單。

        形狀跟以前 orders.order_accounts 回傳的一樣（多一個 "order"），下面
        三支 plan_* 照吃不必改。"order" 直接用選單上那一列的號碼，不重新從 1
        數起：那個號碼是「報酬率由低到高的第幾位」，執行預覽的「順序」欄要顯示
        的正是它——現在一次只做一位，欄位裡永遠寫 1 的話那一欄就只是噪音了。

        報酬率讀不到（None）的也照樣回傳，不像 orders.order_accounts 那樣踢進
        skipped：那支的顧慮是「用猜的決定誰先執行」，而這裡根本沒有第二位可以
        排序（見 _fill_order_accounts）。
        """
        sheet = self._order_sheet()
        if sheet is None:
            return []
        return [{"sheet": sheet, "return_rate": self.order_return_rates.get(sheet),
                 "order": self.order_account_rank.get(sheet, "－")}]

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
        比重／價格（或追價檔數）／選中的帳戶，任何一個變了就整份重算重畫——
        跟更新分頁 fill_sync_tree() 同一個做法，整份重建比自己追蹤哪一列該
        更新可靠。
        """
        # 數量那一欄的欄名跟著單位走：整張是「張數」、零股是「股數」（2026/09/01
        # 使用者指定）。單位寫在欄名上，格子裡就只放數字（見 orders._lots_text）。
        # 放在這裡而不是切單位的那個 handler：這支是「畫面上的設定變了就整份重算」
        # 的唯一入口，切單位、切作業、重讀 Excel 都會經過，欄名不會有哪一條路漏掉。
        self.order_preview.heading(
            "lots", text=orders.UNIT_COLUMN_TITLES[self.order_unit.get()])

        if not self._order_job_ready():
            self._render_order_preview([], [
                f"「{orders.JOB_NAMES[self.order_job.get()]}」還沒接上，"
                f"目前只有「{orders.JOB_NAMES[orders.JOB_CLEAR]}」可以執行"
                f"（落地順序見 docs/介面規劃.md 9.7）。"])
            return

        ordered = self._order_execution_accounts()
        hints = []
        if not ordered:
            hints.append(f"⚠ {self._order_no_account_text()}")

        if self.order_job.get() == orders.JOB_TRADE:
            # 買賣股票：不讀畫面上的任何數字，張數與價格都來自各帳戶自己那一頁
            # 的下單試算（見 orders.plan_trade_orders）。
            codes = [row["code"] for row in self.order_rows]
            preview = orders.plan_trade_orders(
                [{"code": row["code"], "name": row["name"]} for row in self.order_rows],
                ordered, self.order_plans, self.order_holdings, self.order_unit.get())
            self._update_trade_row_labels([account["sheet"] for account in ordered])
            if codes and not self.order_plans:
                hints.append("⚠ 還沒讀到下單試算，先按「讀取持股」。")
        elif self._order_intraday():
            stock_settings = self._order_stock_settings()
            ticks = self._order_ticks_setting()
            if ticks is None:
                preview = []
                hints.append("⚠ 追價檔數要填 0 以上的整數。")
            else:
                preview = orders.plan_intraday_orders(
                    stock_settings, ordered, self.order_holdings, ticks,
                    self.order_side.get(),
                    prices=self.order_prices, quotes=self.order_quotes)
        else:
            preview = orders.plan_stock_orders(self._order_stock_settings(), ordered,
                                               self.order_holdings, self.order_side.get())

        self._render_order_preview(preview, hints)

    def _update_trade_row_labels(self, sheets):
        """
        買賣股票：左邊每一列那句「試算多少股」跟著選中的帳戶更新。

        sheets 一定是 0 或 1 個（一次一位，見 _order_execution_accounts），
        2026/09/01 之前是勾選的那幾位，所以這裡還要處理「各帳戶試算不同」——
        現在沒有那種情況了，同一檔股票在這一輪只有一個數字。
        """
        for row in self.order_rows:
            label = row.get("plan_label")
            if label is None:
                continue
            qty = next((int((self.order_plans.get((sheet, row["code"])) or {}).get("qty") or 0)
                        for sheet in sheets), None)
            if qty is None:
                text = "還沒選帳戶"
            elif qty == 0:
                text = "這一位沒有試算"
            else:
                text = f"試算 {show(qty)} 股"
            label.configure(text=text)

    def run_auto_calc(self):
        """
        「更新→自動計算」：對**選中的那一位**依序觸發使用者既有的「更新股價」
        （I4:I8）與「自動計算」（依 M4:M8 目標比重試算，寫進 M14:N18）兩支巨集，
        跑完把試算重讀回來。取代原本的「初始化下單」（M ← O 欄加碼股數）——
        使用者確認不需要那條路，一律用自動計算的比重試算結果（2026/08/31）。

        **按下去直接跑，不跳確認**（2026/08/31 使用者要求拿掉）。原本會先跳一個
        列出「會對哪幾個分頁跑、M14:N18 會被蓋掉」的確認框，理由是從程式按跟
        自己在 Excel 上按不一樣——他看不到程式對哪幾個分頁按了幾次。現在那件事
        改由底部常駐狀態列講（「更新→自動計算 執行中……」，跑完換成跑了幾個
        分頁），從「事前擋一次」換成「事中、事後看得到」：這顆按鈕平常一輪要按
        好幾次，每次都要按掉一個只是重述按鈕名字的視窗太吵。

        巨集本身還是可能因為股價抓不到或輸入不完整跳出自己的訊息框（見
        excel_io.run_auto_calc_macro：9.6 那套事前檢查刻意還沒做，真的卡住了
        才回頭做）——那是 Excel 跳的，不是這裡跳的，拿掉確認框不影響它。
        """
        if self._excel_in_use() or not self._require_excel():
            return
        sheets = [account["sheet"] for account in self._order_execution_accounts()]
        if not sheets:
            show_info(self.root, "還沒選帳戶", self._order_no_account_text())
            return

        self.order_busy = True
        self._apply_busy_state()
        self.order_status.configure(text="更新／自動計算中…")
        # 底部常駐狀態列：一個分頁跑兩支巨集、還要一檔一檔打 Yahoo，是會讓人等
        # 的操作，而現在沒有確認框了，「按下去了、正在跑什麼」只剩這裡在講。
        # 現在一次只有一位，分頁名字列得出來（以前最多 20 個名字會把這一列撐爆，
        # 它是單行 Label，超出就直接切掉，所以那時只寫幾個分頁）。
        #
        # 刪節號用中文的「……」（兩個字、六個點）不是「…」：這句話是「還沒完，
        # 還在等」，中文全形刪節號本來就是六點（2026/08/31 使用者要求）。
        self._say(f"更新→自動計算 執行中……（{sheets[0]}）")
        threading.Thread(target=self._order_auto_calc_worker,
                         args=(self.path, sheets), daemon=True).start()

    def _order_auto_calc_worker(self, path, sheets):
        """
        背景執行緒：一頁 Activate 一次，依序跑「更新股價」「自動計算」，
        順手把新的試算讀回來（兩支巨集都只認 ActiveSheet，見 excel_io）。
        """
        import pythoncom

        pythoncom.CoInitialize()
        excel = workbook = sheet = None
        payload = {}
        try:
            with excel_io.opened(path, True) as (excel, workbook, _attached):
                plans, errors = {}, {}
                with excel_io.keep_active_sheet(workbook):
                    for name in sheets:
                        sheet, error = excel_io.find_sheet(workbook, name)
                        if sheet is None:
                            errors[name] = error
                            continue
                        excel_io.run_update_price_macro(
                            excel, sheet, on_stuck=self._macro_stuck_notifier("更新股價", name))
                        excel_io.run_auto_calc_macro(
                            excel, sheet, on_stuck=self._macro_stuck_notifier("自動計算", name))
                        plans[name] = excel_io.read_order_plan(sheet)
                # 兩支巨集都寫過格子（I4:I8、M14:N18），理由同 _order_read_worker。
                workbook.Save()
                payload = {"plans": plans, "errors": errors}
        except Exception as exc:
            payload = {"error": str(exc)}
        finally:
            sheet = excel = workbook = None
            pythoncom.CoUninitialize()
        self.queue.put(("order_auto_calc", payload))

    def _on_order_auto_calc(self, payload):
        """
        「更新→自動計算」的背景回話：把重讀回來的試算換掉，重畫執行預覽。

        底部狀態列也要收尾——按下去的時候把它設成「執行中……」（見 run_auto_calc），
        沒人換掉的話它會一直停在那句話上，看起來像還在跑。
        """
        self.order_busy = False
        self._apply_busy_state()

        if "error" in payload:
            self.order_status.configure(text="更新／自動計算失敗")
            self._say("更新→自動計算 失敗。")
            show_error(self.root, "更新／自動計算失敗", payload["error"])
            return

        for name, plan in payload["plans"].items():
            for code, values in plan.items():
                self.order_plans[(name, code)] = values

        errors = payload["errors"]
        note = f"　（找不到分頁：{'、'.join(errors)}）" if errors else ""
        # 名字寫出來，不寫「N 個分頁」：一次只有一位，而這一趟是**會改 Excel**
        # 的（M14:N18 被重算），改到誰身上要看得見（跟 _on_order_data 那句同一個
        # 態度）。
        if payload["plans"]:
            done = f"已對 {'、'.join(payload['plans'])} 跑過「更新／自動計算」。{note}"
        else:
            done = f"沒有跑到任何分頁。{note}"
        self.order_status.configure(text=done)
        self._say(f"更新→自動計算：{done}")
        self._recompute_order_preview()

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
            # 真的會發生的一筆交易（見 ui_layout._build_order_preview 的 tag_configure）。
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
            # 順序要跟 ui_layout._build_order_preview 的 columns 一模一樣：
            # 順序／帳戶／買賣／股票／張數／價格／持股／備註。Treeview 是照位置
            # 對欄位的，換了順序卻沒改這裡不會報錯，只是每一欄顯示別欄的值。
            #
            # 「張數」欄：買賣股票選零股時要把「另有幾張沒送」寫出來（9.4），
            # 選整張只顯示張數；其他作業沒有這個問題，就是一個數字（沒有
            # lots_text 就退回 lots）。持股最小單位是 1 股，不需要小數點；股數
            # 本來就可能上看百萬，千分位才看得出位數（util.show 是全專案統一
            # 用的數字顯示格式）。
            self.order_preview.insert("", "end", values=(
                item["order"], item["sheet"], side_names.get(item["side"], item["side"]),
                f"{item['code']} {item['name']}",
                item.get("lots_text") or item["lots"], price_text, show(item["held_qty"]),
                item["note"],
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
        # 面板高度跟欄寬是同一種「清單結構變了才要重算」的東西，四個呼叫端
        # （新增／移除／切模式／切作業）也完全一樣，所以掛在這裡一起做，不另外
        # 去那四個地方各加一行。
        self._resize_order_stock_panel()

    def _resize_order_stock_panel(self):
        """
        「指定股票」那一格的高度＝目前檔數 × 一檔的高度，**最多三檔**
        （使用者指定「留三檔的空間就好」），第四、五檔靠右邊的捲軸。

        不是固定三檔高：它跟執行預覽上下疊在同一欄（見 ui_layout 的
        _build_order_tab），空著的時候佔的每一像素都是從預覽身上扣的，而預覽
        的列數是「帳戶 × 股票」的交叉，20 組 × 5 檔就是 100 列——清單還沒加
        東西的時候，那些高度該歸預覽。

        清單空的時候留一檔高，不縮到 0：整格塌成一條線看起來像壞掉，而且
        「新增」進來的第一檔要有地方落地。
        """
        shown = min(max(len(self.order_rows), 1), ORDER_STOCK_ROWS_SHOWN)
        self.order_stock_canvas.configure(height=ORDER_STOCK_ROW_H * shown)
        # 改完 Canvas 高度還要自己把分隔推過去：ttk 的 Panedwindow 只在第一次
        # 排版時照子元件的 reqheight 定分隔位置，之後子元件長高了它不會跟——
        # 實測加到五檔，那一格還是停在一檔高，畫面上看起來就是「新增沒反應」
        # （其實加進去了，只是被壓在看不見的地方）。update_idletasks 不能省：
        # reqheight 要等這一輪版面算完才是新的值。
        box = self.order_stock_canvas.master
        box.update_idletasks()
        try:
            self.order_body_paned.sashpos(0, box.winfo_reqheight())
        except tk.TclError:
            pass    # 視窗還沒排版出來（開機那一次），初始位置本來就是對的

    def _resize_order_sheet_column(self):
        """
        「帳戶」欄寬跟著這次「讀取持股」讀回來的帳戶名單重量一次——
        只在 _fill_order_accounts 換了一批名單時呼叫，理由跟
        _resize_order_stock_column 一樣：名單只在讀取的當下換一批，不會因為
        使用者操作畫面上其他東西（換帳戶、改比重）而變動。
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
    # 不能跟更新分頁或下單依序執行同時搶同一顆瀏覽器。

    def fetch_order_quotes(self):
        """
        觸發背景查詢；結果回來見 _on_order_quotes_fetched。

        報價是公開資料，不因帳戶而不同（跟 order_exec_prices 那份 Excel
        股價「哪個帳戶先讀到就先用哪個」同一種態度）——這裡拿選中的那位去
        登入，純粹是借「已經登入」這件事開 FastQuote 彈出視窗，不代表這批
        報價只給那個帳戶用。
        """
        if self.busy or self.order_quotes_busy:
            return
        if not self._order_quotes_available():
            return

        codes = sorted({row["code"] for row in self.order_rows})
        if not codes:
            show_info(self.root, "還沒有股票", "請先加入至少一檔股票。")
            return

        ordered = self._order_execution_accounts()
        if not ordered:
            show_info(self.root, "還沒選帳戶",
                      f"{self._order_no_account_text()}（查詢委買賣要借一組帳戶登入。）")
            return
        order_number = self._order_number_for_sheet(ordered[0]["sheet"])
        if order_number is None:
            show_error(self.root, "找不到帳戶", f"{ordered[0]['sheet']} 對不到任何一組帳號。")
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
        # 回傳的就是要送回主執行緒的那份 payload（見 ui_background 的 simple_jobs：
        # 那張表直接把 job 的回傳值當 payload 送出去）。
        return {"quotes": quotes}

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
            show_error(self.root, "查詢委買賣失敗", text)
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

    def _order_quotes_available(self):
        """
        「查詢委買賣」現在有沒有意義：只有「出清股票」有盤前／盤中這個設定，
        而追價比價是盤中限定的（9.3 把盤前／盤中降級成出清作業自己的設定）。
        """
        return self._order_intraday()

    def _update_order_quotes_ui(self):
        if self.order_quotes_busy:
            self.order_quotes_button.configure(text="查詢中…", state="disabled")
            return
        available = self._order_quotes_available()
        self.order_quotes_button.configure(
            text="查詢委買賣", state="normal" if available and not self.busy else "disabled")

    def _order_number_for_sheet(self, name):
        """分頁名 -> 第幾組帳號。找不到（理論上不會，執行預覽的名字都從 trader_of 長出來）就回 None。"""
        return next((order for order, sheet in self.trader_of.items() if sheet == name), None)
