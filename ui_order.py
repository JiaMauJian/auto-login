"""
下單分頁「按下開始下單之前」的那一半：收使用者的輸入（選股票、填比重、勾
帳戶）、讀 Excel（持股、B22 報酬率、成交價）、查即時委買賣一，算出執行預覽。

按下去之後的另一半在 ui_order_exec.py——那裡從凍結這一輪的設定開始，一筆
一筆送出、跑完一輪再決定要不要接下一輪。界線是「這一輪要送什麼」（這裡）
跟「怎麼把它送出去」（那裡），見那個檔案開頭的說明。

「盤前」（股票／比重／價格設定）跟「盤中」（股票／比重／追價檔數設定，價格
用 Excel 成交價＋下單前查對手方第一檔算出來的）共用同一套選帳戶、算執行
預覽、依序執行的機制，只有「股票設定要填什麼」「怎麼組出執行清單」不一樣
（見 `_on_order_mode_changed`／ui_order_exec 的 `start_order_execution`）。

比重→張數、帳戶依 B22 報酬率排序、組出預覽清單、追價檔數換算價格，全部是
orders.py 的純函式，這裡只負責收輸入、讀 Excel（含成交價，盤中新增股票／
讀取試算時順便觸發「更新股價」巨集）、查即時對手方第一檔、把結果畫出來。
"""

import threading
import tkinter as tk

import ttkbootstrap as ttk

import excel_io
import order_fill
import orders
import stockinfo
from ui_common import (
    FONT_SIZE, ORDER_STOCK_ROW_H, ORDER_STOCK_ROWS_SHOWN, PRICE_PENDING_TEXT,
    col_width, show_error, show_info, wide,
)
from util import show


# 「指定股票」每一列的欄位起點（照 10 級字的像素，實際用 wide() 換算）。
# 每一列各自是一個 Frame，但欄的 minsize 全部用這同一組，所以三檔的「比重」
# 「價格」「移除」會上下對齊成一張小表——不是靠股票名稱剛好一樣長。
#
# 股票那一欄 110 夠放「2330 台積電」這種四碼＋三到五個中文；真的更長也不會被
# 切掉（minsize 是下限不是上限），只是那一列的後面幾欄會往右推、跟別列對不齊。
ORDER_STOCK_COL_W = (26, 110, 130, 140, 60)

# 「查詢委買賣」整條路連不上時，錯誤視窗最上面那段人話（下面接原始錯誤）。
# 重點是講清楚「不影響下單」——這一步本來就只是提前看價格，沒查到的話下單前
# 那一刻還是會照原本方式再查一次（見 ui_order_exec 追價那段）。
QUOTES_OFFLINE_HINT = (
    "連不上行情伺服器，這一趟沒有查到任何一檔。"
    "下單前還是會照原本方式即時查一次，不影響下單。"
)


class UiOrderMixin:
    # ---------- 下單分頁：盤前模式 ----------

    def _order_init_state(self):
        """SyncApp.__init__ 呼叫一次。"""
        self.order_rows = []              # 這一輪加進來的股票設定列（見 add_order_stock）
        self.order_holdings = {}          # (分頁名, 股票代號) -> 股數，按「讀取試算」才會更新
        # (分頁名, 股票代號) -> Treeview iid，_render_order_preview 每次重畫都會
        # 整份重建；出清進度那一欄靠這份對照單格更新（見 ui_order_exec.
        # _refresh_order_progress_cells），不必為了刷新「▒」重畫整張表。
        self.order_preview_iid = {}
        # 「指定股票」下拉的候選是兩份資料併起來的（見 _rebuild_order_names）：
        # order_stock_catalog 是「讀取ＯＯ持股」讀回來的第一個分頁 D4~D13——
        # 跟勾了誰無關，2026/09/02 起這顆按鈕搬到左邊「執行帳戶」，一開檔就
        # 可以按；order_holding_labels 是「讀取試算」讀到那幾位帳戶手上實際
        # 有的那幾檔（會併進來，理由見 excel_io.read_stock_list：某一位手上有、
        # 第一頁沒列到的那幾檔不併就選不到）。order_names 是兩者合併後的結果，
        # 真正給畫面用的就是這一份，不要直接改它。
        self.order_stock_catalog = []     # [(代號, D欄原文), ...]
        self.order_stock_list_sheet = None  # 候選清單來自哪個分頁名；還沒讀過是 None
        self.order_holding_labels = {}    # 股票代號 -> 名稱，來自「讀取試算」讀到的持股列
        self.order_names = {}             # 股票代號 -> 名稱，畫面顯示用（合併後，見上）
        self.order_prices = {}            # 股票代號 -> Excel I 欄讀回來的股價；盤中模式這份就是
                                           # chase_price 的 pricenow 來源（見 start_order_execution），
                                           # 不只是畫面顯示用（跟 order_names 平行）
        self.order_return_rates = {}      # 分頁名 -> B22 報酬率或 None（讀不到）
        # (分頁名, 股票代號) -> {"name", "qty"(股數，正買負賣), "price"}：買賣股票作業
        # 用的下單試算 M19:N28（見 excel_io.read_order_plan）。**每個帳戶各有一份**
        # ——它是那一頁自己的試算結果，不像股價那樣全帳戶共用一個值，所以 key 要
        # 帶分頁名，跟 order_holdings 同一種形狀，不是 order_prices 那種。
        self.order_plans = {}
        # 「執行帳戶」表（左欄的 Treeview，一列一位，列首是 ☐／☑）：**可以勾
        # 好幾位，一次跑完**（2026/09/02 使用者決定，推翻 09/01 那版「一次一位」，
        # 見 docs/介面規劃.md 9.3 第 5 點）。理由是買賣股票的操作變成「勾帳戶 →
        # 選股票 → 執行」：張數與價格全在各帳戶自己那一頁的下單試算裡，人不必
        # 逐位確認任何數字，一位一位選反而跟手動下單一樣慢。
        #
        # 存的是**分頁名的集合**，不是畫面上那一列的文字——那一列還帶著勾選
        # 記號、順位號碼與報酬率，三者都會隨著補讀 B22 而改變，拿它當識別的話，
        # 每重畫一次清單就對不回同一個人。
        self.order_checked = set()
        self.order_account_rank = {}      # 分頁名 -> 報酬率由低到高的順位（讀不到是「－」）
        self.order_account_order = []     # 清單上由上到下的分頁名＝執行順序
        self.order_account_label = {}     # 分頁名 -> 那一列不含 ☐／☑ 的文字
        # 上一次「讀取試算」真的讀到的分頁。勾了帳戶但還沒讀的那幾位，手上一格
        # 資料都沒有，看起來會跟「下單試算是空的」一模一樣——執行預覽要分得出
        # 這兩件事（見 orders.REASON_NOT_LOADED）。
        self.order_loaded = set()
        # 「正在去 Excel 讀帳戶清單」（見 refresh_order_accounts）。那一趟只讀
        # B22，不跑巨集，但動的是同一份活頁簿，所以照樣要算進 _excel_in_use()。
        self.order_rates_busy = False
        self.order_busy = False           # 「讀取試算」還在跑
        self.order_stock_list_busy = False  # 「讀取ＯＯ持股」還在跑
        # 股票代號 -> {"bid","ask","last"}，「查詢委買賣」按鈕查回來的即時
        # 委買賣一（見 fetch_order_quotes／stockinfo.quote 的形狀）。有這份資料
        # 時 orders.plan_intraday_orders 會直接算出實際會送出的價格，不是只留
        # 一句「下單前會再查」的說明文字（2026/08/29 使用者要求）。切模式會清空
        # 重來，理由跟 order_rows 一樣。
        #
        # 只用股票代號當 key、不分整股零股，是因為整份預覽同一時間必定只有一種：
        # 「出清零股」跟「出清整張・盤中」是 _recompute_order_preview 裡互斥的
        # 兩條作業分支。哪天這個前提變了（同一份預覽混著兩種單位），這份字典的
        # key 要跟著改成 (代號, 是不是零股)——兩者是兩本不同的簿子，同一時刻
        # 可以差好幾檔，共用一個 key 會靜靜拿錯。
        self.order_quotes = {}
        self.order_quotes_busy = False
        self._order_quotes_requested = []  # 上一次按「查詢委買賣」實際問了哪幾檔，回話時算漏了誰用

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
        # 沒有一個需要人選方向（出清永遠是賣；買賣股票由 M19:M28 的正負逐檔
        # 決定；全持股交易也是算出來的）。方向只該出現在執行預覽的「買賣」欄
        # 跟紅綠底色。這個變數留著是因為 orders.py 那幾支 plan_* 還是收一個
        # side 參數——差別在它現在由作業決定，不是由人選。
        self.order_side = tk.StringVar(value=orders.SIDE_SELL)

        # 自動送出開關：關＝半自動，跟原本一樣停在委託確認視窗給人看、給人按，
        # 視窗關掉後還是要人按「下一筆」才送下一筆（半自動流程裡使用者唯一的
        # 節奏控制點，這條沒有要動）；開（預設，2026/09/02 使用者改）＝程式
        # 自己按下確認視窗裡的「確認」，委託真的送出去，視窗一關就自動接下一筆
        # （見 ui_order_exec._on_order_dialog_closed），整批自己跑到結束或遇到
        # 失敗才停，不會停下來等人按。2026/08/28 加這個開關時 AskUserQuestion
        # 確認過「節奏不變、不管開不開都要按一次下一筆」，2026/09/02 使用者
        # 推翻了這個決定——自動送出時不再要求多按那一下。
        self.order_auto_confirm = tk.BooleanVar(value=True)
        self._order_auto_last = True

        # 「開始下單」之後那一整套凍結狀態（queue、這一輪的模式／追價檔數／
        # 買賣方向、多輪的第幾輪…）在 ui_order_exec.py，跟這裡「畫面上現在填了
        # 什麼」的狀態刻意分開——見那個檔案開頭的說明。
        self._order_exec_init_state()

    def _rebuild_order_names(self):
        """
        把「指定股票」下拉可以選的東西重算一次：order_stock_catalog（讀取ＯＯ
        持股讀到的第一個分頁 D4~D13）併上 order_holding_labels（讀取試算讀到
        那幾位帳戶手上實際有的股票）。兩份資料各自獨立更新（見兩個 _on_order_*
        handler），順序無所謂，每次任何一份變了就重算整份，不嘗試局部更新。

        併起來而不是只用第一頁，是因為某一位手上有、第一頁沒列到的那幾檔，
        不併的話根本選不到；反過來第一頁有、大家都沒有的那幾檔選得到，執行
        預覽會一列一列寫「這一位沒有這檔」。
        """
        self.order_names = {}
        for code, label in self.order_stock_catalog:
            self.order_names.setdefault(code, label.split("(")[0].split("（")[0].strip())
        for code, label in self.order_holding_labels.items():
            self.order_names.setdefault(code, label)
        choices = sorted(f"{code} {name}" for code, name in self.order_names.items())
        self.order_stock_pick.configure(values=choices)

    def _update_order_stock_list_button(self):
        """
        左欄那顆按鈕的文字固定是「讀取持股」，不再依讀到的分頁名動態改字
        （2026/09/02 使用者要求改回固定文字）。
        """
        self.order_stock_list_button.configure(text="讀取持股")

    # ---------- 讀取ＯＯ持股（第一個分頁的股票候選，跟勾了誰無關） ----------

    def refresh_order_stock_list(self):
        """
        只讀**第一個分頁**的 D4~D13，當「指定股票」下拉的候選（2026/09/02
        使用者指定，見 excel_io.read_stock_list）。

        **不看勾了誰、也不用勾才能按**：這顆按鈕搬到左邊「執行帳戶」之後
        （2026/09/02 使用者要求），跟「執行帳戶」清單一樣是「Excel 那一頭的
        答案」，開檔就問得到，不必等人先勾人。要讀哪幾位的試算是另一顆
        「讀取試算」的事（見 refresh_order_plans），這裡完全不碰。
        """
        # 看 _excel_in_use() 而不是只看 order_stock_list_busy：更新分頁的寫入、
        # 「讀取試算」、「新增」股票附帶的股價重讀，動的都是同一份活頁簿。
        if self._excel_in_use() or not self._require_excel():
            return
        self.order_stock_list_busy = True
        self._apply_busy_state()
        self.order_status.configure(text="讀取中…")
        threading.Thread(target=self._order_stock_list_worker,
                         args=(self.path,), daemon=True).start()

    def _order_stock_list_worker(self, path):
        """背景執行緒：只讀第一個分頁的名字與 D4~D13，不跑巨集、不動其他分頁。"""
        import pythoncom

        pythoncom.CoInitialize()
        excel = workbook = None
        payload = {}
        try:
            with excel_io.opened(path, False) as (excel, workbook, _attached):
                sheet = excel_io.first_visible_sheet(workbook)
                payload = {
                    "sheet_name": sheet.Name.strip() if sheet is not None else None,
                    "stocks": excel_io.read_stock_list(workbook),
                }
        except Exception as exc:
            payload = {"error": str(exc)}
        finally:
            excel = workbook = None
            pythoncom.CoUninitialize()
        self.queue.put(("order_stock_list", payload))

    def _on_order_stock_list(self, payload):
        self.order_stock_list_busy = False
        self._apply_busy_state()

        if "error" in payload:
            self.order_status.configure(text="讀取失敗")
            show_error(self.root, "讀取失敗", payload["error"])
            return

        self.order_stock_catalog = payload["stocks"]
        self.order_stock_list_sheet = payload["sheet_name"]
        self._update_order_stock_list_button()
        self._rebuild_order_names()

        if payload["sheet_name"] is None:
            done = "沒有看得見的分頁，讀不到股票清單。"
        elif not payload["stocks"]:
            done = f"已讀取「{payload['sheet_name']}」，D4~D13 沒有股票代號。"
        else:
            done = f"已讀取「{payload['sheet_name']}」的持股清單（{len(payload['stocks'])} 檔）。"
        self.order_status.configure(text=done)
        self._recompute_order_preview()

    # ---------- 讀取試算（勾選那幾位的持股／股價／下單試算） ----------

    def refresh_order_plans(self):
        """
        重新讀 Excel：把**勾選的那幾位**的持股（D~F 欄）、B22、股價（I 欄）讀
        一次，買賣股票作業再多讀下單試算 M19:N28。**一定要先勾至少一位**——
        這顆按鈕做的事整個就是「勾了誰、去讀誰」，沒有對象可讀（見下面
        _order_sheets 為空就擋住）。

        讀的範圍就是勾了誰讀誰：09/01 那版一次只讀一位，是因為當時的操作是
        一次處理一位；09/02 改成可以勾好幾位一次跑完，這裡自然跟著變成「勾了
        幾位就讀幾頁」。**沒有「全部帳戶都讀一遍」這條路**——20 個分頁裡沒勾的
        那幾位這一輪一格都用不到，盤中模式更貴：每一頁都要各跑一次「更新股價」
        巨集，一頁 5 檔就是 5 次 Yahoo HTTP（todo.txt C2 記的那 100 次就是這麼
        來的）。要跑全部就自己按「全選」，那是人的決定，不是程式偷偷替他決定。

        盤中模式讀之前先觸發一次「更新股價」巨集再讀（2026/08/29 使用者確認：
        盤中追價用的成交價來自這裡讀到的 order_prices，新增股票／讀取試算
        這一步就要盡量拿到新的價格，不能留著上次殘留的舊數字）；盤前
        模式的價格是人手動填的，不需要 Excel 股價，維持原本只讀不觸發巨集。

        「指定股票」下拉的候選不歸這裡管，見 refresh_order_stock_list——兩顆
        按鈕以前是同一顆，2026/09/02 使用者要求拆開：候選清單開檔就該看得到，
        不該卡在「還沒勾帳戶」。
        """
        # 看 _excel_in_use() 而不是只看 order_busy：更新分頁的寫入、「讀取ＯＯ
        # 持股」、「新增」股票附帶的股價重讀、多輪之間的重讀，動的都是同一份
        # 活頁簿（見那個述詞）。
        if self._excel_in_use() or not self._require_excel():
            return
        names = self._order_sheets()
        if not names:
            show_info(self.root, "還沒勾帳戶", self._order_no_account_text())
            return

        # 追價的兩條路（出清整張・盤中、出清零股）才要先跑「更新股價」：它們的
        # 委託價以 Excel I 欄的成交價為基準（見 orders.chase_price），基準是舊的
        # 就整輪都算錯。盤前的價格是人填的、買賣股票的價格來自試算，兩者都不必
        # 為此付一頁 10 次 Yahoo HTTP 的代價。
        run_macro = self._order_uses_excel_price()
        # 買賣股票的張數與價格來自下單試算 M19:N28，那是另外幾格，只有這個作業
        # 要讀——其他作業讀它只是多 10 格 COM 往返。
        read_plan = self.order_job.get() == orders.JOB_TRADE
        self.order_busy = True
        self._apply_busy_state()
        self.order_status.configure(text="更新股價、讀取中…" if run_macro else "讀取中…")
        threading.Thread(target=self._order_plans_worker,
                         args=(self.path, names, run_macro, read_plan), daemon=True).start()

    def _order_plans_worker(self, path, names, run_macro, read_plan=False):
        """
        背景執行緒：用 COM 讀 D~F 欄、B22、I 欄。run_macro 為真的話，每個
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
                # 巨集寫過 I4:I13 就要存檔，不然沒接上使用者既有視窗時
                # close_workbook 會 Close(False) 把這次更新的股價丟掉。
                if run_macro:
                    workbook.Save()
                payload = {"sheets": sheets, "errors": errors}
        except Exception as exc:
            payload = {"error": str(exc)}
        finally:
            sheet = excel = workbook = None
            pythoncom.CoUninitialize()
        self.queue.put(("order_plans", payload))

    def _on_order_plans_data(self, payload):
        self.order_busy = False
        self._apply_busy_state()

        if "error" in payload:
            self.order_status.configure(text="讀取失敗")
            show_error(self.root, "讀取失敗", payload["error"])
            return

        # 持股／股價／試算整份換掉：這一趟讀的就是勾選那幾位的全部，留著
        # 上一批的只會讓「這一位到底有沒有這檔」變成看運氣（改勾選那一刻
        # 其實已經清過一次，見 _on_order_account_toggled，這裡是第二道）。
        #
        # order_return_rates 例外，不清：它是**所有**帳戶的（那份清單本身，見
        # refresh_order_accounts），清掉的話讀幾位就把左邊其他人整個弄丟。
        # 讀到的那幾位 B22 順手更新（下面那一行）——這一趟本來就讀到了，不用
        # 再跑一次帳戶清單那條路。
        self.order_holdings, self.order_holding_labels, self.order_prices = {}, {}, {}
        self.order_plans = {}
        self.order_loaded = set(payload["sheets"])
        for name, data in payload["sheets"].items():
            self.order_return_rates[name] = data["return_rate"]
            for code, plan in (data.get("plan") or {}).items():
                self.order_plans[(name, code)] = plan
            for row in data["rows"]:
                self.order_holdings[(name, row["code"])] = row["qty"]
                self.order_holding_labels.setdefault(
                    row["code"], row["label"].split("(")[0].split("（")[0].strip())
                # 哪個帳戶先讀到就先用哪個，跟 _on_order_price_refresh 彙整
                # order_exec_prices 同一個態度——同一檔股票的 Excel 股價不會
                # 因為帳戶不同而不同，不比對多帳戶是否一致。讀不到（None）
                # 就不佔位，讓 add_order_stock 那邊看到「沒有」而不是猜一個值。
                if row["price"] is not None:
                    self.order_prices.setdefault(row["code"], row["price"])
        self._rebuild_order_names()
        self._fill_order_accounts()

        # 盤中模式的股票列不畫價格輸入框，改顯示 order_prices（見
        # _build_order_stock_row 的說明）——這一趟讀完要就地把畫面上已加入
        # 股票的價格文字補上，不然新增股票時附帶的這次重讀只更新了
        # self.order_prices，畫面還停在加入那一刻的舊快照（2026/09/03 起
        # add_order_stock 改成每次都呼叫這裡，取代原本專職做這件事的
        # _refresh_added_stock_price）。
        for row in self.order_rows:
            label = row.get("price_label")
            if label is None:
                continue
            excel_price = self.order_prices.get(row["code"])
            price_text = f"Excel股價 {show(excel_price)} 元" if excel_price is not None else "Excel股價：讀不到"
            label.configure(text=price_text)

        # 「版面對不對得上」不在這裡問：那件事在**開檔那一刻**就驗過了（A22
        # 錨點，見 ui_background.check_excel_layout），對不上的話 excel_open
        # 是 False，「讀取試算」這顆按鈕根本按不下去，走不到這裡。

        # 一次可能好幾位，所以句子要能列出名字（讀到 4 位以上只報數字，不然
        # 那一行會長到把狀態列擠爆）。讀不到任何一位的時候整句換掉，不是把
        # 名字的位置填一句「沒有分頁」——那會變成「已讀取 沒有分頁 的試算」
        # 這種讀起來卡住的句子。這顆按鈕一定是勾了才按得下去（見
        # refresh_order_plans），所以這裡不會有「根本沒請求」那種空。
        errors = payload["errors"]
        if len(payload["sheets"]) > 3:
            done = f"已讀取 {len(payload['sheets'])} 位的持股與試算。"
        elif payload["sheets"]:
            done = f"已讀取 {'、'.join(payload['sheets'])} 的持股與試算。"
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
        每次都真的重讀一遍，不記「掃過了沒」——它一位只讀兩格（A22 錨點、B22
        報酬率），20 位也只是幾十毫秒（真正貴的是巨集那些 Yahoo HTTP），
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
        """背景執行緒：列出交易人分頁與 B22（見 refresh_order_accounts）。"""
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
        現在能不能改帳戶的勾選。有任何一條路在動那份活頁簿（_excel_in_use）、
        或這一輪委託還沒跑完（order_exec_queue）都不行——理由見
        _order_excel_buttons 與 _on_order_account_toggled。
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
        重畫左欄的「執行帳戶」表：整份清掉重建，勾好的那幾位留著。

        排序用 orders.order_accounts，排出來的就是**執行順序**（規格「報酬率
        低的先執行」）：號碼小的先送。2026/09/01～09/02 之間一次只跑一位，那
        段期間它排的只是「建議的處理順序」，現在一次跑好幾位，這個號碼又變回
        真正的執行順序了。

        B22 讀不到的那幾位排在最後、報酬率欄寫「讀不到」，但**照樣勾得到**：
        報酬率在這裡只決定順序，不決定能不能下單。他們之間的先後就照活頁簿
        裡分頁本來的順序（orders.order_accounts 把他們原樣放在 skipped 裡），
        不假裝算得出一個順位——所以號碼欄寫「－」，不是給一個看起來像順位的
        數字。

        每一列的 iid 就是分頁名，所以重建之後把勾選畫回去只要看
        `sheet in self.order_checked`——那一列的文字（勾選記號、號碼、百分比）
        每次補讀 B22 都可能變，拿文字當識別的話，勾好的帳戶會在每次重畫時被
        清掉一次。

        勾選狀態存在 self.order_checked（純資料），畫面只是它的投影：Treeview
        本身沒有 checkbox，列首那個 ☐／☑ 就是 #0 那一欄文字的一部分。所以這裡
        不必像 09/01 那版擋 `<<TreeviewSelect>>`——重畫清單不會誤觸任何東西，
        真正會改變勾選的只有使用者點下去那一下（_on_order_account_toggled）。
        """
        ordered, skipped = orders.order_accounts(
            [{"sheet": name, "return_rate": self.order_return_rates.get(name)}
             for name in self._order_known_sheets()])

        rows = []
        for account in ordered:
            # B22 存的是小數（0.185222... 代表 18.5%），畫面要顯示的是百分比，
            # 這裡要乘 100——漏了這一步會把 18.5% 顯示成 0.2%，跟現金查詢
            # Amount 除以 100 是同一種「單位不對但不會報錯」的坑（CLAUDE.md）。
            rows.append((account["sheet"], account["order"],
                         f"{account['order']}　{account['sheet']}",
                         f"{account['return_rate'] * 100:.1f}%"))
        for account in skipped:
            # 沒有號碼可給——排序這件事對他不成立，用「－」佔位，不要給一個
            # 看起來像順位的數字（執行預覽的「順序」欄顯示的也是這個值）。
            rows.append((account["sheet"], "－", f"－　{account['sheet']}", "讀不到"))

        self.order_accounts.delete(*self.order_accounts.get_children())
        self.order_account_rank = {}
        self.order_account_label = {}
        self.order_account_order = [sheet for sheet, _rank, _name, _rate in rows]
        for sheet, rank, name_text, rate_text in rows:
            self.order_account_rank[sheet] = rank
            self.order_account_label[sheet] = name_text
            self.order_accounts.insert("", "end", iid=sheet,
                                       text=self._order_account_text(sheet, name_text),
                                       values=(rate_text,))
        # 從名單上消失的那幾位（換了 Excel 檔、Excel 裡刪了一頁）連著勾選一起
        # 忘掉，不留一個對不到任何一列的名字——那種名字看不見卻還會被算進
        # 「這一輪要跑誰」，是最難發現的一種錯。
        gone = self.order_checked - set(self.order_account_rank)
        if gone:
            self.order_checked -= gone
        self._resize_order_sheet_column()
        self._order_excel_buttons()

    def _order_account_text(self, sheet, name_text):
        """
        帳戶那一列的文字：`✅ 1　交易人A`。勾選記號直接寫在 #0 那一欄的文字裡，
        不另外開一欄——這一格只有 260 像素寬，已經放了名字與報酬率兩欄，再切
        一欄出來給一個字元的記號，名字那一欄就會開始被切掉。

        用 ✅／⬜ 而不是 ☑／☐：Treeview 一格文字只有一種字級，符號沒辦法只放大
        自己、名字維持原大小——換成 Windows 上會走彩色 emoji 字型畫的符號，
        同樣字級視覺上就大顆很多，不必真的調整字型大小（2026/09/02 使用者確認）。
        """
        return f"{'✅' if sheet in self.order_checked else '⬜'} {name_text}"

    def _order_sheets(self):
        """
        這一輪要動哪幾位（分頁名），照執行順序（報酬率低的先）排好。

        沒勾就是空清單——**不預設「沒勾等於全部」**：這一格決定的是真的會送出
        委託的帳戶，把「沒動作」解釋成「全部」是最貴的一種貼心。
        """
        return [sheet for sheet in getattr(self, "order_account_order", [])
                if sheet in self.order_checked]

    def _order_no_account_text(self):
        """
        沒勾帳戶時要說的話。清單根本是空的（Excel 還沒開）跟只是還沒勾，
        要人做的事不一樣。

        「還沒登入」不在這裡講：帳戶清單來自 Excel 分頁，跟登入無關（見
        refresh_order_accounts）；沒登入的話會在真的要送委託那一刻才擋下來。
        """
        if not self._order_known_sheets():
            return "還沒讀到任何帳戶——請先按左上角的「開啟EXCEL」。"
        return "請先在左邊的「執行帳戶」勾至少一位（要全部就按「全選」）。"

    def _on_order_account_click(self, event):
        """
        帳戶清單被點了一下。只有真的點在某一列上才算——點表頭（要排序沒有、
        要調欄寬有）、點欄與欄之間那條分隔線都不該把某一位勾起來：調個欄寬
        順手把一位帳戶勾進這一輪，是不會有人發現的那種錯。
        """
        if self.order_accounts.identify_region(event.x, event.y) not in ("tree", "cell"):
            return
        self._on_order_account_toggled(self.order_accounts.identify_row(event.y))

    def _on_order_account_toggled(self, sheet):
        """
        點了帳戶清單的某一列＝把那一位的勾選翻面（2026/09/02 起可以勾好幾位）。

        點整列都算數，不必剛好點在那個 ☐ 上：那個記號只是 #0 欄文字的第一個
        字元，沒有真正的 checkbox 可以瞄準，硬要人點準一個字寬的目標只是把
        「看起來像 checkbox」變成「用起來不像」。

        鎖住的時候（`_order_locked()`：Excel 正在被動，或這一輪委託還沒跑完）
        直接不理，不是先改再扳回去——勾選是這邊自己存的一份資料（
        self.order_checked），畫面只是它的投影，不像 09/01 那版 Treeview 選取
        會被 Tk 自己改掉，沒有「已經在路上的那一次選取」要追。

        **改勾選＝把讀進來的資料整批清掉**（見 _clear_order_round）：畫面上
        留著的持股、試算、股價是上一批人的，勾進來的新人一格都還沒讀。留著
        不清的話，新勾的那位會照著別人的試算數字算出一列「看起來可以送」的
        委託——不報錯，只會把 A 的數字掛到 B 的帳上（CLAUDE.md 講 _revisit
        那四道身分核對時擔心的就是同一件事，只是那邊是機器搞錯，這裡是畫面
        騙人）。股票清單本身留著：那是「這一輪要動哪幾檔」，跟勾了誰無關，
        每加一位就要重選一次股票只是白工。
        """
        if self._order_locked() or sheet not in self.order_account_rank:
            return
        if sheet in self.order_checked:
            self.order_checked.discard(sheet)
        else:
            self.order_checked.add(sheet)
        self._after_order_accounts_changed()

    def _on_order_accounts_all(self, checked):
        """「全選」／「全不選」：一次把清單上每一位都勾起來或取消。"""
        if self._order_locked():
            return
        self.order_checked = set(self.order_account_rank) if checked else set()
        self._after_order_accounts_changed()

    def _after_order_accounts_changed(self):
        """
        勾選變了之後要做的事，兩個入口（單列、全選）共用：把每一列的 ☐／☑
        重畫，把上一批讀進來的資料清掉。

        **不在這裡觸發「讀取試算」**（2026/09/03 推翻 2026/09/02 那版「勾了就
        自動讀」）：勾帳戶當下往往還沒決定要動哪幾檔股票，這裡就先去讀
        Excel（尤其追價那兩種作業還會順便跑「更新股價」巨集）等於在使用者
        還沒做完設定時就先付一次代價。讀取試算的時機改到 add_order_stock——
        「指定股票」按下「新增」那一刻，見那邊的說明。

        重畫是逐列改文字，不是整份 _fill_order_accounts()——後者會連帶重算欄寬、
        重新排序，而勾選這件事一格報酬率都沒動，清單的內容與順序完全一樣。
        """
        for sheet, label in self.order_account_label.items():
            self.order_accounts.item(sheet, text=self._order_account_text(sheet, label))
        self._clear_order_round()

    def _clear_order_round(self, keep_stocks=True):
        """
        把「上一批人手上的東西」清掉：持股、試算、查回來的報價，也就是每一項
        都跟「是誰」綁在一起的資料。改勾選（_on_order_account_toggled）與換
        Excel 檔（ui_background._forget_round）都走這裡——兩邊要清的是同一份
        東西，列兩份遲早會有一邊漏掉一項。

        `keep_stocks` 決定要不要連「指定股票」那一格一起清掉：

        - 改勾選（預設，留著）：股票清單是「這一輪要動哪幾檔」，跟勾了誰無關
          （某一位沒有那一檔，執行預覽會自己寫「這一位沒有這檔」）。每加一位
          就要重選一次股票只是白工。
        - 換 Excel 檔：連候選清單本身都換了一份，留著等於拿舊檔的股票去對新檔
          的分頁。

        這裡本身不重讀——重讀的時機是 add_order_stock 按「新增」那一刻
        （2026/09/03 起改成這裡，之前是勾帳戶就自動讀，見
        _after_order_accounts_changed 的說明）；換 Excel 檔（_forget_round）
        走同一個函式，那時候 Excel 才剛開，也沒有勾選可言。
        """
        self.order_holdings, self.order_plans, self.order_quotes = {}, {}, {}
        self.order_holding_labels = {}
        self.order_loaded = set()
        if not keep_stocks:
            # 連「讀取ＯＯ持股」讀到的候選也要一起忘掉：換了 Excel 檔，舊檔
            # 第一個分頁的股票清單、分頁名字都對不上新檔了（見
            # refresh_order_stock_list）。按鈕文字跟著退回還沒讀過的樣子。
            self.order_stock_catalog, self.order_stock_list_sheet = [], None
            self._update_order_stock_list_button()
            self.order_prices = {}
            self._rebuild_order_names()   # catalog／holding_labels 都空了，順便清空下拉
            self.order_stock_pick.set("")
            for row in list(self.order_rows):
                row["frame"].destroy()
            self.order_rows = []

        # 這句只是個過場：勾了人的話 _after_order_accounts_changed 緊接著就會
        # 呼叫 refresh_order_plans，狀態列馬上被它的「讀取中…」蓋掉，這裡不用
        # 再講「接著按讀取試算」（2026/09/02 拿掉按鈕後那句話已經不成立）。
        sheets = self._order_sheets()
        if not sheets:
            text = ""
        elif len(sheets) > 3:
            text = f"已勾 {len(sheets)} 位。"
        else:
            text = f"已勾 {'、'.join(sheets)}。"
        self.order_status.configure(text=text)
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
        self._sync_order_clear_controls()
        # 「查詢委買賣」是追價限定功能，跟 order_ticks_entry 同一個道理
        # （disabled 不整個藏起來）。切模式代表股票清單整批清掉重來（下面），
        # 舊查到的報價沒有對象可用，一併清空——不留著一份查不到任何一列在
        # 用的舊資料。
        self.order_quotes = {}
        self._update_order_quotes_ui()
        self._reset_order_stock_rows()

    def _sync_order_clear_controls(self):
        """
        出清作業第二列那幾組控制項（時機、追價檔數、多輪那一列）的亮暗。切時機、
        切單位、切作業三條路共用這一支——三個地方各寫一份的話，之後任何一條規則
        改了都會有一邊沒跟上，而且不會報錯，只是某顆按鈕停在不該有的狀態。

        兩條規則：

        - **零股沒有盤前那一版。** 規劃文件「出清股票－零股」整節只有一組設定，
          而且走的是盤中零股那一場（`order_fill.TAB1_ODD = '5'`，見
          orders.BS_FLAG_ODD）。所以選了零股就把時機固定在盤中、「盤前」那顆
          變灰——那不是「還沒接」，是這個單位底下根本不存在的選項。
        - **追價檔數與多輪是「價格不是人填的」那幾種情況才有意義**（出清整張・
          盤中、出清零股兩種），所以問的是 _order_uses_excel_price，不是時機
          本身。盤前的價格是人一格一格填的，追不追價無從談起。「自動更新股價」
          再多一層：沒勾多輪它不會發生（見 _on_order_multi_round_changed）。
        """
        odd = self._order_clear_odd()
        if odd and self.order_mode.get() != "intraday":
            # 固定成盤中。set() 不會觸發 command，所以 _order_mode_last 要自己
            # 跟上——不然下一次切時機被 busy 擋下來時，那個「還原回上一個值」
            # 的還原點會是一個已經不存在的選擇。
            self.order_mode.set("intraday")
            self._order_mode_last = "intraday"
        for value, radio in self.order_mode_radios.items():
            radio.configure(state="disabled" if odd and value == "pre" else "normal")

        chase = self._order_uses_excel_price()
        self.order_ticks_entry.configure(state="normal" if chase else "disabled")
        self.order_multi_round_check.configure(state="normal" if chase else "disabled")
        if not chase:
            self.order_multi_round.set(False)
            self.order_auto_price.set(False)
            self.order_auto_price_check.configure(state="disabled")

    def _reset_order_stock_rows(self):
        """
        股票清單整批清掉重選，再把跟著它變的東西重畫一遍。切時機、切單位、切
        作業三條路共用（9.3 第 1 點）：每一列的形狀跟著設定走（出清整張有比重、
        出清零股沒有、盤前還多一格價格），與其想辦法把舊的列轉成新形狀，不如
        整批清掉重來。
        """
        for row in list(self.order_rows):
            row["frame"].destroy()
        self.order_rows = []
        self._resize_order_stock_column()
        self._recompute_order_preview()
        # 執行按鈕上寫著「整張・盤中」還是「零股・盤中」（見 _order_exec_label），
        # 要跟著換。
        self._update_order_exec_ui()
        # 「新增」能不能按跟這一輪要不要跑巨集有關（見 _order_excel_buttons），
        # 切設定要重算一次。
        self._apply_busy_state()

    def _on_order_job_changed(self):
        """
        切作業。跟切模式同一條規矩：**股票清單整批清掉重選**（9.3 第 1 點）
        ——「比重」在出清股票是人填的設定，在買賣股票根本不存在（張數與價格
        來自 Excel 的下單試算 M19:N28），沿用舊的列會讓人以為兩邊是同一個
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

        # 「單位」兩個作業各畫一組單選鈕，綁的卻是同一個變數（見
        # ui_layout._build_order_unit）——切過來之後那個值在新作業可能是還沒接
        # 上的那一段（例如全持股交易只有整張），單選鈕是灰的，變數卻還停在上
        # 面：按鈕上的字、送出去的量都會照一個按不到的選項走。跟 _order_intraday
        # 要連作業一起問 order_mode 是同一類的錯——不會報錯，只會做了不該做的事。
        #
        # **要先扳正單位再同步第二列**：_sync_order_clear_controls 問的「現在是
        # 不是出清零股」取決於單位（見 _order_clear_odd），反過來做會拿還沒扳正
        # 的值去決定時機那兩顆的亮暗。
        if not orders.unit_ready(job, self.order_unit.get()):
            self.order_unit.set(orders.UNIT_LOT)

        # 先清空股票清單再改追價檔數：order_ticks.set() 會立刻觸發
        # _recompute_order_preview（見 __init__ 的 trace_add），這時候如果
        # order_rows 還是切作業前的舊列（形狀跟到手的新作業對不上，例如買賣
        # 股票的列沒有 weight_pct），plan_stock_orders 會 KeyError。
        self._reset_order_stock_rows()

        if job == orders.JOB_CLEAR:
            self.order_ticks.set(orders.DEFAULT_TICKS[self.order_unit.get()])
            self._sync_order_clear_controls()

    def _on_order_unit_changed(self):
        """
        切整張／零股。**三種作業一律股票清單整批清掉重選**，執行預覽跟著空掉。

        出清股票本來就是這樣：出清整張的每一列有「比重」，出清零股沒有（規劃
        文件「出清股票－零股」那一節的設定裡就沒有比重——零股是整段賣掉，沒有
        賣幾成可以斟酌），列的形狀不一樣，沿用舊列會留下一個再也不會被讀到的
        輸入框，跟 9.3 第 1 點對切作業的規矩是同一條。

        買賣股票原本不清——理由是它只是同一個試算股數換送另一半（見
        orders.plan_trade_orders），列的形狀沒變，重算預覽就夠了。2026/09/04
        使用者確認**三種作業統一照出清那條走**：切單位就是「這一輪整套設定重來」
        這個意思，一個作業一套規矩的話，人得先想起自己站在哪一邊才知道清單會不會
        留著。代價是切完要重按「新增」（那也會順便重讀一次下單試算，見
        add_order_stock），使用者確認接受。

        全持股交易（JOB_FULL）現在只開放整張（orders.UNITS_READY），實務上切不到
        單位，但規則一樣套在它身上——哪天零股接上了不必回來補這一條。

        切時機（_on_order_mode_changed）、切作業（_on_order_job_changed）也都是
        無條件清空，三條路現在講的是同一句話。

        追價檔數跟著換成那個單位的預設值（整張 2 檔、零股 3 檔，規劃文件各自寫在
        自己那一節）；這是出清限定的收尾，另外兩個作業根本沒有這個欄位。使用者
        自己改過的數字一樣會被蓋掉，跟股票清單被清掉是同一種取捨：留一半舊的比
        全部重設更難察覺。
        """
        # 三種作業共同的那一步，擺在分支之前——這就是「統一作法」本身，不是
        # 剛好三條路各自都寫了一次。順序也不能反過來：下面 order_ticks.set()
        # 會立刻觸發 _recompute_order_preview（見 __init__ 的 trace_add），
        # 這時候 order_rows 要是還留著切單位前的舊列形狀就會 KeyError
        # （同 _on_order_job_changed 的說明）。
        self._reset_order_stock_rows()

        if self.order_job.get() != orders.JOB_CLEAR:
            # 買賣股票／全持股交易沒有追價檔數，也用不到即時報價
            # （見 _order_quotes_available），清單清掉就結束了。
            return

        self.order_ticks.set(orders.DEFAULT_TICKS[self.order_unit.get()])
        self._sync_order_clear_controls()
        # 舊的即時報價跟著作廢，理由同切時機：清單都清空了，沒有任何一列在用它。
        self.order_quotes = {}
        self._update_order_quotes_ui()

    def _order_intraday(self):
        """
        現在是不是「出清股票・**整張**・盤中」。

        盤前／盤中是**出清作業自己的設定**（9.3 第 4 點），所以問「是不是盤中」
        一定要連作業一起問：從出清・盤中切到買賣股票的時候 order_mode 還留著
        "intraday"，只看它的話，買賣股票會莫名其妙跑去追價、跑去觸發更新股價
        巨集。這種錯不會報錯，只會做了一堆不該做的事。

        2026/09/03 起還要連單位一起問：出清・零股接上之後，它的時機也是盤中
        （而且被固定成盤中），但走的是完全不同的一條路——比重、IOC、收斂條件
        全都不一樣（見 _order_clear_odd）。這一支從此專指整張那一版，「兩種
        出清都算」的那個問題改問 _order_uses_excel_price。
        """
        return (self.order_job.get() == orders.JOB_CLEAR
                and self.order_mode.get() == "intraday"
                and self.order_unit.get() == orders.UNIT_LOT)

    def _order_clear_odd(self):
        """
        現在是不是「出清股票・零股」（規劃文件「出清股票－零股」那一節）。

        跟 _order_intraday 一樣要連作業一起問，不能只看單位：零股在買賣股票是
        「照下單試算的股數送出去」那半段，在出清股票才是「全部掛賣單、幾秒後
        全部取消」這一整套流程——兩邊差的不是幾個欄位，是整條執行路徑（見
        orders.UNITS_READY 的說明）。
        """
        return (self.order_job.get() == orders.JOB_CLEAR
                and self.order_unit.get() == orders.UNIT_ODD)

    def _order_uses_excel_price(self):
        """
        這一輪的委託價要不要以 Excel I 欄的成交價為基準追價（見
        orders.chase_price）。出清整張・盤中與出清零股都算：兩者的價格都不是人
        填的，所以都要在「新增」股票時順便觸發一次「更新股價」巨集、都用得到
        「查詢委買賣」把即時委買一先查回來。

        跟 _order_intraday 分成兩支而不是把零股塞進去：那一支問的是「時機是不是
        盤中」，而零股根本沒有盤前那一版（時機被固定成盤中，見
        _sync_order_mode_for_unit）。混在一起的話，「盤前／盤中」這個設定會同時
        代表兩件事，之後每一個問它的地方都要自己再分一次。
        """
        return self._order_intraday() or self._order_clear_odd()

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
        切「半自動」／「自動送出」。自動送出是這裡的預設（2026/09/02 使用者改，
        半自動曾經是預設值、也是最早實測過整條路能通的模式，見記憶
        order-exec-sequential-wired-up）。

        勾選當下不再跳確認框（2026/09/03 使用者要求拿掉）——「開始下單」那顆
        確認框本來就會把目前是自動還是手動模式紅字標出來（見
        ui_order_exec._start_order_batch 的 emphasize 那段），真正送出委託之前
        還是會看到、還是要按一次確認，不用兩顆確認框講同一件事。
        """
        if self.busy:
            self.order_auto_confirm.set(self._order_auto_last)
            show_info(self.root, "忙碌中", "現在有背景工作在跑，先等它結束才能切換。")
            return

        self._order_auto_last = self.order_auto_confirm.get()

    def _order_excel_buttons(self):
        """
        下單分頁裡「按下去會用 COM 動 Excel」的按鈕：只要有任何一條路正在動那份
        活頁簿（或那份活頁簿根本沒開著）就變灰（見 ui_background._apply_busy_state，
        它負責在 _excel_in_use() 那些旗標或 excel_open 變動時呼叫這裡）。

        原本「讀取持股」（現拆成「讀取ＯＯ持股」與按「新增」時觸發的「讀取
        試算」，2026/09/03 起入口從勾帳戶改成「新增」）與「新增」各自只看自己
        那一個旗標——前者跑著的時候後者還是亮的，而兩條路都會一頁一頁 Activate
        再跑巨集，交錯之後巨集會跑在別人剛切過去的那一頁上（見
        excel_io._EXCEL_LOCK 的說明）。

        擋住而不是排隊：跟 refresh_order_plans 開頭那道 `_excel_in_use()` guard
        同一個態度——按下去卻剛好撞上忙碌就默默不做事，不是排進佇列，下一次
        「新增」還會再有機會補上。
        """
        busy = self._excel_in_use()
        # 「讀取ＯＯ持股」還要 Excel 真的開著才亮：這顆做的事整個就是讀那份
        # 活頁簿，沒開著根本無事可做——跟更新分頁的「更新全部帳戶」同一個規矩
        # （見 ui_sync._sync_buttons）。2026/08/31 之前這裡只看忙碌旗標，所以
        # Excel 沒開的時候「更新」那顆是灰的、這顆卻亮著，按下去換來一個
        # 「Excel 沒開著」的視窗；現在一起灰，那個視窗也跟著拿掉了（見
        # ui_background._require_excel）。「讀取試算」2026/09/02 拿掉按鈕之後
        # 不必再管它的灰／亮，勾帳戶那條路自己會被 _order_locked() 擋住。
        state = "normal" if self.excel_open and not busy else "disabled"
        self.order_stock_list_button.configure(state=state)
        # 「新增」是否要跟著變灰，現在看**有沒有勾帳戶**而不是看作業種類：
        # 2026/09/03 起不管哪種作業，「新增」都會在勾了帳戶時附帶呼叫
        # refresh_order_plans（見 add_order_stock 的說明），所以只要 Excel
        # 忙著、而且這一按會真的去讀，才需要擋。沒勾帳戶的話 add_order_stock
        # 根本不會碰 COM，跟以前「盤前完全不碰 COM，沒有理由跟著變灰」是同一個
        # 判斷，只是條件換成勾選狀態。它也不跟著 excel_open 走：Excel 沒開的話
        # refresh_order_plans 自己會被 _require_excel() 擋下，不必特地在這裡
        # 先擋一次。
        add_busy = busy and bool(self._order_sheets())
        self.order_add_button.configure(state="disabled" if add_busy else "normal")
        # 「執行帳戶」在有事情在跑的時候鎖住。改勾選會把持股、試算整批清掉
        # （見 _on_order_account_toggled），而背景那一趟讀回來的是**改之前**
        # 那批人的資料，兩件事撞在一起的結果是清單上勾著 B、手上的試算卻是 A
        # 的——不會報錯，只會拿 A 的數字算 B 的張數。依序執行中一樣鎖住：queue
        # 在按下「開始下單」那一刻就凍結了，改勾選動不到它，但畫面會在跑到一半
        # 的時候被清空，看起來像壞掉。
        #
        # Treeview 的 disabled 是用 state() 給的，不是 configure(state=...)，
        # 而且**擋不住 <Button-1>**（那是綁在元件上的事件，不是內建的選取行為）
        # ——它只負責「看得出來現在不能動」，真正擋下來的是
        # _on_order_account_toggled 開頭那道 _order_locked()。
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

        按下「新增」之後，勾了帳戶的話就重讀一次「試算」（見
        refresh_order_plans）——2026/09/03 起這裡才是唯一的觸發點，勾帳戶
        本身不再自動讀（見 _after_order_accounts_changed 的說明）。每按一次
        「新增」都重讀，不管這一批帳戶是不是已經讀過：買賣股票只讀
        D~F／B22／M19:N28，成本很低；出清整張・盤中與出清零股會多跑一次
        「更新股價」巨集，但那本來就是追價要用的成交價基準，加了新股票就該
        用最新的（使用者 2026/09/03 確認接受這個代價，不必只在「還沒讀過」時
        才讀）。剛加的這一檔如果原本沒被最近一次試算涵蓋到（例如本來沒
        持股），這一步會補上。

        這次重讀特意放在「已經加過了」那個重複檢查**之前**：勾帳戶不會自動
        讀之後，「重按一次新增」是使用者唯一能自己觸發重讀的方式——如果
        因為股票已經在清單裡就整段跳過，剛勾的新帳戶會一直卡在讀不到試算，
        而使用者手上所有能按的動作都只換來一句「已經加過了」，沒有任何
        辦法補救。
        """
        raw = self.order_stock_pick.get().strip()
        if not raw:
            return
        code = raw.split(" ")[0].strip().upper()
        if self._order_sheets():
            self.refresh_order_plans()
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
        elif self._order_clear_odd():
            # 出清零股也不填比重：規劃文件「出清股票－零股」的設定裡沒有這一項，
            # 送出去的量固定是持股的零股那一段（見 orders.plan_clear_odd_orders）。
            # 留一個永遠不會被讀的輸入框比沒有更糟——人填了數字卻不影響任何結果。
            row.pop("weight", None)
        elif self.order_mode.get() == "pre":
            row["price"] = tk.StringVar()
            row["price"].trace_add("write", lambda *_a: self._recompute_order_preview())
        self._build_order_stock_row(row)
        self.order_rows.append(row)
        self.order_stock_pick.set("")
        self._resize_order_stock_column()
        self._recompute_order_preview()

    @staticmethod
    def _order_weight_key_ok(value):
        """
        比重輸入框的按鍵驗證（Entry 的 validatecommand，value 是那一鍵按下去之後
        會變成的完整字串）：放行空字串（打到一半、或刪光重打）跟 0～100 之間的
        數字，其餘一律擋下，包含整段貼上——貼「150」這種一次到位的輸入也一樣會
        被擋，不會先貼進去再讓後面的重算邏輯發現算出來的張數不合理。
        """
        if value == "":
            return True
        try:
            num = float(value)
        except ValueError:
            return False
        return 0 <= num <= 100

    @staticmethod
    def _order_price_key_ok(value):
        """
        出清股票・整張・盤前那格價格輸入框的按鍵驗證，跟 _order_weight_key_ok
        同一個做法：放行空字串（打到一半、刪光重打），只放行非負的數字，其餘
        （中文、負號、整段貼上非數字的內容）一律在按鍵層級擋下，不是送出前
        才靠 REASON_NO_PRICE 那句提醒去發現填錯。
        """
        if value == "":
            return True
        try:
            num = float(value)
        except ValueError:
            return False
        return num >= 0

    def _build_order_stock_row(self, row):
        """
        一檔股票一列，**一行排完**：買賣別、股票、比重、價格（或試算／
        Excel 股價）、移除。

        2026/08/31 之前是分兩行的，那是為了「指定股票」還在 300 像素寬的左欄時
        設計的——一行塞不下「名稱＋比重＋價格＋移除」，價格輸入框會被擠到剩沒
        幾個像素、打不進去字。那一格搬到右欄之後有一千多像素寬，前提不成立了，
        而且分兩行等於把橫向那一大片空白換成縱向的浪費：一檔 70 像素降到 39，
        省下來的高度全部流給下面的執行預覽（它是「帳戶 × 股票」的交叉，永遠
        不夠——見 ui_layout._build_order_tab）。

        每一列自己是一個 Frame（用 pack 疊起來，不是整個清單一張大 grid）：
        移除中間一列時不會留下空位，不必自己重新排列剩下的列。列**內部**才用
        grid，而且每一列的欄 minsize 都用同一組 ORDER_STOCK_COL_W，所以三檔的
        「比重」「價格」「移除」會上下對齊成一張小表——不是靠股票名稱剛好一樣長。

        買賣別跟著這一輪的方向走（見 _order_init_state 的 order_side：9.3 之後
        方向是作業算出來的，不是人選的），一整批
        股票共用同一個方向，加進來那一刻就定案——切買賣的時候整批清單會被
        清空重選（見 _on_order_job_changed），不會出現舊列還留著舊方向的
        情況。底色跟網站本身買紅賣綠的配色一致（Sell.TLabel／Buy.TLabel 在
        ui_layout._build() 裡註冊），不必看文字就認得出方向。

        `"price" in row` 決定要不要畫價格輸入框——盤中模式的 row 沒有這個
        key（見 add_order_stock），不是留白也不是畫一個不會被讀的欄位。

        盤中模式沒有價格輸入框的位置改顯示 Excel 讀回來的股價（self.
        order_prices，「讀取試算」讀回來的那一份，跟 order_names 不是同一份
        ——後者還併了「讀取ＯＯ持股」的候選，見 _rebuild_order_names）——
        這不只是給人參考，開始下單那一刻會拿 order_prices 當第一輪
        chase_price 的 pricenow（見 start_order_execution），追價檔數還是
        要在下單前用這個基準再算一次邊界、查一次對手方第一檔（見
        orders.chase_price），這裡顯示的數字就是實際會拿去算價的那一個
        （不是另一條網頁現查的路，2026/08/29 使用者確認拿掉了）。這裡存了
        Label 物件本體（row["price_label"]）而不是只畫一次文字，因為每次按
        「新增」都會觸發一次背景重讀（2026/09/03 起併進 refresh_order_plans，
        見 _on_order_plans_data 補畫 price_label 那一段），回來要能就地更新
        這一列的文字，不是只有新加的那一列，而是畫面上全部盤中列一起刷新。
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

        if trade:
            # 逐帳戶的試算股數不在這一格重複講了（2026/09/02 拿掉「N 位有
            # 試算」那句）：一位一列才看得到真正的買賣與股數，右邊執行預覽
            # 那張表已經是唯一的答案，這裡再摘要一次只是同一件事講兩次。
            # 這個模式沒有比重／價格，「移除」緊接在股票後面就好，不必為了
            # 對齊硬跳到最後一欄。
            ttk.Button(block, text="移除", bootstyle="danger-outline",
                      command=lambda: self.remove_order_stock(row)).grid(
                row=0, column=2, sticky="w", padx=(6, 0))
            row["frame"] = block
            return

        # 比重、價格各自包一個小 Frame 再放進格子裡：label＋entry＋單位是一組
        # 三件套，讓它們在組內用 pack 貼在一起，組跟組之間才靠 grid 的欄對齊。
        #
        # `"weight" in row` 決定畫不畫比重那一組（出清零股沒有，見
        # add_order_stock），跟下面 `"price" in row` 是同一種做法：這一列有哪
        # 幾個設定是加進來那一刻就定案的，不是在這裡再判斷一次作業／單位——
        # 兩邊各判一次遲早會分岔，而分岔的結果是畫得出來卻讀不到的輸入框。
        weight_box = ttk.Frame(block)
        weight_box.grid(row=0, column=2, sticky="w", padx=(12, 0))
        if "weight" in row:
            ttk.Label(weight_box, text="比重").pack(side="left")
            # 比重是「持股 × 這個百分比」（orders.lots_from_weight），沒有上限的話
            # 打錯一個 0（例如 150）會算出比實際持股還多的張數，出清股票時可能因此
            # 送出一張比帳上還多的委託——按鍵層級擋掉範圍外的輸入，不是送出前才報錯。
            if not hasattr(self, "_order_weight_vcmd"):
                self._order_weight_vcmd = (self.root.register(self._order_weight_key_ok), "%P")
            ttk.Entry(weight_box, textvariable=row["weight"], width=6,
                     font=(self.family, FONT_SIZE),
                     validate="key", validatecommand=self._order_weight_vcmd).pack(side="left", padx=(4, 0))
            ttk.Label(weight_box, text="%").pack(side="left", padx=(2, 0))
        else:
            # 出清零股：量是算出來的（持股的零股那一段，1~999 股），不是人填的。
            # 這一格空著不寫字的話，三檔上下對齊的那張小表會缺一塊，看起來像
            # 沒載入完；寫一句話同時把「為什麼這裡不能填」講掉。
            ttk.Label(weight_box, text="全部零股", style="Hint.TLabel").pack(side="left")

        price_box = ttk.Frame(block)
        price_box.grid(row=0, column=3, sticky="w", padx=(12, 0))
        if "price" in row:
            ttk.Label(price_box, text="價格").pack(side="left")
            if not hasattr(self, "_order_price_vcmd"):
                self._order_price_vcmd = (self.root.register(self._order_price_key_ok), "%P")
            ttk.Entry(price_box, textvariable=row["price"], width=8,
                     font=(self.family, FONT_SIZE),
                     validate="key", validatecommand=self._order_price_vcmd).pack(side="left", padx=(4, 0))
            ttk.Label(price_box, text="元").pack(side="left", padx=(2, 0))
        else:
            excel_price = self.order_prices.get(row["code"])
            price_text = f"Excel股價 {show(excel_price)} 元" if excel_price is not None else "Excel股價：讀不到"
            label = ttk.Label(price_box, text=price_text, style="Hint.TLabel")
            label.pack(side="left")
            row["price_label"] = label

        ttk.Button(block, text="移除", bootstyle="danger-outline",
                  command=lambda: self.remove_order_stock(row)).grid(
            row=0, column=4, sticky="w", padx=(12, 0))
        row["frame"] = block

    def remove_order_stock(self, row):
        row["frame"].destroy()
        self.order_rows.remove(row)
        self._resize_order_stock_column()
        self._recompute_order_preview()

    # ---------- 執行預覽 ----------

    def _order_execution_accounts(self):
        """
        這一輪要動誰——勾起來的那幾位，照清單上的順序（＝報酬率由低到高，
        規格「報酬率低的先執行」）排好；沒勾就是空清單。

        形狀跟 orders.order_accounts 回傳的一樣（多一個 "order"），下面三支
        plan_* 照吃不必改。"order" 直接用清單上那一列的號碼，不重新從 1 數起：
        那個號碼是「報酬率由低到高的第幾位」，跳著勾的時候重新編號會讓執行
        預覽的「順序」欄跟左邊那張清單對不起來。

        報酬率讀不到（None）的也照樣回傳，不像 orders.order_accounts 那樣踢進
        skipped：那支的顧慮是「用猜的決定誰先執行」，而這裡是使用者自己一位
        一位勾的，勾了就是要跑，他們排在最後、號碼寫「－」（見
        _fill_order_accounts）。
        """
        return [{"sheet": sheet, "return_rate": self.order_return_rates.get(sheet),
                 "order": self.order_account_rank.get(sheet, "－")}
                for sheet in self._order_sheets()]

    def _order_stock_settings(self):
        """
        把畫面上股票清單的目前輸入值（比重，盤前模式再加價格）整理成
        orders.plan_stock_orders／plan_intraday_orders 吃的格式。給
        _recompute_order_preview 跟 start_order_execution 共用——後者要用
        「跟畫面上一模一樣」的設定去組執行清單，不是自己另外算一次。
        """
        stock_settings = []
        for row in self.order_rows:
            setting = {"code": row["code"], "name": row["name"]}
            # 出清零股那一列沒有比重（見 add_order_stock），連 weight_pct 這個鍵
            # 都不給——plan_clear_odd_orders 本來就不看它，補一個 0 進去只會讓
            # 「這個作業沒有比重」變成「比重是 0」，兩者在別的 plan_* 底下差很多。
            if "weight" in row:
                try:
                    setting["weight_pct"] = float(row["weight"].get())
                except ValueError:
                    setting["weight_pct"] = 0
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
            ready = "、".join(orders.JOB_NAMES[job] for job in orders.JOBS_READY)
            self._render_order_preview([], [
                f"「{orders.JOB_NAMES[self.order_job.get()]}」還沒接上，"
                f"目前可以執行的是「{ready}」"
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
                ordered, self.order_plans, self.order_holdings, self.order_unit.get(),
                loaded_sheets=self.order_loaded)
            # 勾了但還沒讀的那幾位：他們的每一列都會寫「還沒讀取試算，略過」，
            # 但那是一列一列講的，整批漏讀（最常見的情況：勾完帳戶、加了股票，
            # 但那一批試算剛好撞上 Excel 忙碌而被 refresh_order_plans 靜靜跳過，
            # 見 add_order_stock 的說明）要在這裡講一次。2026/09/03 起讀取試算
            # 的入口改成「新增」，「重新勾選」已經不會觸發重讀，remedy 改成
            # 「按新增」——哪怕股票已經在清單裡也一樣會重讀（見 add_order_stock
            # 把重讀擺在「已經加過了」檢查之前的理由）。
            missing = [account["sheet"] for account in ordered
                       if account["sheet"] not in self.order_loaded]
            if codes and missing:
                who = "、".join(missing) if len(missing) <= 3 else f"{len(missing)} 位"
                hints.append(f"⚠ {who}還沒讀到下單試算，讀取中或請按「新增」重新觸發。")
        elif self._order_clear_odd():
            # 出清零股：量是持股的零股那一段，沒有比重可讀；價格跟盤中出清同一條
            # 追價路（見 orders.plan_clear_odd_orders）。
            ticks = self._order_ticks_setting()
            if ticks is None:
                preview = []
                hints.append("⚠ 追價檔數要填 0 以上的整數。")
            else:
                preview = orders.plan_clear_odd_orders(
                    self._order_stock_settings(), ordered, self.order_holdings, ticks,
                    prices=self.order_prices, quotes=self.order_quotes)
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
        self.order_preview_iid = {}

        # 「出清進度」欄只有出清股票這個作業才有意義（見 orders.PROGRESS_NONE
        # 的說明），問 preview 每一列自己帶的 "clearing"（三支 plan_* 只有出清
        # 那三支會標）而不是直接問 self.order_job：多輪出清重畫「下一輪還剩
        # 什麼」時（見 _on_order_price_refresh）只有 preview 拿得到，畫面上的
        # order_job 隨時可能已經被切到別的作業。preview 是空的（沒有帳戶、還
        # 沒選股票）就沒有列可以問，退回問 order_job。
        showing_progress = (any(item.get("clearing") for item in preview) if preview
                            else self.order_job.get() == orders.JOB_CLEAR)
        base_columns = self.order_preview_columns
        self.order_preview["displaycolumns"] = (
            base_columns if showing_progress
            else tuple(key for key in base_columns if key != "progress"))

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
            if item["price"] is None and item.get("chase"):
                # 追價那兩條路（出清整張・盤中、出清零股）才會「現在還沒有價格，
                # 下單前才算」——買賣股票的價格是 Excel 試算給的，讀不到就是這
                # 一列根本不會送（沒這一檔、試算是空的），寫「依 Excel 成交價
                # 追價」會讓人以為它還會去查。問 "chase" 不問 bs_flag：零股同樣
                # 是追價來的，委託別卻是 ROD（見 orders.BS_FLAG_ODD）。
                price_text = PRICE_PENDING_TEXT
            elif item["price"] is None:
                price_text = "－"
            elif isinstance(item["price"], str):
                price_text = item["price"]
            else:
                price_text = show(item["price"])

            if item.get("clearing"):
                unit = item.get("unit", orders.UNIT_LOT)
                # 基準拿凍結好的快照（_snapshot_order_base_qty），但只在「這一批
                # 還在跑」的時候信它（order_exec_active）——這支同時被兩條路呼叫：
                # 多輪重讀時（_on_order_price_refresh）呼叫這裡的時候 active 保證
                # 是真的（那支開頭就會在 active 是假的時候提早 return）；使用者
                # 還沒按「開始下單」、單純在調整設定的那條路（_recompute_order_
                # preview）active 是假的，這時候即使上一批（可能是別的單位／別的
                # 持股水位）留下的快照剛好對得到同一個 (帳戶,股票)，也不能拿來用
                # ——那是上一批的分母，不是這一批的。停止之後要「留著讓人看最後
                # 結果」是 _refresh_order_progress_cells 單格更新在做的事，不靠
                # 這裡的 fallback，所以這裡收斂成只問 active 不會漏掉那個需求。
                base_qty = (self.order_exec_base_qty.get((item["sheet"], item["code"]))
                           if self.order_exec_active else None)
                if base_qty is None:
                    # 還沒開始執行（或快照跟這一批對不上）→ 用這一列當下的量當
                    # 基準，分子減分母是 0，畫出來是「條全空 0%」，不是「－」，
                    # 因為這一段真的有東西可以清，只是還沒開始清。
                    base_qty = orders.clearable_qty(item["held_qty"], unit)
                now_qty = orders.clearable_qty(item["held_qty"], unit)
                sent_qty = self._order_sent_qty(item["sheet"], item["code"])
                progress_cell = orders.progress_text(base_qty, now_qty, sent_qty)
            else:
                progress_cell = ""

            # 順序要跟 ui_layout._build_order_preview 的 columns 一模一樣：
            # 順序／帳戶／買賣／股票／張數／價格／持股／出清進度／備註。
            # Treeview 是照位置對欄位的，換了順序卻沒改這裡不會報錯，只是每一
            # 欄顯示別欄的值。
            #
            # 「張數」欄：買賣股票選零股時要把「另有幾張沒送」寫出來（9.4），
            # 選整張只顯示張數；其他作業沒有這個問題，就是一個數字（沒有
            # lots_text 就退回 lots）。持股最小單位是 1 股，不需要小數點；股數
            # 本來就可能上看百萬，千分位才看得出位數（util.show 是全專案統一
            # 用的數字顯示格式）。
            iid = self.order_preview.insert("", "end", values=(
                item["order"], item["sheet"], side_names.get(item["side"], item["side"]),
                f"{item['code']} {item['name']}",
                item.get("lots_text") or item["lots"], price_text, show(item["held_qty"]),
                progress_cell, item["note"],
            ), tags=(tag,) if tag else ())
            self.order_preview_iid[(item["sheet"], item["code"])] = iid

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
        「帳戶」欄寬跟著這次重讀到的帳戶名單重量一次——
        只在 _fill_order_accounts 換了一批名單時呼叫，理由跟
        _resize_order_stock_column 一樣：名單只在讀取的當下換一批，不會因為
        使用者操作畫面上其他東西（換帳戶、改比重）而變動。
        """
        self.order_preview.column(
            "sheet", width=col_width(self.family, list(self.order_return_rates), minimum=wide(90)))

    # ---------- 查詢委買賣（盤中限定） ----------
    #
    # 「查詢委買賣」按鈕：先幫目前清單裡的股票查一次即時委買賣一，讓執行預覽
    # 直接顯示 orders.chase_price 算出來的實際價格，不用等「開始下單」依序跑到
    # 那一筆才臨時查（2026/08/29 使用者要求：出清股票時想在按下去之前就看到會
    # 發生什麼事）。
    #
    # 2026/09/04 起資料來源從 fastquote 的 WebSocket 換成 stockinfo 的 HTTP
    # （見 stockinfo.py 模組說明）。連帶三件事跟著變，都是刻意的：
    #   1. 不用登入、不碰瀏覽器，所以不再借帳戶、不再進瀏覽器背景執行緒的
    #      queue，改成自己開一條 thread（跟 _order_rates_worker 同一種寫法）。
    #   2. 不再佔用 self.busy——查報價不會跟更新分頁或下單搶瀏覽器，沒有理由
    #      在這段期間鎖住「登入／更新／全部登出」。
    #   3. 一檔一個 GET。stockinfo.quote 刻意不做批次（理由見它的模組說明：
    #      零股不能批次，而且整股零股混查會靜靜降級成整股）。

    def fetch_order_quotes(self):
        """
        觸發背景查詢；結果回來見 _on_order_quotes_fetched。

        報價是公開資料，不因帳戶而不同，所以這裡完全不看帳戶——**沒登入也能
        按**。以前要借一組登入過的帳戶去開 FastQuote 彈出視窗，那個限制是舊
        資料來源的技術債，換成 HTTP 之後沒有存在的理由了。

        `self.busy` 仍然擋著（見 _update_order_quotes_ui 的按鈕狀態）：不是
        技術上不行，是「下單依序執行到一半」重查會改掉畫面上的執行預覽價格，
        而正在跑的那一輪用的是開始下單那刻凍結的 order_exec_quotes，兩邊對不
        起來，人看了會誤會（2026/09/04 使用者確認保留這一條）。
        """
        if self.busy or self.order_quotes_busy:
            return
        if not self._order_quotes_available():
            return

        codes = sorted({row["code"] for row in self.order_rows})
        if not codes:
            show_info(self.root, "還沒有股票", "請先加入至少一檔股票。")
            return

        self._order_quotes_requested = codes
        self.order_quotes_busy = True
        self._update_order_quotes_ui()
        # 這一趟查整股還是零股，跟著現在選的作業走——「出清零股」送出去的是零股
        # 委託，要比的就是零股那本簿子（見 stockinfo.quote 的 odd 參數）。整份
        # 預覽同一時間只會有一種，理由見 self.order_quotes 那段說明。
        threading.Thread(target=self._order_quotes_worker,
                         args=(codes, self._order_clear_odd()), daemon=True).start()

    def _order_quotes_worker(self, codes, odd):
        """
        背景執行緒：一檔一個 HTTP GET 查即時委買賣一（見 stockinfo.quote）。

        查不到的代號就不會出現在回傳的字典裡，不是塞一個 None 佔位，呼叫端
        （_on_order_quotes_fetched）自己比對哪些代號漏了——這是換資料來源之前
        就有的約定，沒有變。

        「這一檔查不到」（收盤後、盤中零股第一盤之前、代號不存在）跟「整條路
        壞了」（連不上、逾時）要分開：前者是正常情況，跳過那一檔繼續查下一檔；
        後者每一檔都會壞，記下第一個錯誤就整趟結束，讓畫面說得出是連線問題，
        不要讓人以為是這幾檔剛好沒行情。
        """
        quotes = {}
        for code in codes:
            try:
                quote = stockinfo.quote(code, odd=odd)
            except Exception as exc:
                self.queue.put(("order_quotes_fetched",
                                {"error": str(exc), "hint": QUOTES_OFFLINE_HINT}))
                return
            if quote:
                quotes[code] = quote
        self.queue.put(("order_quotes_fetched", {"quotes": quotes}))

    def _on_order_quotes_fetched(self, payload):
        """
        fetch_order_quotes 的背景回話。查到的併進 self.order_quotes（不是
        整份換掉——重複按「查詢委買賣」，這次沒查到的代號還留著上次查到的
        舊值，比整份清空更安全，見下面 missing 那段的說明），再重算一次
        執行預覽讓畫面反映最新算出來的價格。

        不動 browser_waiting／_set_busy：2026/09/04 起這一趟走自己的 thread、
        不碰瀏覽器也不佔 self.busy（見 fetch_order_quotes）。這裡要是還留著
        `_set_busy(False)`，等於幫別人（更新分頁、下單）把 busy 清掉。
        """
        self.order_quotes_busy = False
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
        「查詢委買賣」現在有沒有意義：只有價格是追價算出來的那兩條路用得到即時
        委買賣一（出清整張・盤中、出清零股）。盤前的價格是人一格一格填的，買賣
        股票的價格來自 Excel 試算，查回來的報價沒有任何一列會用到。
        """
        return self._order_uses_excel_price()

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
