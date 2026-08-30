"""
下單分頁「按下開始下單之前」的那一半：收使用者的輸入（選股票、填比重、勾
帳戶）、讀 Excel（持股、B17 報酬率、成交價）、查即時委買賣一，算出執行預覽。

按下去之後的另一半在 ui_order_exec.py——那裡從凍結這一輪的設定開始，一筆
一筆送出、跑完一輪再決定要不要接下一輪。界線是「這一輪要送什麼」（這裡）
跟「怎麼把它送出去」（那裡），見那個檔案開頭的說明。

「盤前」（股票／比重／價格設定）跟「盤中」（股票／比重／追價檔數設定，價格
用 Excel 成交價＋下單前查對手方第一檔算出來的）共用同一套勾帳戶、算執行
預覽、依序執行的機制，只有「股票設定要填什麼」「怎麼組出執行清單」不一樣
（見 `_on_order_mode_changed`／ui_order_exec 的 `start_order_execution`）。

比重→張數、帳戶依 B17 報酬率排序、組出預覽清單、追價檔數換算價格，全部是
orders.py 的純函式，這裡只負責收輸入、讀 Excel（含成交價，盤中新增股票／
重新整理時順便觸發「更新股價」巨集）、查即時對手方第一檔、把結果畫出來。
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

        # 作業（見 docs/介面規劃.md 9.2／9.3）：買賣股票／出清股票／全持股交易
        # 三選一，是這個分頁最上層的選擇。第二列跟著整列換成「那個作業自己的
        # 設定」，右半邊三個作業共用同一份 widget。
        self.order_job = tk.StringVar(value=orders.JOB_CLEAR)
        self._order_job_last = orders.JOB_CLEAR

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
            messagebox.showinfo("忙碌中", "現在有背景工作在跑，先等它結束才能切換作業。",
                                parent=self.root)
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

        for row in list(self.order_rows):
            row["frame"].destroy()
        self.order_rows = []
        self._resize_order_stock_column()
        self._recompute_order_preview()
        self._update_order_exec_ui()
        self._apply_busy_state()

    def _on_order_unit_changed(self):
        """切整張／零股。目前只影響執行按鈕上的字（零股還沒實作，選不到）。"""
        self._update_order_exec_ui()

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

        買賣別跟著這一輪的方向走（見 _order_init_state 的 order_side：9.3 之後
        方向是作業算出來的，不是人選的），一整批
        股票共用同一個方向，加進來那一刻就定案——切買賣的時候整批清單會被
        清空重選（見 _on_order_job_changed），不會出現舊列還留著舊方向的
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
        if not self._order_job_ready():
            self._render_order_preview([], [
                f"「{orders.JOB_NAMES[self.order_job.get()]}」還沒接上，"
                f"目前只有「{orders.JOB_NAMES[orders.JOB_CLEAR]}」可以執行"
                f"（落地順序見 docs/介面規劃.md 9.7）。"])
            return

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
        if not self._order_quotes_available():
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

    def _order_quotes_available(self):
        """
        「查詢委買賣」現在有沒有意義：只有「出清股票」有盤前／盤中這個設定，
        而追價比價是盤中限定的（9.3 把盤前／盤中降級成出清作業自己的設定）。
        """
        return (self.order_job.get() == orders.JOB_CLEAR
                and self.order_mode.get() == "intraday")

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
