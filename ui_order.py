"""
下單分頁的行為：「盤前」（股票／比重／價格設定）跟「盤中」（股票／比重／
追價檔數設定，價格是下單前現查成交價算出來的）共用同一套勾帳戶、算執行
預覽、依序半自動填單的機制，只有「股票設定要填什麼」「怎麼組出執行清單」
不一樣（見 `_on_order_mode_changed`／`start_order_execution`）。

半自動填單只做到「開出委託確認視窗」為止，不會按裡面的「確認」——那一步
要送出真實委託，留給人自己決定，見 order_fill.py／`_order_fill_job`。

比重→張數、帳戶依 B17 報酬率排序、組出預覽清單、追價檔數換算價格，全部是
orders.py 的純函式，這裡只負責收使用者的輸入、讀 Excel、查即時成交價、
操作瀏覽器、把結果畫出來。
"""

import threading
import tkinter as tk
from tkinter import messagebox

import ttkbootstrap as ttk
from playwright.sync_api import Error as PlaywrightError

import excel_io
import fetch as fetch_mod
import order_fill
import orders
from ui_common import FONT_SIZE, ask_confirm
from util import show


class UiOrderMixin:
    # ---------- 下單分頁：盤前模式 ----------

    def _order_init_state(self):
        """SyncApp.__init__ 呼叫一次。"""
        self.order_rows = []              # 這一輪加進來的股票設定列（見 add_order_stock）
        self.order_holdings = {}          # (分頁名, 股票代號) -> 股數，「重新整理」才會更新
        self.order_names = {}             # 股票代號 -> 名稱，畫面顯示用
        self.order_return_rates = {}      # 分頁名 -> B17 報酬率或 None（讀不到）
        self.order_account_vars = {}      # 分頁名 -> 勾選狀態的 BooleanVar（見 _fill_order_accounts）
        self.order_busy = False

        # 「盤前」／「盤中」模式（見 ui_layout._build_order_tab）。追價檔數
        # 是盤中模式整批共用的一個值（使用者 2026/08/28 確認過，不是像比重
        # 那樣每檔股票各自設定），trace 讓改了立刻反映在執行預覽上，跟股票
        # 設定列的 weight/price 是同一個做法。
        self.order_mode = tk.StringVar(value="pre")
        self._order_mode_last = "pre"
        self.order_ticks = tk.StringVar(value="2")
        self.order_ticks.trace_add("write", lambda *_a: self._recompute_order_preview())

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
        self.order_exec_last_note = ""    # 上一筆自動送出的結果，_update_order_exec_ui 顯示用

    def refresh_order_data(self):
        """
        重新讀 Excel：把每個已知名字的帳戶分頁的持股（E/F 欄）跟 B17 都讀一次。

        只讀不寫。帳戶名單只能從 self.trader_of 來——那是「登入過才知道名字」
        的既有規則（見 ui.py），還沒登入過的帳戶這裡也看不到，跟同步分頁的
        範圍選單是同一個限制，不是這裡另外加的。
        """
        if self.order_busy or not self._require_excel():
            return
        names = sorted(set(self.trader_of.values()))
        if not names:
            messagebox.showinfo(
                "還沒有帳戶名字",
                "還沒有任何帳戶登入過，名字都還不知道。\n請先到「同步」分頁按「登入」。",
                parent=self.root)
            return

        self.order_busy = True
        self.order_refresh_button.configure(state="disabled")
        self.order_status.configure(text="讀取中…")
        threading.Thread(target=self._order_read_worker, args=(self.path, names), daemon=True).start()

    def _order_read_worker(self, path, names):
        """背景執行緒：只用 COM 讀 E/F 欄跟 B17，不寫任何東西。"""
        import pythoncom

        pythoncom.CoInitialize()
        excel = workbook = sheet = None
        payload = {}
        try:
            excel, workbook, attached = excel_io.open_workbook(path, False)
            try:
                sheets, errors = {}, {}
                for name in names:
                    sheet, error = excel_io.find_sheet(workbook, name)
                    if sheet is None:
                        errors[name] = error
                        continue
                    data = excel_io.read_sheet(sheet)
                    data["return_rate"] = excel_io.read_return_rate(sheet)
                    sheets[name] = data
                payload = {"sheets": sheets, "errors": errors}
            finally:
                excel_io.close_workbook(excel, workbook, attached)
        except Exception as exc:
            payload = {"error": str(exc)}
        finally:
            sheet = excel = workbook = None
            pythoncom.CoUninitialize()
        self.queue.put(("order_data", payload))

    def _on_order_data(self, payload):
        self.order_busy = False
        self.order_refresh_button.configure(state="normal")

        if "error" in payload:
            self.order_status.configure(text="讀取失敗")
            messagebox.showerror("讀取失敗", payload["error"])
            return

        self.order_holdings, self.order_return_rates, self.order_names = {}, {}, {}
        for name, data in payload["sheets"].items():
            self.order_return_rates[name] = data["return_rate"]
            for row in data["rows"]:
                self.order_holdings[(name, row["code"])] = row["qty"]
                self.order_names.setdefault(row["code"], row["label"].split("(")[0].split("（")[0].strip())

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

        self.order_stock_box.configure(
            text="指定股票（比重整批共用「追價檔數」）" if intraday
            else "指定股票（各自設定比重與價格）")
        self.order_ticks_entry.configure(state="normal" if intraday else "disabled")

        for row in list(self.order_rows):
            row["frame"].destroy()
        self.order_rows = []
        self._recompute_order_preview()

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
        self._recompute_order_preview()

    def _build_order_stock_row(self, row):
        """
        一檔股票一區塊，分兩行：股票名稱＋買賣別在上面，比重（／價格）在下面。
        分兩行是因為左邊這個窄面板一行塞不下「名稱＋比重＋價格＋移除」，
        會把價格輸入框擠到剩沒幾個像素、打不進去字——用 pack 而不是 grid
        編號，移除中間一列時不會留下空位，不必自己重新排列剩下的列。

        買賣別目前固定是「賣」（盤前／盤中規劃都只支援賣出，見
        orders.SIDE_SELL），底色跟網站本身買紅賣綠的配色一致（Sell.TLabel／
        Buy.TLabel 在 ui_layout._build() 裡註冊），不必看文字就認得出方向。

        `"price" in row` 決定要不要畫價格輸入框——盤中模式的 row 沒有這個
        key（見 add_order_stock），不是留白也不是畫一個不會被讀的欄位。
        """
        block = ttk.Frame(self.order_stock_frame)
        block.pack(fill="x", pady=(0, 8))

        head = ttk.Frame(block)
        head.pack(fill="x")
        ttk.Label(head, text="賣", style="Sell.TLabel", width=2, anchor="center").pack(side="left")
        ttk.Label(head, text=f" {row['code']} {row['name']} ", style="Sell.TLabel").pack(side="left")
        ttk.Button(head, text="移除", bootstyle="danger-outline",
                  command=lambda: self.remove_order_stock(row)).pack(side="right")

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
        row["frame"] = block

    def remove_order_stock(self, row):
        row["frame"].destroy()
        self.order_rows.remove(row)
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
        for item in self.order_preview.get_children():
            self.order_preview.delete(item)

        ordered, skipped = orders.order_accounts(self._selected_order_accounts())
        stock_settings = self._order_stock_settings()
        hints = []

        if self.order_mode.get() == "intraday":
            ticks = self._order_ticks_setting()
            if ticks is None:
                preview = []
                hints.append("⚠ 追價檔數要填 0 以上的整數。")
            else:
                preview = orders.plan_intraday_orders(stock_settings, ordered, self.order_holdings, ticks)
        else:
            preview = orders.plan_stock_orders(stock_settings, ordered, self.order_holdings)

        side_names = {"B": "買", "S": "賣"}
        for item in preview:
            # 跳過的列不上買賣底色，只淡化文字——已經跳過了，不該看起來像
            # 真的會發生的一筆交易（見 ui_layout._build_order_right 的 tag_configure）。
            tag = "skip" if item["skip"] else {"B": "buy", "S": "sell"}.get(item["side"], "")
            # 盤中模式的 price 是 None（還沒查到即時成交價，見
            # orders.plan_intraday_orders），畫面上顯示文字說明，不是空白
            # 也不是猜一個數字。
            price_text = item["price"] if item["price"] is not None else "下單前才查"
            self.order_preview.insert("", "end", values=(
                item["order"], item["sheet"], side_names.get(item["side"], item["side"]),
                f"{item['code']} {item['name']}",
                # 持股最小單位是 1 股，不需要小數點；股數本來就可能上看百萬，
                # 千分位才看得出位數（util.show 是全專案統一用的數字顯示格式）。
                show(item["held_qty"]), item["lots"], price_text, item["note"],
            ), tags=(tag,) if tag else ())

        if skipped:
            hints.append(f"⚠ {'、'.join(a['sheet'] for a in skipped)} 讀不到報酬率，沒有排進執行順序。")
        self.order_preview_hint.configure(text="　".join(hints))

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

        模式（盤前／盤中）、盤中的追價檔數，都在這裡凍結成
        self.order_exec_mode／self.order_exec_ticks，之後每一筆都用凍結的值，
        不是每次都重讀畫面——理由跟凍結 queue 一樣：按下「開始下單」之後
        使用者還是可以去改上面的設定，那些改動只影響「下一輪」，不能半路
        插進正在跑的這一輪。
        """
        if self.order_exec_queue:
            self._dispatch_next_order()
            return

        if self.busy:
            return

        mode = self.order_mode.get()
        ordered, _skipped = orders.order_accounts(self._selected_order_accounts())
        stock_settings = self._order_stock_settings()

        if mode == "intraday":
            ticks = self._order_ticks_setting()
            if ticks is None:
                messagebox.showerror("追價檔數不對", "追價檔數要填 0 以上的整數。", parent=self.root)
                return
            preview = orders.plan_intraday_orders(stock_settings, ordered, self.order_holdings, ticks)
            queue_rows = orders.executable_intraday_orders(preview)
        else:
            ticks = None
            preview = orders.plan_stock_orders(stock_settings, ordered, self.order_holdings)
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

        if mode == "intraday":
            head = (
                f"即將依序處理 {len(queue_rows)} 筆委託（共 {total_lots} 張），用 IOC。\n\n"
                f"每一筆的價格會在下單前那一刻現查成交價、往下追 {ticks} 檔算出來，"
                f"不是現在看到的數字；沒成交的部位 IOC 會自動取消，不會掛著。\n"
            )
        else:
            head = f"即將依序處理 {len(queue_rows)} 筆委託（共 {total_lots} 張），用 ROD-當日有效。\n\n"

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

        self.order_exec_queue = queue_rows
        self.order_exec_pos = 0
        self.order_exec_mode = mode
        self.order_exec_ticks = ticks
        self.order_exec_auto = auto
        self._set_busy(True, "下單：準備第 1 筆…")
        self._dispatch_next_order()

    def _dispatch_next_order(self):
        row = self.order_exec_queue[self.order_exec_pos]
        order_number = self._order_number_for_sheet(row["sheet"])
        if order_number is None:
            messagebox.showerror("找不到帳戶", f"{row['sheet']} 對不到任何一組帳號，這筆沒辦法執行，這一輪停止。")
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
                      self.order_exec_ticks, self.order_exec_auto)))

    def stop_order_execution(self):
        """
        放棄這一輪。瀏覽器裡如果還留著一個沒處理的委託確認視窗，程式不會再
        幫忙追蹤它——那一頁還在，使用者自己回去按「確認」或「取消」就好，
        只是「下一筆」的自動流程到這裡為止。
        """
        if not self.order_exec_queue:
            return
        if not ask_confirm(
                self.root, "停止下單",
                "確定要停止這一輪嗎？\n\n如果瀏覽器裡還留著一個沒處理的委託確認視窗，"
                "程式不會再幫你追蹤它，請自己到瀏覽器裡按「確認」或「取消」。",
                confirm_style="primary"):
            return
        self.order_exec_queue = []
        self.order_exec_pos = 0
        self.order_exec_busy = False
        self.order_exec_watching = False
        self.order_exec_last_note = ""
        self._set_busy(False)
        self._update_order_exec_ui()
        self._say("下單：已停止這一輪。")

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
            self._set_busy(False)
            self._say("下單：這一輪已經跑完。")
        self._update_order_exec_ui()

    def _update_order_exec_ui(self):
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
        if self.order_exec_watching:
            self.order_exec_button.configure(text="下一筆", state="disabled")
            note = self.order_exec_last_note or (
                "委託確認視窗已開啟，去瀏覽器確認或取消，視窗關閉後才能按「下一筆」。")
            self.order_exec_status.configure(text=f"({pos + 1}/{total}) {label}：{note}")
        elif self.order_exec_busy:
            self.order_exec_button.configure(text="處理中…", state="disabled")
            self.order_exec_status.configure(text=f"({pos + 1}/{total}) {label}：登入／填單中…")
        else:
            # 不在等視窗、也不在忙，卡在這裡代表上一筆失敗了（剛開始執行的當下
            # 一定馬上進 order_exec_busy，不會停留在這個分支，見 start_order_execution）。
            self.order_exec_button.configure(text=f"下一筆（{pos + 1}/{total}）", state="normal")
            self.order_exec_status.configure(text=f"({pos + 1}/{total}) {label}：上一筆失敗，可以重試或停止。")

    def _order_fill_job(self, context, store, order_number, account, row, mode, ticks, auto):
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

        mode=="intraday" 時 row["price"] 是 None（見 orders.plan_intraday_orders），
        要在這裡現查一次成交價、用 orders.chase_price 往下追 ticks 檔算出真正
        要送出的價格——查價要在換去下單頁之前做，跟 collect() 查「未實現
        損益」用的是同一支查詢、同一個頁面情境（帳戶頁），這是唯一驗證過
        這支查詢能跑的地方，不要在別的頁面上呼叫它。
        """
        page, session, problems = fetch_mod.ensure_logged_in(
            context, [(order_number, account)], store)[order_number]
        if problems:
            raise RuntimeError("；".join(problems))

        price = row["price"]
        if mode == "intraday":
            bid, cid = session.get("branch_id"), session.get("cust_id")
            expect_code = fetch_mod.account_code(session)
            pricenow = fetch_mod.current_price(page, bid, cid, row["code"], expect_code)
            if pricenow is None:
                raise RuntimeError(f"查不到 {row['code']} 現在的成交價，這一筆沒辦法算追價。")
            price = orders.chase_price(pricenow, ticks)

        # 委託別跟著模式走，不是兩邊共用同一個值（見 orders.BS_FLAG_PRE 的
        # 說明）：盤前開盤前還沒有連續交易，只能用 ROD；盤中規劃文件明講
        # 用 IOC。2026/08/28 使用者更正過，之前這裡兩種模式都寫死 IOC 是錯的。
        bs_flag = orders.BS_FLAG_INTRADAY if mode == "intraday" else orders.BS_FLAG_PRE

        page.goto(order_fill.ORDER_ENTRY_PAGE, wait_until="domcontentloaded")
        order_fill.open_order_form(page)
        order_fill.select_stock(page, row["code"])
        order_fill.fill_order(page, side="S", qty=row["lots"], price=price, bs_flag=bs_flag)

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
