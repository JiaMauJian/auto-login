"""
下單分頁的行為：目前只有「盤前」模式——股票／比重／價格設定、勾帳戶、
算出執行預覽。

只做到「看得到等一下會發生什麼事」為止，還沒有真的操作下單頁面送出委託——
那一步要動真實表單，風險不一樣，先把這裡的計算跟畫面搞對再說。

比重→張數、帳戶依 B17 報酬率排序、組出預覽清單，全部是 orders.py 的純函式，
這裡只負責收使用者的輸入、讀 Excel、把結果畫出來。
"""

import threading
import tkinter as tk
from tkinter import messagebox

import ttkbootstrap as ttk

import excel_io
import orders
from ui_common import FONT_SIZE
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

    # ---------- 股票設定 ----------

    def add_order_stock(self):
        """
        把下拉選單（或手動輸入）裡的股票加進設定清單，一檔一列，各自的比重／
        價格獨立輸入——這是使用者確認過的決定：不同股票通常想賣的比例、
        價格都不一樣，共用一個值沒意義。
        """
        raw = self.order_stock_pick.get().strip()
        if not raw:
            return
        code = raw.split(" ")[0].strip().upper()
        if any(row["code"] == code for row in self.order_rows):
            messagebox.showinfo("已經加過了", f"{code} 已經在清單裡了。", parent=self.root)
            return
        name = self.order_names.get(code) or (raw.split(" ", 1)[1].strip() if " " in raw else code)

        row = {"code": code, "name": name, "weight": tk.StringVar(), "price": tk.StringVar()}
        row["weight"].trace_add("write", lambda *_a: self._recompute_order_preview())
        row["price"].trace_add("write", lambda *_a: self._recompute_order_preview())
        self._build_order_stock_row(row)
        self.order_rows.append(row)
        self.order_stock_pick.set("")
        self._recompute_order_preview()

    def _build_order_stock_row(self, row):
        """
        一檔股票一區塊，分兩行：股票名稱＋買賣別在上面，比重／價格在下面。
        分兩行是因為左邊這個窄面板一行塞不下「名稱＋比重＋價格＋移除」，
        會把價格輸入框擠到剩沒幾個像素、打不進去字——用 pack 而不是 grid
        編號，移除中間一列時不會留下空位，不必自己重新排列剩下的列。

        買賣別目前固定是「賣」（盤前規劃只支援賣出比重，見 orders.SIDE_SELL），
        底色跟網站本身買紅賣綠的配色一致（Sell.TLabel／Buy.TLabel 在
        ui_layout._build() 裡註冊），不必看文字就認得出方向。
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

    def _recompute_order_preview(self):
        """
        比重／價格／勾選的帳戶，任何一個變了就整份重算重畫——跟同步分頁
        fill_sync_tree() 同一個做法，整份重建比自己追蹤哪一列該更新可靠。
        """
        for item in self.order_preview.get_children():
            self.order_preview.delete(item)

        ordered, skipped = orders.order_accounts(self._selected_order_accounts())

        stock_settings = []
        for row in self.order_rows:
            try:
                weight = float(row["weight"].get())
            except ValueError:
                weight = 0
            stock_settings.append({
                "code": row["code"], "name": row["name"],
                "weight_pct": weight, "price": row["price"].get(),
            })

        side_names = {"B": "買", "S": "賣"}
        for item in orders.plan_stock_orders(stock_settings, ordered, self.order_holdings):
            # 跳過的列不上買賣底色，只淡化文字——已經跳過了，不該看起來像
            # 真的會發生的一筆交易（見 ui_layout._build_order_right 的 tag_configure）。
            tag = "skip" if item["skip"] else {"B": "buy", "S": "sell"}.get(item["side"], "")
            self.order_preview.insert("", "end", values=(
                item["order"], item["sheet"], side_names.get(item["side"], item["side"]),
                f"{item['code']} {item['name']}",
                # 持股最小單位是 1 股，不需要小數點；股數本來就可能上看百萬，
                # 千分位才看得出位數（util.show 是全專案統一用的數字顯示格式）。
                show(item["held_qty"]), item["lots"], item["price"], item["note"],
            ), tags=(tag,) if tag else ())

        if skipped:
            names = "、".join(a["sheet"] for a in skipped)
            self.order_preview_hint.configure(text=f"⚠ {names} 讀不到報酬率，沒有排進執行順序。")
        else:
            self.order_preview_hint.configure(text="")
