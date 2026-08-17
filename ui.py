"""
持股同步的桌面介面。

    python ui.py

三個分頁分別對應三件事：

    同步      這次要改哪幾格、哪幾格程式不准碰，勾選後寫入
    歷程      誰在什麼時候改了哪一格（程式、人工、交接、補登）
    現金帳本  基準與逐日流水，可以補登漏掉的日子、重新校正基準

畫面只做顯示與操作，所有判斷都來自 planner.py —— 介面與命令列走同一段程式碼，
才不會出現「介面接管的結果跟命令列不一樣」這種最難查的問題。

執行緒
------
登入、抓資料、開 Excel 都很慢，全部丟到背景執行緒，否則視窗會整個凍住。
Tk 的元件只能在主執行緒碰，所以背景做完只把純資料丟進 queue，
由主執行緒定時取出來畫 —— 不在背景執行緒動任何 widget。

背景執行緒要用 COM（Excel）之前一定要先 CoInitialize，這是 Windows 的規定，
少了它 win32com 會直接丟例外。
"""

import datetime
import json
import queue
import threading
import traceback
import tkinter as tk
from tkinter import font as tkfont
from tkinter import messagebox, ttk

from playwright.sync_api import Error as PlaywrightError
from playwright.sync_api import sync_playwright

import excel_io
import ledger as ledger_mod
import planner
from fetch import collect
from login import configure_browsers_path, load_accounts, open_context
from util import show

SOURCE_NAMES = {"program": "程式", "human": "人工", "adopt": "交接", "backfill": "補登"}

CHECKED, UNCHECKED = "☑", "☐"


def pick_font():
    """挑一個有中文字的字型。挑不到就交給 Tk 自己決定。"""
    families = set(tkfont.families())
    for name in ("Microsoft JhengHei UI", "Microsoft JhengHei", "Segoe UI"):
        if name in families:
            return name
    return "TkDefaultFont"


class SyncApp:
    def __init__(self, root):
        self.root = root
        self.path = excel_io.excel_path()
        self.accounts = load_accounts()
        self.today = datetime.date.today()

        self.records = {}        # 分頁名 -> 網頁資料
        self.sheet_data = {}     # 分頁名 -> Excel 現值
        self.proposals = {}      # 分頁名 -> 提案清單
        self.warnings = {}       # 分頁名 -> 提醒
        self.problems = []       # 整組失敗的原因
        self.checked = {}        # (分頁名, 格子) -> 要不要寫
        self.row_map = {}        # Treeview 的 iid -> (分頁名, 提案)
        self.busy = False
        self.queue = queue.Queue()

        self.ledger = None
        self.ledger_error = None
        try:
            self.ledger = ledger_mod.Ledger(self.path)
        except RuntimeError as exc:
            self.ledger_error = str(exc)

        self._build()
        self._drain()

        if not self.path.is_file():
            self._say(f"找不到 Excel：{self.path}")
        elif self.ledger_error:
            messagebox.showerror("紀錄檔有問題", self.ledger_error)
        elif not self.accounts:
            self._say("找不到帳號設定，請先在 .env 填入 TBB_ID_1 / TBB_PASSWORD_1")
        else:
            self._say("按「讀取網頁資料」開始。這會開瀏覽器自動登入，大約要十幾秒。")

        self.refresh_history()
        self.refresh_cash()

    # ---------- 版面 ----------

    def _build(self):
        family = pick_font()
        self.root.title("持股同步")
        self.root.geometry("1180x720")
        self.root.minsize(900, 560)

        style = ttk.Style()
        style.configure("Treeview", font=(family, 10), rowheight=26)
        style.configure("Treeview.Heading", font=(family, 10, "bold"))
        style.configure("TButton", font=(family, 10))
        style.configure("TLabel", font=(family, 10))
        style.configure("Hint.TLabel", font=(family, 9), foreground="#666666")
        style.configure("Big.TButton", font=(family, 10, "bold"))
        self.family = family

        top = ttk.Frame(self.root, padding=(12, 10, 12, 6))
        top.pack(fill="x")

        ttk.Label(top, text="Excel").grid(row=0, column=0, sticky="w")
        ttk.Label(top, text=str(self.path), style="Hint.TLabel").grid(row=0, column=1, sticky="w", padx=(8, 0))

        ttk.Label(top, text="帳號").grid(row=1, column=0, sticky="w", pady=(6, 0))
        choices = ["全部"] + [f"第 {i} 組" for i in range(1, len(self.accounts) + 1)]
        self.account_choice = ttk.Combobox(top, values=choices, state="readonly", width=10, font=(family, 10))
        self.account_choice.current(0)
        self.account_choice.grid(row=1, column=1, sticky="w", padx=(8, 0), pady=(6, 0))

        self.fetch_button = ttk.Button(top, text="讀取網頁資料", style="Big.TButton", command=self.start_fetch)
        self.fetch_button.grid(row=1, column=2, sticky="e", padx=(16, 0), pady=(6, 0))
        top.columnconfigure(1, weight=1)

        self.progress = ttk.Progressbar(self.root, mode="indeterminate")

        self.tabs = ttk.Notebook(self.root)
        self.tabs.pack(fill="both", expand=True, padx=12, pady=(6, 0))
        self.tabs.bind("<<NotebookTabChanged>>", self._on_tab_changed)

        self._build_sync_tab()
        self._build_history_tab()
        self._build_cash_tab()

        self.status = ttk.Label(self.root, text="", style="Hint.TLabel", anchor="w", padding=(12, 6))
        self.status.pack(fill="x")

    def _build_sync_tab(self):
        frame = ttk.Frame(self.tabs, padding=8)
        self.tabs.add(frame, text="  同步  ")

        columns = ("cell", "label", "current", "web", "proposed", "status", "note")
        self.tree = ttk.Treeview(frame, columns=columns, show="tree headings", selectmode="browse")
        self.tree.heading("#0", text="寫入")
        self.tree.column("#0", width=110, minwidth=90, stretch=False)
        for key, title, width, anchor in (
            ("cell", "格子", 60, "w"),
            ("label", "項目", 230, "w"),
            ("current", "Excel 現值", 100, "e"),
            ("web", "網頁值", 100, "e"),
            ("proposed", "會變成", 100, "e"),
            ("status", "狀態", 80, "center"),
            ("note", "說明", 380, "w"),
        ):
            self.tree.heading(key, text=title)
            self.tree.column(key, width=width, anchor=anchor, stretch=(key == "note"))

        bar = ttk.Scrollbar(frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=bar.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        bar.grid(row=0, column=1, sticky="ns")

        # 顏色只用來輔助，不是唯一線索 —— 狀態欄本身就寫著字。
        self.tree.tag_configure("write", background="#eaf4ea")
        self.tree.tag_configure("manual", foreground="#a34a00")
        self.tree.tag_configure("untracked", foreground="#777777")
        self.tree.tag_configure("sheet", font=(self.family, 10, "bold"))

        self.tree.bind("<Button-1>", self._on_tree_click)

        self.warn_box = tk.Text(frame, height=6, wrap="word", font=(self.family, 9),
                                background="#fbfbfb", relief="flat", state="disabled")
        self.warn_box.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(8, 0))

        buttons = ttk.Frame(frame)
        buttons.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(8, 0))
        self.check_all_button = ttk.Button(buttons, text="全選", command=lambda: self._check_all(True))
        self.check_all_button.pack(side="left")
        self.uncheck_all_button = ttk.Button(buttons, text="全不選", command=lambda: self._check_all(False))
        self.uncheck_all_button.pack(side="left", padx=(6, 0))
        self.takeback_button = ttk.Button(buttons, text="交還給程式", command=self.takeback, state="disabled")
        self.takeback_button.pack(side="left", padx=(16, 0))
        self.write_button = ttk.Button(buttons, text="寫入", style="Big.TButton",
                                       command=self.start_write, state="disabled")
        self.write_button.pack(side="right")

        self.tree.bind("<<TreeviewSelect>>", lambda _event: self._sync_buttons())

        frame.rowconfigure(0, weight=1)
        frame.columnconfigure(0, weight=1)

    def _build_history_tab(self):
        frame = ttk.Frame(self.tabs, padding=8)
        self.tabs.add(frame, text="  歷程  ")

        columns = ("at", "sheet", "cell", "label", "change", "by", "note")
        self.history_tree = ttk.Treeview(frame, columns=columns, show="headings")
        for key, title, width, anchor in (
            ("at", "時間", 140, "w"),
            ("sheet", "分頁", 90, "w"),
            ("cell", "格子", 60, "w"),
            ("label", "項目", 200, "w"),
            ("change", "變化", 170, "w"),
            ("by", "來源", 60, "center"),
            ("note", "說明", 420, "w"),
        ):
            self.history_tree.heading(key, text=title)
            self.history_tree.column(key, width=width, anchor=anchor, stretch=(key == "note"))

        bar = ttk.Scrollbar(frame, orient="vertical", command=self.history_tree.yview)
        self.history_tree.configure(yscrollcommand=bar.set)
        self.history_tree.grid(row=0, column=0, sticky="nsew")
        bar.grid(row=0, column=1, sticky="ns")

        self.history_tree.tag_configure("human", foreground="#a34a00")
        self.history_tree.tag_configure("adopt", foreground="#1a5fb4")

        row = ttk.Frame(frame)
        row.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(8, 0))
        self.history_hint = ttk.Label(row, text="", style="Hint.TLabel")
        self.history_hint.pack(side="left")
        ttk.Button(row, text="重新整理", command=self.refresh_history).pack(side="right")

        frame.rowconfigure(0, weight=1)
        frame.columnconfigure(0, weight=1)

    def _build_cash_tab(self):
        frame = ttk.Frame(self.tabs, padding=8)
        self.tabs.add(frame, text="  現金帳本  ")

        head = ttk.Frame(frame)
        head.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 8))
        ttk.Label(head, text="分頁").pack(side="left")
        self.cash_sheet = ttk.Combobox(head, state="readonly", width=16, font=(self.family, 10))
        self.cash_sheet.pack(side="left", padx=(8, 16))
        self.cash_sheet.bind("<<ComboboxSelected>>", lambda _event: self.refresh_cash(keep=True))
        self.cash_head = ttk.Label(head, text="", style="Hint.TLabel")
        self.cash_head.pack(side="left")

        columns = ("date", "net", "source", "balance")
        self.cash_tree = ttk.Treeview(frame, columns=columns, show="headings")
        for key, title, width, anchor in (
            ("date", "日期", 130, "w"),
            ("net", "淨收付", 110, "e"),
            ("source", "來源", 160, "w"),
            ("balance", "餘額", 110, "e"),
        ):
            self.cash_tree.heading(key, text=title)
            self.cash_tree.column(key, width=width, anchor=anchor, stretch=(key == "source"))
        self.cash_tree.tag_configure("missing", foreground="#a34a00")

        bar = ttk.Scrollbar(frame, orient="vertical", command=self.cash_tree.yview)
        self.cash_tree.configure(yscrollcommand=bar.set)
        self.cash_tree.grid(row=1, column=0, sticky="nsew")
        bar.grid(row=1, column=1, sticky="ns")

        foot = ttk.Frame(frame)
        foot.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(8, 0))
        self.cash_foot = ttk.Label(foot, text="")
        self.cash_foot.pack(side="left")
        ttk.Button(foot, text="重新校正…", command=self.recalibrate).pack(side="right")
        ttk.Button(foot, text="補登…", command=self.backfill).pack(side="right", padx=(0, 6))

        frame.rowconfigure(1, weight=1)
        frame.columnconfigure(0, weight=1)

    # ---------- 背景工作 ----------

    def start_fetch(self):
        if self.busy:
            return
        if not self.path.is_file():
            messagebox.showerror("找不到 Excel", f"{self.path}\n\n可以在 .env 用 EXCEL_PATH 指定位置。")
            return
        if not self.accounts:
            messagebox.showerror("沒有帳號", "請先在 .env 填入 TBB_ID_1 / TBB_PASSWORD_1。")
            return
        if self.ledger_error:
            messagebox.showerror("紀錄檔有問題", self.ledger_error)
            return

        choice = self.account_choice.current()
        numbered = list(enumerate(self.accounts, start=1))
        selected = numbered if choice == 0 else [numbered[choice - 1]]

        self._set_busy(True, "登入中，瀏覽器會自己開起來，請不要關掉它…")
        threading.Thread(target=self._fetch_worker, args=(selected, self.path), daemon=True).start()

    def _fetch_worker(self, selected, path):
        """背景：登入、抓網頁、讀 Excel。只回傳純資料，不回傳任何 COM 物件。"""
        import pythoncom

        payload = {}
        try:
            configure_browsers_path()
            with sync_playwright() as p:
                context, browser = open_context(p)
                try:
                    records = collect(context, selected)
                finally:
                    # 抓完就關瀏覽器。留著它只會佔一個視窗，而且 session 隨時會過期，
                    # 下次讀取本來就得重新登入。
                    try:
                        context.close()
                        if browser is not None:
                            browser.close()
                    except PlaywrightError:
                        pass

            pythoncom.CoInitialize()
            try:
                excel, workbook, attached = excel_io.open_workbook(path, False)
                try:
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
                    excel_io.close_workbook(excel, workbook, attached)
            finally:
                pythoncom.CoUninitialize()

            payload = {"records": records, "sheets": sheets, "sheet_errors": sheet_errors,
                       "attached": attached}
        except Exception:
            payload = {"error": traceback.format_exc()}

        self.queue.put(("fetched", payload))

    def start_write(self):
        if self.busy:
            return

        # 沒勾的就地改成「不寫」，這樣待會兒 commit 記進紀錄檔的內容
        # 才會跟真正寫進 Excel 的一致。
        writes = {}
        total = 0
        for name, items in self.proposals.items():
            for item in items:
                if item["will_write"] and not self.checked.get((name, item["cell"]), False):
                    item["will_write"] = False
            cells = [(i["row"], i["col"], i["proposed"]) for i in items if i["will_write"]]
            if cells:
                writes[name] = cells
                total += len(cells)

        if not total:
            messagebox.showinfo("沒有要寫的", "目前沒有勾選任何要寫入的格子。")
            return

        lines = []
        for name, items in self.proposals.items():
            for item in items:
                if item["will_write"]:
                    lines.append(f"  {name} {item['cell']} {item['label']}："
                                 f"{show(item['current'])} → {show(item['proposed'])}")
        if not messagebox.askokcancel("確認寫入", f"要寫入 {total} 格：\n\n" + "\n".join(lines) +
                                      "\n\n寫入前會自動備份。"):
            return

        self._set_busy(True, "寫入中…")
        threading.Thread(target=self._write_worker, args=(self.path, writes), daemon=True).start()

    def _write_worker(self, path, writes):
        import pythoncom

        payload = {}
        pythoncom.CoInitialize()
        try:
            excel, workbook, attached = excel_io.open_workbook(path, True)
            try:
                saved = excel_io.backup(path)
                for name, cells in writes.items():
                    sheet, error = excel_io.find_sheet(workbook, name)
                    if sheet is None:
                        raise RuntimeError(error)
                    excel_io.write_cells(sheet, cells)
                # 檔案是我們自己開的就要自己存檔，close 是 Close(False)。
                # 接上使用者開著的 Excel 時刻意不存，留給他自己按 Ctrl+S。
                if not attached:
                    workbook.Save()
            finally:
                excel_io.close_workbook(excel, workbook, attached)
            payload = {"backup": str(saved), "attached": attached}
        except Exception:
            payload = {"error": traceback.format_exc()}
        finally:
            pythoncom.CoUninitialize()

        self.queue.put(("written", payload))

    def _drain(self):
        """主執行緒每 120ms 看一次背景有沒有結果。widget 只在這裡被碰。"""
        try:
            while True:
                kind, payload = self.queue.get_nowait()
                if kind == "fetched":
                    self._on_fetched(payload)
                elif kind == "written":
                    self._on_written(payload)
        except queue.Empty:
            pass
        self.root.after(120, self._drain)

    def _on_fetched(self, payload):
        self._set_busy(False)

        if "error" in payload:
            self._say("讀取失敗")
            messagebox.showerror("讀取失敗", payload["error"][-2000:])
            return

        self.records, self.problems = {}, []
        for record in payload["records"]:
            if record["problems"]:
                self.problems.append(f"第 {record['order']} 組：" + "；".join(record["problems"]))
            elif record.get("sheet_name"):
                self.records[record["sheet_name"]] = record
        for name, error in payload["sheet_errors"].items():
            self.problems.append(error)

        self.sheet_data = payload["sheets"]
        self.today = datetime.date.today()
        self.replan()

        if payload.get("attached"):
            self._say("已讀取。這個 Excel 正開著，寫入只會進到記憶體，要自己按 Ctrl+S 才會存檔。")
        else:
            self._say("已讀取。")

    def _on_written(self, payload):
        self._set_busy(False)

        if "error" in payload:
            self._say("寫入失敗")
            messagebox.showerror("寫入失敗", payload["error"][-2000:])
            return

        # 紀錄檔一定在 Excel 寫成功之後才更新。順序反過來的話，寫入失敗會留下
        # 一份「以為自己寫過了」的帳本，之後每次比對都判定成人工改動。
        at = datetime.datetime.now().isoformat(timespec="seconds")
        events = []
        for name, items in self.proposals.items():
            book = self.ledger.sheet(name)
            events.extend(planner.commit(items, book, name, self.today, at))
        self.ledger.save()
        self.ledger.append_history(events)

        # Excel 已經被改過了，手上的現值是舊的，重新讀一次才會準。
        for name, items in self.proposals.items():
            data = self.sheet_data.get(name)
            if not data:
                continue
            for item in items:
                if item["will_write"]:
                    if item["kind"] == "cash":
                        data["balance"] = item["proposed"]
                    else:
                        for line in data["rows"]:
                            if line["row"] == item["row"]:
                                line["qty" if item["which"] == "qty" else "cost"] = item["proposed"]

        self.replan()
        self.refresh_history()

        if payload["attached"]:
            messagebox.showinfo(
                "寫好了",
                f"已備份到\n{payload['backup']}\n\n"
                f"變更已經寫進你開著的 Excel，但還沒存檔：\n"
                f"  滿意   → 切到 Excel 按 Ctrl+S\n"
                f"  不滿意 → 關閉 Excel 時選「不要儲存」\n\n"
                f"程式做的變更沒辦法用 Ctrl+Z 復原，只能靠不存檔或那份備份。",
            )
            self._say("已寫入你開著的 Excel，記得按 Ctrl+S。")
        else:
            messagebox.showinfo("寫好了", f"已存檔。\n\n備份在\n{payload['backup']}")
            self._say("已寫入並存檔。")

    def _set_busy(self, busy, message=""):
        self.busy = busy
        self.fetch_button.configure(state="disabled" if busy else "normal")
        if busy:
            self.progress.pack(fill="x", padx=12, pady=(4, 0), before=self.tabs)
            self.progress.start(12)
            self._say(message)
        else:
            self.progress.stop()
            self.progress.pack_forget()
        self._sync_buttons()

    def _say(self, message):
        self.status.configure(text=message)

    # ---------- 同步分頁 ----------

    def replan(self):
        """重新計算提案並重畫。plan 是純函式，可以隨便重跑。"""
        self.proposals, self.warnings = {}, {}
        for name, record in self.records.items():
            data = self.sheet_data.get(name)
            if not data:
                continue
            book = self.ledger.sheet(name)
            book["account_code"] = record.get("account_code", "")
            items, warns = planner.plan(data, record, book, self.today)
            self.proposals[name] = items
            self.warnings[name] = warns

        self.checked = {
            (name, item["cell"]): item["will_write"]
            for name, items in self.proposals.items() for item in items
        }
        self.fill_sync_tree()
        self.refresh_cash()

    def fill_sync_tree(self):
        self.tree.delete(*self.tree.get_children())
        self.row_map = {}

        for name, items in self.proposals.items():
            parent = self.tree.insert("", "end", text=name, open=True, tags=("sheet",),
                                      values=("", "", "", "", "", "", ""))
            for item in items:
                tags = []
                if item["will_write"]:
                    tags.append("write")
                if item["status"] == ledger_mod.MANUAL:
                    tags.append("manual")
                elif item["status"] in (ledger_mod.UNTRACKED, planner.WEB_MISSING):
                    tags.append("untracked")

                iid = self.tree.insert(
                    parent, "end",
                    text=self._box(name, item),
                    values=(
                        item["cell"], item["label"],
                        show(item["current"]), show(item["web"]),
                        show(item["proposed"]) if item["will_write"] else "—",
                        planner.STATUS_NAMES.get(item["status"], item["status"]),
                        item["note"],
                    ),
                    tags=tuple(tags),
                )
                self.row_map[iid] = (name, item)

        text = []
        for name, warns in self.warnings.items():
            for warning in warns:
                text.append(f"• {warning}")
        for problem in self.problems:
            text.append(f"⚠ {problem}")
        if self.proposals and not any(item["will_write"] for _name, item in self.row_map.values()):
            text.append("目前沒有任何格子需要寫入 —— Excel 上的數字跟網頁已經一致。")
        if not text and self.proposals:
            text.append("沒有需要注意的事。")

        self.warn_box.configure(state="normal")
        self.warn_box.delete("1.0", "end")
        self.warn_box.insert("1.0", "\n".join(text))
        self.warn_box.configure(state="disabled")

        self._sync_buttons()

    def _box(self, name, item):
        if not item["will_write"]:
            return ""
        return CHECKED if self.checked.get((name, item["cell"]), False) else UNCHECKED

    def _on_tree_click(self, event):
        """點第一欄的方框就切換勾選。Treeview 沒有內建 checkbox，只能自己畫。"""
        if self.tree.identify_region(event.x, event.y) != "tree":
            return
        iid = self.tree.identify_row(event.y)
        if iid not in self.row_map:
            return
        name, item = self.row_map[iid]
        if not item["will_write"]:
            return
        key = (name, item["cell"])
        self.checked[key] = not self.checked.get(key, False)
        self.tree.item(iid, text=self._box(name, item))
        self._sync_buttons()

    def _check_all(self, value):
        for iid, (name, item) in self.row_map.items():
            if item["will_write"]:
                self.checked[(name, item["cell"])] = value
                self.tree.item(iid, text=self._box(name, item))
        self._sync_buttons()

    def _sync_buttons(self):
        count = sum(1 for (name, cell), on in self.checked.items() if on)
        self.write_button.configure(
            text=f"寫入勾選的 {count} 項" if count else "寫入",
            state="normal" if (count and not self.busy) else "disabled",
        )

        # 沒有任何可勾的格子時，全選/全不選按了也不會有反應，那就不要亮著。
        # 一顆按得下去卻什麼都不發生的按鈕，只會讓人以為程式壞了。
        checkable = any(item["will_write"] for _name, item in self.row_map.values())
        state = "normal" if (checkable and not self.busy) else "disabled"
        self.check_all_button.configure(state=state)
        self.uncheck_all_button.configure(state=state)

        selected = self.tree.selection()
        item = self.row_map.get(selected[0]) if selected else None
        can_take = bool(item) and item[1]["status"] in (ledger_mod.MANUAL, ledger_mod.UNTRACKED)
        self.takeback_button.configure(state="normal" if (can_take and not self.busy) else "disabled")

    def takeback(self):
        """把選中的那一格交還給程式管理。"""
        selected = self.tree.selection()
        if not selected or selected[0] not in self.row_map:
            return
        name, item = self.row_map[selected[0]]

        included = None
        if item["kind"] == "cash":
            included = ask_calibration(self.root, self.family, item, self.today)
            if included is None:
                return

        at = datetime.datetime.now().isoformat(timespec="seconds")
        book = self.ledger.sheet(name)
        message, event, needs_today = planner.adopt_one(item, book, name, self.today, included, at)

        if needs_today:
            return
        if event is None:
            if message:
                messagebox.showwarning("沒辦法交還", message)
            return

        self.ledger.save()
        self.ledger.append_history([event])
        self.replan()
        self.refresh_history()
        self._say(message)

    # ---------- 歷程分頁 ----------

    def refresh_history(self):
        self.history_tree.delete(*self.history_tree.get_children())
        path = self.ledger.history_path if self.ledger else None

        if not path or not path.is_file():
            self.history_hint.configure(text="還沒有任何歷程。第一次寫入或交接之後就會有。")
            return

        rows = []
        for line in path.read_text(encoding="utf-8").splitlines():
            line = line.strip()
            if not line:
                continue
            try:
                rows.append(json.loads(line))
            except json.JSONDecodeError:
                continue

        for event in reversed(rows):          # 最新的放最上面
            by = event.get("by", "")
            self.history_tree.insert(
                "", "end",
                values=(
                    (event.get("at") or "").replace("T", " "),
                    event.get("sheet", ""), event.get("cell", ""), event.get("label", ""),
                    describe_change(by, event.get("old"), event.get("new")),
                    SOURCE_NAMES.get(by, by),
                    event.get("note", ""),
                ),
                tags=(by,),
            )

        self.history_hint.configure(text=f"{len(rows)} 筆，最新的在最上面。檔案：{path.name}")

    # ---------- 現金帳本分頁 ----------

    def refresh_cash(self, keep=False):
        names = sorted((self.ledger.data.get("sheets") or {}) if self.ledger else {})
        current = self.cash_sheet.get() if keep else None
        self.cash_sheet.configure(values=names)
        if names:
            if current in names:
                self.cash_sheet.set(current)
            elif not self.cash_sheet.get() or self.cash_sheet.get() not in names:
                self.cash_sheet.set(names[0])
        else:
            self.cash_sheet.set("")

        self.cash_tree.delete(*self.cash_tree.get_children())
        name = self.cash_sheet.get()
        if not name:
            self.cash_head.configure(text="還沒有任何分頁的紀錄")
            self.cash_foot.configure(text="")
            return

        cash = self.ledger.sheet(name)["cash"]
        base = cash.get("baseline_value")
        if base is None:
            self.cash_head.configure(text="還沒設定基準。到「同步」分頁選 B8 那一列，按「交還給程式」。")
            self.cash_foot.configure(text="")
            return

        self.cash_head.configure(
            text=f"基準 {cash.get('baseline_date') or '?'}    {base:g} 元"
        )

        applied = cash.get("applied") or {}
        missing = {day.isoformat() for day in ledger_mod.missing_dates(cash, self.today)}
        running = base
        for day in sorted(set(applied) | missing):
            if day in missing:
                self.cash_tree.insert("", "end", values=(day, "——", "沒有紀錄，可補登", "——"),
                                      tags=("missing",))
                continue
            running = round(running + applied[day], 2)
            self.cash_tree.insert("", "end", values=(day, f"{applied[day]:g}", "程式", f"{running:g}"))

        expected = ledger_mod.expected_cash(cash)
        data = self.sheet_data.get(name)
        actual = data["balance"] if data else None
        if actual is None:
            note = "（還沒讀取 Excel）"
        elif abs(actual - expected) <= 0.005:
            note = "與 Excel 相符 ✓"
        else:
            note = f"Excel 上是 {actual:g}，差 {actual - expected:g} —— 同步分頁會把它補上"
        self.cash_foot.configure(text=f"應有餘額 {expected:g}    {note}")

    def backfill(self):
        """補登某一天的淨收付。漏跑的那天靠這個補回來。"""
        name = self.cash_sheet.get()
        if not name:
            return
        cash = self.ledger.sheet(name)["cash"]
        if cash.get("baseline_value") is None:
            messagebox.showinfo("還沒有基準", "要先設定現金基準才能補登。")
            return

        result = ask_backfill(self.root, self.family, self.today)
        if result is None:
            return
        day, amount = result

        at = datetime.datetime.now().isoformat(timespec="seconds")
        before = ledger_mod.expected_cash(cash)
        ledger_mod.record_net(cash, day, amount)
        after = ledger_mod.expected_cash(cash)

        self.ledger.save()
        self.ledger.append_history([{
            "at": at, "sheet": name, "cell": "B8", "label": "現金餘額", "by": "backfill",
            "old": before, "new": after,
            "note": f"補登 {day:%Y/%m/%d} 的淨收付 {amount:g}",
        }])
        self.replan()
        self.refresh_history()
        self._say(f"已補登 {day:%Y/%m/%d} 的淨收付 {amount:g}，應有餘額變成 {after:g}。")

    def recalibrate(self):
        """重新校正基準。等同於同步分頁上對 B8 按「交還給程式」。"""
        name = self.cash_sheet.get()
        if not name:
            return
        items = self.proposals.get(name) or []
        cash_item = next((i for i in items if i["kind"] == "cash"), None)
        if cash_item is None:
            messagebox.showinfo(
                "要先讀取網頁資料",
                "校正需要知道今天的淨收付是多少，才能判斷 Excel 上那個數字算過了沒。\n"
                "請先按上面的「讀取網頁資料」。",
            )
            return

        included = ask_calibration(self.root, self.family, cash_item, self.today)
        if included is None:
            return

        at = datetime.datetime.now().isoformat(timespec="seconds")
        book = self.ledger.sheet(name)
        cash_item = dict(cash_item, status=ledger_mod.MANUAL)   # 強制走交接那條路
        message, event, _ = planner.adopt_one(cash_item, book, name, self.today, included, at)
        if event is None:
            if message:
                messagebox.showwarning("沒辦法校正", message)
            return

        self.ledger.save()
        self.ledger.append_history([event])
        self.replan()
        self.refresh_history()
        self._say(message)

    def _on_tab_changed(self, _event):
        tab = self.tabs.index(self.tabs.select())
        if tab == 1:
            self.refresh_history()
        elif tab == 2:
            self.refresh_cash(keep=True)


def describe_change(by, old, new):
    """
    歷程那一欄要寫什麼。

    「交接」的新舊值講的是**程式的記憶**，不是格子的值：交接就是把
    「我記得這格是多少」改成 Excel 上現在的數字，Excel 本身一格都沒動。
    如果照 program 那樣印成「1 → 0」，看起來會像程式把數字改壞了，
    所以這裡明講是記憶在變。
    """
    if by == "adopt":
        return f"記得 {show(old)} → 改記 {show(new)}"
    if by == "human":
        return f"{show(old)} → {show(new)}（人工）"
    return f"{show(old)} → {show(new)}"


def ask_calibration(parent, family, item, today):
    """
    問「Excel 上這個數字含不含今天的淨收付」。回傳 True / False，取消是 None。

    這一題不能猜也不能給預設值：猜錯的後果是現金餘額從此多扣或少扣一次，
    而且畫面上不會有任何徵兆。所以兩個選項都不預選，非得使用者自己點。
    """
    win = tk.Toplevel(parent)
    win.title("現金餘額重新校正")
    win.transient(parent)
    win.resizable(False, False)

    body = ttk.Frame(win, padding=16)
    body.pack(fill="both", expand=True)

    remembered = "（還沒有紀錄）" if item["current"] is None else ""
    ttk.Label(body, text="以 Excel 上目前的數字為新基準，從今天起重新累加。",
              font=(family, 10)).pack(anchor="w")
    ttk.Label(body, text=f"\nExcel 的 B8 現在是    {show(item['current'])} {remembered}"
                         f"\n今天的淨收付是        {item['net']:g}"
                         f"（{item['net_rows']} 筆成交）",
              font=(family, 10)).pack(anchor="w")

    ttk.Separator(body).pack(fill="x", pady=12)
    ttk.Label(body, text=f"這個 {show(item['current'])} 有沒有已經算進今天的 {item['net']:g}？",
              font=(family, 10, "bold")).pack(anchor="w")

    choice = tk.StringVar(value="")
    ttk.Radiobutton(body, text="已經算進去了（基準會往回推一天份）",
                    variable=choice, value="applied").pack(anchor="w", pady=(8, 0))
    ttk.Radiobutton(body, text="還沒算，請程式幫我扣",
                    variable=choice, value="pending").pack(anchor="w")

    result = {}

    def confirm():
        if not choice.get():
            messagebox.showwarning("還沒選", "請先選一個，程式不會替你猜。", parent=win)
            return
        result["value"] = choice.get() == "applied"
        win.destroy()

    buttons = ttk.Frame(body)
    buttons.pack(fill="x", pady=(16, 0))
    ttk.Button(buttons, text="取消", command=win.destroy).pack(side="right")
    ttk.Button(buttons, text="確定校正", command=confirm).pack(side="right", padx=(0, 8))

    win.grab_set()
    parent.wait_window(win)
    return result.get("value")


def ask_backfill(parent, family, today):
    """問補登的日期與金額。回傳 (date, float)，取消是 None。"""
    win = tk.Toplevel(parent)
    win.title("補登淨收付")
    win.transient(parent)
    win.resizable(False, False)

    body = ttk.Frame(win, padding=16)
    body.pack(fill="both", expand=True)

    ttk.Label(body, text="補登某一天漏掉的淨收付。\n"
                         "金額可以去券商的對帳單查詢頁查那天的「淨收付合計」。",
              font=(family, 10)).pack(anchor="w")

    form = ttk.Frame(body)
    form.pack(fill="x", pady=(12, 0))
    ttk.Label(form, text="日期").grid(row=0, column=0, sticky="w")
    date_entry = ttk.Entry(form, font=(family, 10), width=16)
    date_entry.insert(0, today.isoformat())
    date_entry.grid(row=0, column=1, sticky="w", padx=(8, 0))
    ttk.Label(form, text="YYYY-MM-DD", style="Hint.TLabel").grid(row=0, column=2, sticky="w", padx=(8, 0))

    ttk.Label(form, text="金額").grid(row=1, column=0, sticky="w", pady=(8, 0))
    amount_entry = ttk.Entry(form, font=(family, 10), width=16)
    amount_entry.grid(row=1, column=1, sticky="w", padx=(8, 0), pady=(8, 0))
    ttk.Label(form, text="買進是負數，賣出是正數", style="Hint.TLabel").grid(
        row=1, column=2, sticky="w", padx=(8, 0), pady=(8, 0))

    result = {}

    def confirm():
        try:
            day = datetime.date.fromisoformat(date_entry.get().strip())
        except ValueError:
            messagebox.showwarning("日期看不懂", "請照 2026-08-17 這種格式填。", parent=win)
            return
        try:
            amount = float(amount_entry.get().strip())
        except ValueError:
            messagebox.showwarning("金額看不懂", "請填數字，例如 -107。", parent=win)
            return
        result["value"] = (day, amount)
        win.destroy()

    buttons = ttk.Frame(body)
    buttons.pack(fill="x", pady=(16, 0))
    ttk.Button(buttons, text="取消", command=win.destroy).pack(side="right")
    ttk.Button(buttons, text="補登", command=confirm).pack(side="right", padx=(0, 8))

    win.grab_set()
    amount_entry.focus_set()
    parent.wait_window(win)
    return result.get("value")


def main():
    root = tk.Tk()
    app = SyncApp(root)

    def on_close():
        if app.busy and not messagebox.askokcancel(
            "還在忙", "背景還在登入或寫入。現在關掉可能會留下寫到一半的狀態，確定要關嗎？"
        ):
            return
        root.destroy()

    root.protocol("WM_DELETE_WINDOW", on_close)
    root.mainloop()


if __name__ == "__main__":
    main()
