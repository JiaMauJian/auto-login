"""歷程分頁：篩選、顯示、清除。"""

import datetime
import json

from tkinter import messagebox

from util import show, to_num
from ui_common import ALL_CHOICE, ask_confirm, stock_title, within

def item_order(label):
    """
    項目選單的排序：現金排最前面，其餘照股票分組，同一檔的股數排在成本前面。

    純照字串排會變成「所有股數」一段、「所有成本」另一段，同一檔股票的兩格
    隔了幾十列 —— 但人是先想到哪一檔股票，才想到要看股數還是成本。
    """
    if label.startswith("現金"):
        return (0, "", 0)
    inside = label.partition("（")[2].rstrip("）")
    return (1, inside, 0 if label.startswith("股數") else 1)


def describe_change(old, new):
    """
    歷程那一欄要寫什麼：「舊值 → 新值」。

    2026/08/24 使用者要求簡化：原本「交接」（`by == "adopt"`）印成
    「1 → 改記 0」，用「改記」點明變的是程式記憶、不是 Excel 格子——
    但那句話本身就要多想一下，拿掉之後跟其他行同一個形狀，
    是不是交接已經有 [今日初始餘額] 標籤跟後面的 note 講了。
    """
    return f"{show(old)} → {show(new)}"


def history_line(event):
    """
    一行印完一筆歷程事件（時間／交易人／項目／變化／說明），取代原本
    Treeview 的六個欄位（2026/08/23 改用 Text，理由跟更新分頁訊息框一樣：
    欄寬固定會切字，改成單行文字自動換行就不會切到，見
    ui_layout._build_history_tab）。來源（程式／人工／交接）不再特別強調，
    2026/08/23 使用者要求拿掉；2026/08/24 連 describe_change 裡「交接」
    印成「改記」那條特例也拿掉了，`by` 不再需要傳進去。

    回傳 (文字, tag)：現金餘額變負的那一筆標紅（"neg"），跟更新頁訊息框、
    左邊名單同一套「負現金就紅」的規矩（見 ui_sync._cash_line、
    ui_layout.py 的 "negative" tag），2026/08/23 使用者要求加上。
    """
    time_text = (event.get("at") or "").replace("T", " ")
    change = describe_change(event.get("old"), event.get("new"))
    note = event.get("note", "")
    line = f"{time_text}　{event.get('sheet', '')}　[{event.get('label', '')}]　{change}"
    if note:
        line += f"　{note}"
    new = to_num(event.get("new"), None)
    negative = event.get("label") == "現金餘額" and new is not None and new < 0
    return line, ("neg" if negative else None)


def _merge_stock_group(events):
    """
    同一輪、同一檔股票的股數／成本併成一行（跟更新分頁訊息框
    ui_sync._stock_lines 同一個併法，見 stock_title），但歷程要留舊→新方便
    跨日查核，不像訊息框那樣省掉。
    """
    titles, order = {}, []
    for event in events:
        title = stock_title(event.get("label", ""))
        if title not in titles:
            titles[title] = {}
            order.append(title)
        which = "股數" if event.get("label", "").startswith("股數") else "成本"
        titles[title][which] = event

    lines = []
    for title in order:
        parts = titles[title]
        first = next(iter(parts.values()))
        time_text = (first.get("at") or "").replace("T", " ")
        change_text = "　".join(
            f"{key} {describe_change(parts[key].get('old'), parts[key].get('new'))}"
            for key in ("股數", "成本") if key in parts)
        lines.append((f"{time_text}　{first.get('sheet', '')}　[{title}]　{change_text}", None))
    return lines


def _grouped_lines(rows):
    """
    把 `rows`（依檔案順序、時間由舊到新）依 (交易人, 時間) 分組，同一輪的
    股數／成本併成一行，其餘照舊一行一筆；輸出新的一組排最上面，跟原本
    「最新的放最上面」的慣例一致，組內維持現金先、股票後（跟更新分頁訊息框
    同一個順序，見 docs/更新分頁訊息框改版.md）。
    """
    groups, order = {}, []
    for row in rows:
        key = (row.get("sheet", ""), row.get("at", ""))
        if key not in groups:
            groups[key] = []
            order.append(key)
        groups[key].append(row)

    lines = []
    for key in reversed(order):
        stock_events, other_events = [], []
        for event in groups[key]:
            target = stock_events if event.get("label", "").startswith(("股數（", "成本（")) else other_events
            target.append(event)
        lines.extend(history_line(event) for event in other_events)
        lines.extend(_merge_stock_group(stock_events))
    return lines


class UiHistoryMixin:
    # ---------- 歷程分頁 ----------

    def refresh_history(self):
        """重讀歷程檔。篩選是在記憶體裡做的，換選單不會再碰一次硬碟。"""
        path = self.ledger.history_path if self.ledger else None
        self.history_file = path.name if path else ""
        self.history_rows = []

        if path and path.is_file():
            for line in path.read_text(encoding="utf-8").splitlines():
                line = line.strip()
                if not line:
                    continue
                try:
                    self.history_rows.append(json.loads(line))
                except json.JSONDecodeError:
                    continue

        self._refresh_history_choices()
        self._fill_history()

    def _refresh_history_choices(self):
        """
        照現有的歷程重建兩個選單。

        項目只列出「選中的那個交易人真的有的」—— 把 20 個帳號上百格全部攤平在
        同一個下拉選單裡，跟沒有篩選是一樣的。原本選的那一項還在就留著，
        不在了就退回「全部」：選單上的字還在、表格卻是空的，是最難懂的那種畫面。
        """
        who = self.history_who.get() or ALL_CHOICE
        names = sorted({row.get("sheet", "") for row in self.history_rows if row.get("sheet")})
        self.history_who.configure(values=[ALL_CHOICE] + names)
        if who not in names:
            who = ALL_CHOICE
        self.history_who.set(who)

        item = self.history_item.get() or ALL_CHOICE
        labels = sorted({row.get("label", "") for row in self.history_rows
                         if row.get("label") and (who == ALL_CHOICE or row.get("sheet") == who)},
                        key=item_order)
        self.history_item.configure(values=[ALL_CHOICE] + labels)
        if item not in labels:
            item = ALL_CHOICE
        self.history_item.set(item)

    def _on_history_who(self, _event):
        # 換人之後項目清單要跟著換，不然會停在一個那個人根本沒有的項目上。
        self._refresh_history_choices()
        self._fill_history()

    def _fill_history(self):
        """把通過篩選的事件畫進訊息框，一行一筆，最新的放最上面。"""
        self.history_box.configure(state="normal")
        self.history_box.delete("1.0", "end")
        self._sync_clear_button()

        if not self.history_rows:
            self.history_box.configure(state="disabled")
            return

        who, item, when = (self.history_who.get(), self.history_item.get(),
                           self.history_when.get())
        # 今天是現算的，不是開程式那一刻的 self.today —— 這支程式常常開著過夜，
        # 跨過午夜之後「今天」還停在昨天的話，剛跑完的那一批會整批不見。
        today = datetime.date.today()
        shown = [row for row in self.history_rows
                 if (who == ALL_CHOICE or row.get("sheet") == who)
                 and (item == ALL_CHOICE or row.get("label") == item)
                 and within(row.get("at"), when, today)]

        for index, (text, tag) in enumerate(_grouped_lines(shown)):   # 最新的放最上面
            if index:
                self.history_box.insert("end", "\n")
            self.history_box.insert("end", text, (tag,) if tag else ())
        self.history_box.configure(state="disabled")

        counted = (f"{len(shown)} 筆" if len(shown) == len(self.history_rows)
                   else f"篩出 {len(shown)} 筆／共 {len(self.history_rows)} 筆")
        self.history_hint.configure(text=f"{counted}，最新的在最上面。檔案：{self.history_file}")

    def _sync_clear_button(self):
        """沒有歷程可清、或正在跑的時候，「清除歷程」不要亮著。"""
        usable = bool(self.history_rows) and not self.busy
        self.history_clear.configure(state="normal" if usable else "disabled")

    def clear_history(self):
        """
        把歷程檔刪掉，畫面清空。

        清掉的只有「誰在什麼時候改了哪一格」這本日記，不會動到紀錄檔 ——
        每一格歸誰管、現金的基準與流水都在那邊，所以清完再同步，
        算出來的提案跟清之前一模一樣。這句話也要講給使用者聽：
        「清除」兩個字很容易被讀成「把帳歸零」，那是最貴的誤會。
        """
        if not self.ledger or not self.history_rows:
            return

        if not ask_confirm(
                self.root,
                "清除歷程",
                f"要清掉全部 {len(self.history_rows)} 筆歷程嗎？這會直接刪掉，不會留備份。\n\n"
                "每一格歸誰管、現金的基準都記在另一個紀錄檔裡，不受影響。",
                confirm_style="primary"):
            return

        try:
            self.ledger.clear_history()
        except OSError as exc:
            messagebox.showerror("清不掉", f"歷程檔可能正被別的程式開著：\n{exc}")
            return

        self.refresh_history()
        self._say("歷程已清空。")
