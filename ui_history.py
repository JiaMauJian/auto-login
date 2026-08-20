"""歷程分頁：篩選、顯示、清除。"""

import datetime
import json

from tkinter import messagebox

from util import show
from ui_common import ALL_CHOICE, WHEN_TODAY

# 「補登」這個來源已經沒有入口了（現金帳本分頁拿掉時一起收掉），
# 但舊的歷程檔裡還有這種紀錄，名字要留著才不會顯示成一個英文代號。
SOURCE_NAMES = {"program": "程式", "human": "人工", "adopt": "交接", "backfill": "補登"}


def within(stamp, when, today):
    """這一筆的時間在不在選的期間裡。"""
    if when == ALL_CHOICE:
        return True
    try:
        day = datetime.date.fromisoformat((stamp or "")[:10])
    except ValueError:
        # 時間壞掉的那幾筆只有「全部」看得到。當成今天會更糟 ——
        # 那等於憑空生出一筆今天的異動，而歷程是拿來對帳的東西。
        return False
    if when == WHEN_TODAY:
        return day == today
    return (today - day).days < 7      # 含今天往回數七天


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


def describe_change(by, old, new):
    """
    歷程那一欄要寫什麼。

    「交接」的新舊值講的是**程式的記憶**，不是格子的值：交接就是把
    「我記得這格是多少」改成 Excel 上現在的數字，Excel 本身一格都沒動。
    如果照 program 那樣印成「1 → 0」，看起來會像程式把數字改壞了，
    所以新值那邊寫「改記」，點明變的是記憶。
    """
    if by == "adopt":
        return f"{show(old)} → 改記 {show(new)}"
    if by == "human":
        return f"{show(old)} → {show(new)}（人工）"
    return f"{show(old)} → {show(new)}"


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
        """把通過篩選的事件畫進表格。"""
        self.history_tree.delete(*self.history_tree.get_children())
        self._sync_clear_button()

        if not self.history_rows:
            self.history_hint.configure(text="還沒有任何歷程。第一次寫入或交接之後就會有。")
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

        for event in reversed(shown):         # 最新的放最上面
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

        counted = (f"{len(shown)} 筆" if len(shown) == len(self.history_rows)
                   else f"篩出 {len(shown)} 筆／共 {len(self.history_rows)} 筆")
        self.history_hint.configure(text=f"{counted}，最新的在最上面。檔案：{self.history_file}")

    def _sync_clear_button(self):
        """沒有歷程可清、或正在跑的時候，「清除歷程」不要亮著。"""
        usable = bool(self.history_rows) and not self.busy
        self.history_clear.configure(state="normal" if usable else "disabled")

    def clear_history(self):
        """
        把歷程收進「備份」資料夾，畫面清空。

        清掉的只有「誰在什麼時候改了哪一格」這本日記，不會動到紀錄檔 ——
        每一格歸誰管、現金的基準與流水都在那邊，所以清完再同步，
        算出來的提案跟清之前一模一樣。這句話也要講給使用者聽：
        「清除」兩個字很容易被讀成「把帳歸零」，那是最貴的誤會。
        """
        if not self.ledger or not self.history_rows:
            return

        if not messagebox.askyesno(
                "清除歷程",
                f"要清掉全部 {len(self.history_rows)} 筆歷程嗎？\n\n"
                "舊的歷程檔會改名收進「備份」資料夾，不是真的刪掉。\n"
                "每一格歸誰管、現金的基準都記在另一個紀錄檔裡，不受影響。",
                icon="warning", default="no", parent=self.root):
            return

        try:
            saved = self.ledger.clear_history()
        except OSError as exc:
            messagebox.showerror("清不掉", f"歷程檔可能正被別的程式開著：\n{exc}")
            return

        self.refresh_history()
        # 收去哪裡要寫在狀態列上，不然「備份」兩個字等於白講 —— 人找不到的備份
        # 跟沒有備份是一樣的。
        self._say(f"歷程已清空。舊的收在 {saved}" if saved else "歷程已清空。")
