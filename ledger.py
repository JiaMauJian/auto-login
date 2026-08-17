"""
程式自己的記帳本：記住「我上次寫進 Excel 的是什麼」，以及現金的逐日流水。

Excel 完全保持原本的版面，所有輔助資訊都放在這裡，檔案就放在 Excel 旁邊。

紀錄檔不見了不會算錯帳，只會變成「每一格都不認得」—— 全部當成人工維護、
一格都不敢寫，然後請使用者重新設定基準。壞掉的方向是安全的那一邊。

誰在管這一格
------------
每一格各自有狀態，不是整個分頁一起切。判定規則只有一條：

    Excel 現值 == 紀錄檔的 last_written   ->  自動（程式還管著，可以覆蓋）
    Excel 現值 != 紀錄檔的 last_written   ->  手動（有人改過，程式不准碰）
    紀錄檔裡沒有這一格                     ->  未接管（要人明確交給程式）

自動轉手動由程式自己偵測、悄悄發生，目的是保護使用者的手改；
手動轉自動一定要人明確要求（CLI 的 --adopt、介面上的「交還給程式」按鈕）。

現金為什麼要記流水
------------------
網頁只回答「今天的淨收付是多少」，從來不告訴你「你現在有多少現金」，
所以餘額只能從某個基準往後累加：

    應有餘額 = baseline_value + applied 裡所有日期的淨收付總和

applied 是 {日期: 金額} 的字典而不是一個累計數字，這帶來兩件事：
同一天重跑只是覆蓋同一個 key，跑幾次結果都一樣；
漏跑的那天會在字典裡缺一個 key，看得見、也補得回來 —— 舊的「前日餘額 +
今日淨收付」模型漏掉一天就永遠追不回來，這是換成流水的主因。

baseline_value 的定義是「起算日當天，套用 applied 之前的餘額」。
校正時如果使用者說「Excel 上這個數字已經含今天的淨收付了」，
基準就要往回推一天份（balance - net），這樣加回來才會等於畫面上的數字。
"""

import datetime
import json

from util import values_match

AUTO, MANUAL, UNTRACKED = "auto", "manual", "untracked"

LEDGER_SUFFIX = "-同步紀錄.json"
HISTORY_SUFFIX = "-同步歷程.jsonl"

# 檔名跟著 Excel 的檔名走，而不是固定叫「持股同步紀錄.json」。
# 萬一同一個資料夾裡有第二份持股表，兩者的現金基準混在一起是災難級的錯誤，
# 而且不會有任何徵兆 —— 用檔名綁死就沒有這個破口。


def new_field():
    """一格（股數或成本）的狀態。"""
    return {"mode": UNTRACKED, "last_written": None, "last_written_at": None}


def new_cash():
    """現金的狀態。比一般格子多了基準與流水。"""
    return {
        "mode": UNTRACKED,
        "baseline_date": None,
        "baseline_value": None,
        "applied": {},
        "last_written": None,
        "last_written_at": None,
    }


def new_sheet():
    return {"account_code": "", "cash": new_cash(), "holdings": {}}


class Ledger:
    """紀錄檔的讀寫。資料就是一個 dict，需要的人直接改，改完呼叫 save()。"""

    def __init__(self, excel_path):
        self.path = excel_path.parent / (excel_path.stem + LEDGER_SUFFIX)
        self.history_path = excel_path.parent / (excel_path.stem + HISTORY_SUFFIX)
        self.existed = self.path.is_file()
        self.data = {"version": 1, "sheets": {}}

        if self.existed:
            try:
                self.data = json.loads(self.path.read_text(encoding="utf-8"))
            except (OSError, json.JSONDecodeError) as exc:
                # 讀不到就當作沒有紀錄，全部退回「未接管」。
                # 不要嘗試修補一個壞掉的帳本，那比沒有帳本更危險。
                raise RuntimeError(
                    f"紀錄檔讀不出來：{self.path}\n{exc}\n"
                    f"可以把它改名備份起來，程式會重新建立（所有格子要重新交接一次）。"
                ) from exc

    def sheet(self, name):
        """取得某個分頁的紀錄，沒有就建一個空的。"""
        sheets = self.data.setdefault("sheets", {})
        book = sheets.setdefault(name, new_sheet())
        book.setdefault("cash", new_cash())
        book.setdefault("holdings", {})
        return book

    def field(self, book, code, which):
        """取得某一檔股票的 qty / cost 狀態，沒有就建一個。"""
        holding = book["holdings"].setdefault(code, {})
        return holding.setdefault(which, new_field())

    def save(self):
        """先寫暫存檔再換掉本尊，避免寫到一半斷電留下一個殘缺的帳本。"""
        text = json.dumps(self.data, ensure_ascii=False, indent=2)
        temp = self.path.with_name(self.path.name + ".tmp")
        temp.write_text(text, encoding="utf-8")
        temp.replace(self.path)
        self.existed = True

    def append_history(self, events):
        """歷程只增不改，一行一筆 JSON。"""
        if not events:
            return
        with self.history_path.open("a", encoding="utf-8") as handle:
            for event in events:
                handle.write(json.dumps(event, ensure_ascii=False) + "\n")


def status_of(state, current):
    """
    這一格現在誰在管。current 是 Excel 上的現值。

    以 mode 為準而不是看 last_written 有沒有值：交接一個空白格
    （例如新加的一列還沒填股數）也要算成已接管，否則它會永遠停在未接管。
    """
    mode = (state or {}).get("mode", UNTRACKED)
    if mode == UNTRACKED:
        return UNTRACKED
    if mode == MANUAL:
        return MANUAL
    if not values_match(state.get("last_written"), current):
        return MANUAL          # 這次才偵測到的手改
    return AUTO


def mark_written(state, value, at):
    """程式寫進去之後，把「我記得寫了什麼」更新掉。"""
    state["mode"] = AUTO
    state["last_written"] = value
    state["last_written_at"] = at
    state.pop("since", None)


def mark_manual(state, at):
    """偵測到人工改動，轉為手動並記下時間（是偵測到的時間，不是實際改動的時間）。"""
    state["mode"] = MANUAL
    state["since"] = at


def expected_cash(cash):
    """依基準與流水算出「現金餘額應該是多少」。沒有基準就回 None。"""
    base = cash.get("baseline_value")
    if base is None:
        return None
    return round(base + sum((cash.get("applied") or {}).values()), 2)


def cash_after(cash, day, net):
    """如果把某一天的淨收付記成 net，餘額會變成多少。不會改到 cash 本身。"""
    base = cash.get("baseline_value")
    if base is None:
        return None
    applied = dict(cash.get("applied") or {})
    applied[day.isoformat()] = round(net, 2)
    return round(base + sum(applied.values()), 2)


def record_net(cash, day, net):
    """把某一天的淨收付記進流水。同一天重記就是覆蓋，所以重跑不會重複扣。"""
    cash.setdefault("applied", {})[day.isoformat()] = round(net, 2)


def calibrate(cash, balance, day, net, today_included, at):
    """
    以 Excel 現在的數字為新基準，從 day 起重新起算。

    today_included=True  代表 balance 已經含了今天的 net，基準要往回推一天份。
    today_included=False 代表還沒含，程式待會兒會把 net 加上去。

    舊的 applied 全部丟掉：既然基準重設在今天，之前的流水已經被 balance 這個
    數字吸收掉了，留著會被重複加一次。
    """
    cash["baseline_date"] = day.isoformat()
    cash["baseline_value"] = round(balance - net, 2) if today_included else round(balance, 2)
    cash["applied"] = {day.isoformat(): round(net, 2)}
    cash["mode"] = AUTO
    cash["last_written"] = round(balance, 2)
    cash["last_written_at"] = at
    cash.pop("since", None)


def missing_dates(cash, today):
    """
    基準日到今天之間，流水裡沒有紀錄的工作日。

    只跳週末，不管國定假日，所以會有假警報 —— 但方向是對的：
    寧可多提醒幾天讓人去看一眼，也不要漏掉真的沒跑到的那天。
    """
    raw = cash.get("baseline_date")
    if not raw:
        return []
    try:
        start = datetime.date.fromisoformat(raw)
    except ValueError:
        return []

    applied = cash.get("applied") or {}
    missing = []
    day = start
    while day < today:
        if day.weekday() < 5 and day.isoformat() not in applied:
            missing.append(day)
        day += datetime.timedelta(days=1)
    return missing
