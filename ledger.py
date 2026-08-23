"""
程式自己的記帳本：現金的基準與逐日流水。

Excel 完全保持原本的版面，所有輔助資訊都放在這裡，檔案就放在 Excel 旁邊。

紀錄檔不見了不會算錯帳，股數/成本/現金下次讀取或登入都會照 Excel 現值重算，
現金基準會被當成今天第一次登入重新設定一次。

股數與成本：一律覆蓋
--------------------
網頁庫存是唯一真相，程式每次都直接把算出來的值寫進 Excel，不記、不比對
Excel 上原本是什麼。修改 Excel 的風險交給操作的人自己管控。

現金：每天重新起算，不回頭算舊帳
--------------------------------
網頁只回答「今天的淨收付是多少」，從來不告訴你「你現在有多少現金」，
所以餘額要有一個起點才算得出來：

    應有餘額 = baseline_value + 今天的淨收付

baseline_value 就是介面上那個「今日初始現金餘額」—— 今天第一次登入時 Excel 上的
B8。那個時間點今天要買賣什麼都還沒發生，所以它必定是今天的起點
（見 planner.initialize）。

跨日累加的版本試過：基準只設一次，applied 把每天的淨收付一路往後疊。它的代價是
漏跑一天就永遠少一段，而且要補得回頭一天一天對。改成每天重設之後，昨天有沒有跑過、
跑得對不對都不影響今天 —— 今天的起點就是今天看到的 B8，沒有舊帳要算。

applied 仍然是 {日期: 金額} 的字典，只是裡面永遠只有今天一筆。留著它是因為
同一天重跑只是覆蓋同一個 key，跑幾次結果都一樣。

同一天內想改餘額
----------------
基準每天由「當天第一次登入」自己設好，正常情況沒有人要碰它。會碰到的是當天
第二次以後登入：基準已經設過、不會再跟著 B8 走，這時候手改 B8 下次寫入會被
程式直接蓋掉（餘額是 baseline + 今天淨收付算出來的，不是抄 B8）。要改就直接
改介面上那個「今日初始現金餘額」，走的還是 calibrate，只是 balance 由人給
（見 planner.apply_cash_reset）。
"""

import datetime
import json

LEDGER_SUFFIX = "-同步紀錄.json"
HISTORY_SUFFIX = "-同步歷程.jsonl"
HISTORY_KEEP_DAYS = 10

# 檔名跟著 Excel 的檔名走，而不是固定叫「持股同步紀錄.json」。
# 萬一同一個資料夾裡有第二份持股表，兩者的現金基準混在一起是災難級的錯誤，
# 而且不會有任何徵兆 —— 用檔名綁死就沒有這個破口。


def new_cash():
    """現金的狀態：基準與流水。"""
    return {
        "baseline_date": None,
        "baseline_value": None,
        "applied": {},
        "last_written": None,
        "last_written_at": None,
    }


def new_sheet():
    return {"account_code": "", "cash": new_cash()}


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
                # 讀不到就報錯給人看，不要嘗試修補一個壞掉的帳本，
                # 那比沒有帳本更危險。
                raise RuntimeError(
                    f"紀錄檔讀不出來：{self.path}\n{exc}\n"
                    f"可以把它改名備份起來，程式會重新建立（現金基準下次登入會重設）。"
                ) from exc

        self.prune_history()

    def sheet(self, name):
        """取得某個分頁的紀錄，沒有就建一個空的。"""
        sheets = self.data.setdefault("sheets", {})
        book = sheets.setdefault(name, new_sheet())
        book.setdefault("cash", new_cash())
        return book

    def setting(self, key, default=None):
        """介面用的偏好設定。跟著這份 Excel 走，不是全域的。"""
        return self.data.setdefault("settings", {}).get(key, default)

    def set_setting(self, key, value):
        self.data.setdefault("settings", {})[key] = value
        self.save()

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

    def prune_history(self, keep_days=HISTORY_KEEP_DAYS):
        """
        只留最近 keep_days 天的歷程，超過的直接丟掉，不進備份資料夾。

        跟 clear_history 不一樣：那是使用者按按鈕、整批收進「備份」；這是開檔
        時自動做的保養，本來就打算讓舊的消失，不留副本才叫「自動清除」。
        解析不出日期的行（舊格式、手改壞掉）保留，寧可多留不砍錯。
        """
        if not self.history_path.is_file():
            return

        cutoff = (datetime.date.today()
                  - datetime.timedelta(days=keep_days)).isoformat()
        kept = []
        dropped = False
        for line in self.history_path.read_text(encoding="utf-8").splitlines():
            stripped = line.strip()
            if not stripped:
                continue
            try:
                event = json.loads(stripped)
                at = event.get("at") or ""
            except json.JSONDecodeError:
                at = ""
            if len(at) >= 10 and at[:10] < cutoff:
                dropped = True
                continue
            kept.append(stripped)

        if not dropped:
            return

        text = "".join(line + "\n" for line in kept)
        temp = self.history_path.with_name(self.history_path.name + ".tmp")
        temp.write_text(text, encoding="utf-8")
        temp.replace(self.history_path)

    def clear_history(self):
        """
        清掉歷程。回傳收起來的檔案位置，本來就沒有歷程就回傳 None。

        檔案不刪，改名收進「備份」資料夾。歷程是拿來對帳的東西 ——「那天那一格
        是誰改的」問不出來的代價，遠大於多留一個檔案；按錯一顆按鈕就永久失去
        追溯能力，那顆按鈕不該存在。

        只動歷程，不動 self.data：現金基準跟流水都記在紀錄檔那邊，
        清歷程等於撕掉日記本，不是把帳算掉。
        """
        if not self.history_path.is_file():
            return None
        folder = self.history_path.parent / "備份"
        folder.mkdir(exist_ok=True)
        stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        dest = folder / f"{self.history_path.stem}_{stamp}{self.history_path.suffix}"
        self.history_path.replace(dest)
        return dest


def cash_after(cash, day, net):
    """如果把某一天的淨收付記成 net，餘額會變成多少。不會改到 cash 本身。"""
    base = cash.get("baseline_value")
    if base is None:
        return None
    applied = dict(cash.get("applied") or {})
    applied[day.isoformat()] = round(net, 2)
    return round(base + sum(applied.values()), 2)


def opening_balance(cash):
    """
    今日初始現金餘額 —— 今天第一次登入時 Excel 上的 B8，餘額就是從它加起來的。

    它就是 baseline_value 本身。基準每天重設一次（見 planner.initialize），
    所以不必把前幾天的流水加回來 —— 這支程式不回頭算舊帳。
    """
    return cash.get("baseline_value")


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
    cash["last_written"] = round(balance, 2)
    cash["last_written_at"] = at
