"""
把「網頁資料 × Excel 現值 × 紀錄檔」算成一張變更提案清單。

這裡不碰 Excel、不碰網路，只有純計算，所以命令列跟之後的介面可以共用
同一份判斷邏輯，也可以拿假資料直接測。

提案是一格一列（E5 跟 F5 分開），因為自動/手動是一格一格判定的 ——
同一檔股票的股數還在自動、成本被手改成手動，是完全可能的狀態。
畫面上要把兩格併成一列顯示是顯示層的事。
"""

import ledger
from excel_io import COL_COST, COL_QTY, CELL_BALANCE
from recon import TRADE_NAMES
from util import cell_name, same_number, show, to_num

# 網頁沒有這一檔，不是「誰在管」的問題，但畫面上要跟自動/手動並排顯示，
# 所以放在同一組狀態值裡。
WEB_MISSING = "web_missing"

STATUS_NAMES = {
    ledger.AUTO: "自動",
    ledger.MANUAL: "手動",
    ledger.UNTRACKED: "未接管",
    WEB_MISSING: "網頁沒有",
}


def merge_holdings(arrays):
    """
    把未實現損益整理成 {股號: {股數, 成本, 名稱}}。

    網頁的資料是以「股票+交易別」為單位，同一檔可能佔好幾列（現股一列、融資一列），
    但 Excel 一檔只有一列，所以要合併：股數直接相加，成本用股數加權平均
    —— 直接取其中一列的均價會是錯的。
    另外照網頁自己的規則濾掉當沖與成本股數 0 的列，否則加總會跟畫面對不起來。
    """
    merged = {}
    notes = []

    for item in arrays:
        stkno = (item.get("stkno") or "").strip()
        if not stkno:
            continue
        trade = str(item.get("trade") or "")
        if trade in ("9", "A"):
            continue
        qty = to_num(item.get("costqtyn"))
        if qty == 0:
            continue

        record = merged.setdefault(stkno, {"qty": 0.0, "amount": 0.0, "name": "", "trades": []})
        record["qty"] += qty
        record["amount"] += qty * to_num(item.get("priceavgn"))
        record["name"] = record["name"] or (item.get("stkna") or "").strip()
        record["trades"].append(TRADE_NAMES.get(trade, trade))

    for stkno, record in merged.items():
        record["cost"] = round(record["amount"] / record["qty"], 4) if record["qty"] else 0.0
        if len(record["trades"]) > 1:
            notes.append(
                f"{stkno} {record['name']} 在網頁上佔了 {len(record['trades'])} 列"
                f"（{'、'.join(record['trades'])}），已用股數加權平均合併成本"
            )

    return merged, notes


def settlement_total(arrays):
    """當日淨收付合計。跳過成交股數 0 的列，跟網頁畫面的算法一致。"""
    net = 0.0
    rows = 0
    for item in arrays:
        for mat in item.get("matdat") or []:
            if str(mat.get("qty") or "").strip() == "0":
                continue
            rows += 1
            net += to_num(mat.get("payn"))
    return net, rows


def traded_today(pnl_arrays, today):
    """未實現損益的持股明細裡有沒有「今天」成交的紀錄。用來交叉檢查淨收付是不是漏了。"""
    stamp = today.strftime("%Y%m%d")
    for item in pnl_arrays:
        for detail in item.get("stkdat") or []:
            if str(detail.get("tdate") or "").strip() == stamp:
                return True
    return False


def plan(sheet_data, record, book, today):
    """
    算出這個分頁的提案。回傳 (提案清單, 提醒清單)。

    純讀取，不會動到 book —— 要落實到紀錄檔是 commit() 的事，
    這樣試算模式才能保證真的什麼都沒改到。
    """
    proposals = []
    warnings = []

    holdings, notes = merge_holdings(record.get("未實現損益", []))
    warnings.extend(notes)

    seen = set()
    for line in sheet_data["rows"]:
        code = line["code"]
        seen.add(code)
        found = holdings.get(code)

        if found is None:
            # 網頁已經沒有這檔了。刻意不自動歸零 —— 使用者的流程是「先刪 Excel 再賣」，
            # 所以這通常代表忘了刪，該由人確認而不是程式清掉。
            warnings.append(
                f"第 {line['row']} 列「{line['label']}」在網頁庫存中已不存在，"
                f"股數與成本維持原樣，請確認是否忘記刪除這一列"
            )
            for which, col, current in (("qty", COL_QTY, line["qty"]),
                                        ("cost", COL_COST, line["cost"])):
                proposals.append(_row(line["row"], col, "holding", code, which,
                                      f"{'股數' if which == 'qty' else '成本'}（{line['label']}）",
                                      current, None, None, WEB_MISSING, "網頁庫存已無此檔"))
            continue

        label = f"{code} {found['name']}"
        for which, col, current, web in (
            ("qty", COL_QTY, line["qty"], int(found["qty"])),
            ("cost", COL_COST, line["cost"], found["cost"]),
        ):
            state = (book["holdings"].get(code) or {}).get(which)
            status = ledger.status_of(state, current)
            note = ""
            if status == ledger.MANUAL:
                note = f"手動改過（程式記得 {state.get('last_written')}），不會覆蓋"
            elif status == ledger.UNTRACKED:
                note = "尚未交給程式管理"
            proposals.append(_row(
                line["row"], col, "holding", code, which,
                f"{'股數' if which == 'qty' else '成本'}（{label}）",
                current, web, web, status, note,
            ))

    for code, found in holdings.items():
        if code not in seen:
            warnings.append(
                f"網頁有 {code} {found['name']}（{int(found['qty'])} 股）但 Excel 沒有這一列，"
                f"不會自動新增，請自行決定放在哪一列"
            )

    proposals.append(_cash(sheet_data, record, book, today, warnings))
    return proposals, warnings


def _row(row, col, kind, key, which, label, current, web, proposed, status, note):
    """組一列提案。will_write 只有「自動」而且值真的不一樣時才會是 True。"""
    will_write = (
        status == ledger.AUTO
        and proposed is not None
        and not same_number(current, proposed)
    )
    return {
        "row": row, "col": col, "cell": cell_name(row, col),
        "kind": kind, "key": key, "which": which, "label": label,
        "current": current, "web": web, "proposed": proposed,
        "status": status, "note": note, "will_write": will_write,
        "reset_to": None,
    }


def apply_cash_reset(item, opening):
    """
    把人填的「今天開盤前的現金」套進現金那一列。就地改，回傳同一個 item。

    只能在 plan() 之後叫得動 —— 它要用今日淨收付，而那是網頁資料，
    登入的當下還沒去查（login_only 只登入、不抓資料）。這也正是它值得晚一步
    才問的原因：算得出結果，就不必請人回答「你填的數字含不含今天的成交」，
    那個問題每次都要人回想今天做過什麼，答錯的代價剛好是一整天的淨收付。

    這是唯一能蓋過「手動」的東西。人已經明講了正確的數字，
    程式沒有理由再守著一個它自己也知道過時的值。
    """
    target = round(opening + item["net"], 2)
    item["proposed"] = target
    item["status"] = ledger.AUTO
    item["reset_to"] = target
    item["record_net"] = False          # calibrate 會把流水一起寫好
    item["will_write"] = not same_number(item["current"], target)
    item["note"] = (f"開盤前 {show(opening)} + 今日淨收付 {show(item['net'])}"
                    f"（{item['net_rows']} 筆成交），從今天起重新起算")
    return item


def _cash(sheet_data, record, book, today, warnings):
    """
    現金那一列。跟持股不同的地方全部集中在這裡。

    現金是累加出來的：網頁只說「今天淨收付 -107」，不會說「你現在有多少錢」。
    所以它的提議值不是網頁值，而是「基準 + 流水」算出來的應有餘額。
    """
    row, col = CELL_BALANCE
    balance = sheet_data["balance"]
    cash = book["cash"]
    net, rows = settlement_total(record.get("當日淨收付", []))

    proposal = _row(row, col, "cash", "cash", "balance", "現金餘額",
                    balance, net, None, ledger.status_of(cash, balance), "")
    proposal["net"] = net
    proposal["net_rows"] = rows
    proposal["record_net"] = False
    # 淨收付本身信不過的時候立起來。這一格連「重設基準」都不該讓人按 ——
    # 拿一個已知是錯的 net 去 calibrate，等於把今天的成交永久算進基準裡。
    proposal["blocked"] = False

    if balance is None:
        proposal["note"] = "B8 是空的或不是數字，請先填一個現金餘額"
        return proposal

    # 交叉檢查：未實現損益說今天有成交，淨收付卻是 0，兩邊矛盾。
    # 最可能的原因是收盤結帳後查不到當日資料了，這時候記 0 會讓餘額默默漏掉今天。
    # 只擋現金這一格，持股照樣可以寫。
    if net == 0 and traded_today(record.get("未實現損益", []), today):
        proposal["note"] = "未實現損益顯示今天有成交、淨收付卻是 0（收盤結帳後可能查不到當日資料），現金這格先不動"
        proposal["blocked"] = True
        warnings.append("[現金] " + proposal["note"])
        return proposal

    if proposal["status"] == ledger.UNTRACKED:
        proposal["note"] = "還沒設定現金基準，要先校正一次才會開始自動累加"
        return proposal

    if proposal["status"] == ledger.MANUAL:
        proposal["note"] = (
            f"手動改過（程式記得 {cash.get('last_written')}），"
            f"要交還給程式要先重新校正一次基準"
        )
        return proposal

    proposal["proposed"] = ledger.cash_after(cash, today, net)
    proposal["will_write"] = not same_number(balance, proposal["proposed"])
    proposal["record_net"] = True
    proposal["note"] = f"今日淨收付 {show(net)}（{rows} 筆成交）"

    for day in ledger.missing_dates(cash, today):
        warnings.append(
            f"[現金] {day:%Y/%m/%d} 沒有淨收付紀錄。那天如果有成交，餘額會少算，"
            f"請對照對帳單直接把正確餘額填進 Excel 的 B8，下次登入程式會以它為準"
            f"（國定假日可忽略）"
        )

    return proposal


# 歷程的「說明」欄。項目、動作、新舊值三欄已經講完「哪一格、做了什麼、變成多少」，
# 說明再覆述一次等於空白，所以這裡只補那三欄講不出來的一件事：數字是從 Excel 抄來的。
ADOPTED_NOTE = "以 Excel 上的數字為準"


def adopt(proposals, book, sheet_name, today, today_included, at):
    """
    把「未接管 / 手動」的格子交還給程式管理。回傳 (訊息清單, 歷程項目, 還缺什麼)。

    持股沒有副作用：網頁庫存就是唯一真相，交還就是下次以網頁值重抄一次。
    現金不一樣 —— 程式無法從一個數字看出它含不含今天的淨收付，猜錯就是多扣或
    少扣一次，所以一定要 today_included 明講，沒講就不動現金那一格。

    交接本身也要進歷程。它是唯一一種「人決定、程式照做」的異動，
    日後回頭查帳最需要看到的就是這種：餘額為什麼從這天開始不一樣了。
    """
    messages, events = [], []
    needs_today = False

    for item in proposals:
        message, event, missing = adopt_one(item, book, sheet_name, today, today_included, at)
        needs_today = needs_today or missing
        if message:
            messages.append(message)
        if event:
            events.append(event)

    return messages, events, needs_today


def adopt_one(item, book, sheet_name, today, today_included, at):
    """
    交還單獨一格，回傳 (訊息, 歷程項目, 是否還缺 today_included)。

    介面上的「交還給程式」按鈕是一次處理一格，命令列的 --adopt 是全部跑一遍，
    兩邊走的是同一段程式碼，才不會出現「介面接管的結果跟命令列不一樣」。
    """
    if item["status"] not in (ledger.MANUAL, ledger.UNTRACKED):
        return None, None, False

    if item["kind"] == "cash":
        if item["current"] is None:
            return "B8 現金餘額是空的，沒有東西可以當基準，請先在 Excel 填一個數字", None, False
        if today_included is None:
            return None, None, True

        cash = book["cash"]
        was = cash.get("last_written")
        ledger.calibrate(cash, item["current"], today, item["net"], today_included, at)
        message = (
            f"現金基準改為 {show(cash['baseline_value'])}"
            f"（Excel 上的 {show(item['current'])} "
            f"{'已含' if today_included else '尚未含'}今日淨收付 {show(item['net'])}），"
            f"從 {today:%Y/%m/%d} 起重新起算"
        )
        return message, _event(at, sheet_name, item, "adopt", was, item["current"], message), False

    state = book["holdings"].setdefault(item["key"], {}).setdefault(item["which"], ledger.new_field())
    was = state.get("last_written")
    state["mode"] = ledger.AUTO
    state["last_written"] = item["current"]
    state["last_written_at"] = at
    state.pop("since", None)
    message = f"{item['cell']} {item['label']} 交還給程式管理"
    return message, _event(at, sheet_name, item, "adopt", was, item["current"], ADOPTED_NOTE), False


def initialize(sheet_data, book, sheet_name, today, at):
    """
    以 Excel 現在的數字為準，把一個分頁整個交給程式管理。介面在登入完成後跑一次。

    這裡刻意不需要網頁資料。初始化的定義就是「Excel 上這些數字是今天買賣之前的
    狀態」，今天在網頁上成交了什麼，等按「讀取網頁資料」時再往上加 —— 所以現金
    基準直接取 B8，今天的流水先記 0，之後 commit 會用真正的淨收付覆蓋掉那個 0。
    這也是為什麼不必再問「B8 含不含今天的淨收付」：登入的當下它一定還沒含。

    現金一天只初始化一次（基準日已經是今天就跳過）。今天的淨收付要是已經寫進
    B8 了，再重設一次基準會讓它被加第二次，而且畫面上不會有任何徵兆。

    跳過之後那一格就沒有出口了 —— 當天想改餘額只能由人明講，
    走 apply_cash_reset（介面上是讀完網頁資料、寫入之前跳的「重設現金餘額」對話框）。
    """
    events = []

    for line in sheet_data["rows"]:
        for which, col, current in (("qty", COL_QTY, line["qty"]),
                                    ("cost", COL_COST, line["cost"])):
            state = (book["holdings"].get(line["code"]) or {}).get(which)
            item = _row(line["row"], col, "holding", line["code"], which,
                        f"{'股數' if which == 'qty' else '成本'}（{line['label']}）",
                        current, None, None, ledger.status_of(state, current), "")
            _message, event, _needs = adopt_one(item, book, sheet_name, today, False, at)
            if event:
                events.append(event)

    balance = sheet_data["balance"]
    cash = book["cash"]
    if (balance is not None
            and ledger.status_of(cash, balance) != ledger.AUTO
            and cash.get("baseline_date") != today.isoformat()):
        row, col = CELL_BALANCE
        item = _row(row, col, "cash", "cash", "balance", "現金餘額",
                    balance, None, None, ledger.UNTRACKED, "")
        was = cash.get("last_written")
        ledger.calibrate(cash, balance, today, 0.0, False, at)
        events.append(_event(at, sheet_name, item, "adopt", was, balance, ADOPTED_NOTE))

    return events


def commit(proposals, book, sheet_name, today, at):
    """
    寫入成功之後更新紀錄檔，回傳要追加到歷程的項目。

    偵測到的人工改動也在這裡落地：狀態轉成手動、歷程記一筆。
    時間只能記「偵測到的時間」，因為程式沒在旁邊看著，實際什麼時候改的無從得知。
    """
    events = []

    for item in proposals:
        state = _state_of(book, item)

        if item["status"] == ledger.MANUAL and state.get("mode") != ledger.MANUAL:
            ledger.mark_manual(state, at)
            events.append(_event(at, sheet_name, item, "human",
                                 state.get("last_written"), item["current"],
                                 "執行時偵測到人工改動，已轉為手動"))

        if not item["will_write"]:
            continue

        events.append(_event(at, sheet_name, item, "program",
                             item["current"], item["proposed"], item["note"]))
        ledger.mark_written(state, item["proposed"], at)

    for item in proposals:
        if item["kind"] != "cash":
            continue
        if item["reset_to"] is not None:
            # calibrate 一次做完基準、流水、last_written 三件事，所以不必再 record_net。
            # reset_to 是「開盤前 + 今日淨收付」，所以 today_included 是 True，
            # calibrate 往回推一天份之後剛好回到人填的那個開盤前數字。
            was = book["cash"].get("last_written")
            ledger.calibrate(book["cash"], item["reset_to"], today, item["net"], True, at)
            if not item["will_write"]:
                # 算出來剛好等於 B8 上的數字，所以沒有格子要寫 —— 但基準確實被人
                # 改掉了。上面那一圈只替「寫過的格子」記歷程，這一筆不補就完全沒有
                # 痕跡：帳從此算得不一樣，卻沒有人知道是誰、什麼時候、改成什麼。
                # 記成 adopt 是因為變的是程式的記憶，Excel 一格都沒動。
                events.append(_event(at, sheet_name, item, "adopt", was,
                                     item["reset_to"], item["note"]))
        elif item.get("record_net"):
            ledger.record_net(book["cash"], today, item["net"])

    return events


def _state_of(book, item):
    if item["kind"] == "cash":
        return book["cash"]
    return book["holdings"].setdefault(item["key"], {}).setdefault(item["which"], ledger.new_field())


def _event(at, sheet_name, item, by, old, new, note):
    return {
        "at": at, "sheet": sheet_name, "cell": item["cell"], "label": item["label"],
        "by": by, "old": old, "new": new, "note": note,
    }
