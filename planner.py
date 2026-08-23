"""
把「網頁資料 × Excel 現值 × 紀錄檔」算成一張變更提案清單。

這裡不碰 Excel、不碰網路，只有純計算，所以命令列跟之後的介面可以共用
同一份判斷邏輯，也可以拿假資料直接測。

提案是一格一列（E5 跟 F5 分開），股數與成本分開算。畫面上要把兩格併成
一列顯示是顯示層的事。

Excel 上的現值一律被網頁值覆蓋，程式不記、不比對「這一格上次寫了什麼」——
修改 Excel 的風險交給操作的人自己管控。
"""

import ledger
from excel_io import COL_COST, COL_QTY, CELL_BALANCE
from util import cash_formula, cell_name, same_number, show, to_num

# 現金餘額的兩種算法（完整說明見 docs/現金餘額兩種算法.md）。
#
#   opening  今日初始現金餘額 + 今日淨收付
#   bank     銀行餘額 + 還沒交割的淨收付
#
# 兩種並存不是備援，是各有各正確的日子：全額交割股當天，錢在成交當下就從銀行餘額
# 扣走，但同一筆還掛在當日淨收付上（下午才更新改正），那天 bank 會扣兩次、
# opening 才是對的。反過來，匯撥與股利這種非交易現金流不會進淨收付，
# opening 看不到、bank 自動吸收。
METHOD_OPENING, METHOD_BANK = "opening", "bank"

METHOD_NAMES = {
    METHOD_OPENING: "初始餘額累加",
    METHOD_BANK: "銀行餘額推算",
}


def merge_holdings(arrays):
    """
    把未實現損益整理成 {股號: {股數, 成本, 名稱}}。

    網頁的資料是以「股票+交易別」為單位，同一檔理論上可能佔好幾列，所以股數
    直接相加、成本用股數加權平均合併——直接取其中一列的均價會是錯的。這條
    合併邏輯留著是防資料真的不對時算出錯的持股，不是為了配一句提醒（見下）。
    另外照網頁自己的規則濾掉當沖與成本股數 0 的列，否則加總會跟畫面對不起來。

    2026/08/22 使用者確認：同一檔同一天分批用現股買賣，網頁上券商已經先
    平均好、只會有一列；會佔多列的其他交易別（融資等）在目前的交易方式下
    不會發生。原本這裡「同一檔佔多列」時會多印一句「已用股數加權平均合併
    成本」的提醒，因為這個情境不會真的發生，那句提醒已經拿掉。
    """
    merged = {}

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

        record = merged.setdefault(stkno, {"qty": 0.0, "amount": 0.0, "name": ""})
        record["qty"] += qty
        record["amount"] += qty * to_num(item.get("priceavgn"))
        record["name"] = record["name"] or (item.get("stkna") or "").strip()

    for record in merged.values():
        record["cost"] = round(record["amount"] / record["qty"], 4) if record["qty"] else 0.0

    return merged


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


def bank_balance(rows):
    """
    銀行餘額（元）。讀不到、或讀不懂就回 None —— 絕對不要回 0。

    Amount 的單位是分（真的回應長這樣：0000000089300 = 畫面上的 893.00）。
    這是整條路上最容易錯又最不會被發現的一步：弄錯不報錯，只生出一個 100 倍大的數字。

    「讀不懂就回 None」是刻意的：to_num 預設把看不懂的東西當成 0，而 0 在這裡看起來
    像一個答案（這個人沒錢），會被拿去覆蓋 B8。真的遇過一次 —— 模擬頁面把負數補零成
    `000-793672200`，整個人的餘額就變成 0 了。
    """
    if not rows:
        return None
    amount = to_num(rows[0].get("Amount"), None)
    if amount is None:
        return None
    return round(amount / 100, 2)


def pending_rows(rows, today):
    """
    還沒交割的錢，一個成交日一項。回傳 (項目清單, 合計, 今天那一列的金額)。

    來源是交割金額查詢（query610），一天一列，金額已經是那一天的淨額：

        {"trade": "20260821", "cdate": "20260825", "pay_amt": "-238"}

    銀行餘額只含已經交割的錢，所以要補的正好是「cdate 比今天晚」的那幾列。
    交割日等於今天的那一列不補 —— 那筆錢今天早上就扣掉了（見 docs 裡的「前提」）。
    日期是 YYYYMMDD，直接比字串就是比大小。

    項目清單每一項是 (標籤, 金額)，照成交日由新到舊排：最新的那一個成交日是
    `T+0`（正常情況就是今天），往回一個交易日 `T+1`，依此類推。標籤是照回應裡
    的列數往回數的，不是照日曆 —— 中間隔著週末或連假時，日曆算不出「前一個
    交易日」，而這支查詢本來就一個交易日給一列。金額是 0 的那幾天照樣列出來，
    算式上少一項比多一個 0 難懂（畫面與 B8 的公式都直接照這個清單攤開）。

    今天那一列（trade 等於今天）另外回傳給呼叫方做交叉檢查用：未實現損益說
    今天有成交、這裡卻是 0 或根本沒有今天，就是這份資料有問題（見 _cash_blocked）。
    沒有今天那一列時回 None —— 「今天沒有這一列」跟「今天是 0」要分得出來。
    """
    stamp = today.strftime("%Y%m%d")
    items, total, today_amount = [], 0.0, None

    ordered = sorted(rows, key=lambda row: str(row.get("trade") or "").strip(),
                     reverse=True)
    for age, row in enumerate(ordered):
        amount = to_num(row.get("pay_amt"))
        if str(row.get("trade") or "").strip() == stamp:
            today_amount = amount
        if str(row.get("cdate") or "").strip() > stamp:
            items.append((f"T+{age}", amount))
            total += amount

    return items, round(total, 2), today_amount


def traded_today(pnl_arrays, today):
    """未實現損益的持股明細裡有沒有「今天」成交的紀錄。用來交叉檢查淨收付是不是漏了。"""
    stamp = today.strftime("%Y%m%d")
    for item in pnl_arrays:
        for detail in item.get("stkdat") or []:
            if str(detail.get("tdate") or "").strip() == stamp:
                return True
    return False


def plan(sheet_data, record, book, today, method=METHOD_OPENING):
    """
    算出這個分頁的提案。回傳 (提案清單, 提醒清單)。

    method 決定現金那一格用哪一種算法（見 METHOD_OPENING / METHOD_BANK）。
    只算選中的那一種：另一種算出來多少，對「今天該用哪一種」這個決定幫不上忙
    （理由見 _cash），而且它要的資料 fetch 那邊根本不會去抓。

    純讀取，不會動到 book —— 要落實到紀錄檔是 commit() 的事，
    這樣試算模式才能保證真的什麼都沒改到。
    """
    proposals = []
    warnings = []

    holdings = merge_holdings(record.get("未實現損益", []))

    seen = set()
    for line in sheet_data["rows"]:
        code = line["code"]
        seen.add(code)
        found = holdings.get(code)

        if found is None:
            # 網頁已經沒有這檔了。刻意不自動歸零 —— 使用者的流程是「先刪 Excel 再賣」，
            # 所以這通常代表忘了刪，該由人確認而不是程式清掉。
            # 2026/08/22 使用者要求縮短：訊息框現在混進逐筆歷程的時間序裡
            # （見 ui_sync._fill_notes），不必再自帶「股數與成本維持原樣，請確認
            # 是否忘記刪除這一列」這種說明——跟其他行同樣簡短。
            warnings.append(f"第 {line['row']} 列「{line['label']}」在網頁庫存中已不存在")
            for which, col, current in (("qty", COL_QTY, line["qty"]),
                                        ("cost", COL_COST, line["cost"])):
                proposals.append(_row(line["row"], col, "holding", code, which,
                                      f"{'股數' if which == 'qty' else '成本'}（{line['label']}）",
                                      current, None, None,
                                      note="網頁庫存已無此檔", missing=True))
            continue

        label = f"{code} {found['name']}"
        for which, col, current, web in (
            ("qty", COL_QTY, line["qty"], int(found["qty"])),
            ("cost", COL_COST, line["cost"], found["cost"]),
        ):
            proposals.append(_row(
                line["row"], col, "holding", code, which,
                f"{'股數' if which == 'qty' else '成本'}（{label}）",
                current, web, web,
            ))

    for code, found in holdings.items():
        if code not in seen:
            # 2026/08/22 使用者要求縮短，理由同上一句「網頁庫存已無此檔」。
            warnings.append(f"網頁有 {code} {found['name']}（{int(found['qty'])} 股）但 Excel 沒有這一列")

    proposals.append(_cash(sheet_data, record, book, today, warnings, method))
    return proposals, warnings


def _row(row, col, kind, key, which, label, current, web, proposed, note="", missing=False):
    """組一列提案。will_write 只有算出新值而且跟現值不同時才會是 True。"""
    will_write = (
        not missing
        and proposed is not None
        and not same_number(current, proposed)
    )
    return {
        "row": row, "col": col, "cell": cell_name(row, col),
        "kind": kind, "key": key, "which": which, "label": label,
        "current": current, "web": web, "proposed": proposed,
        "note": note, "missing": missing, "will_write": will_write,
        "reset_to": None, "formula": None,
    }


def apply_cash_reset(item, opening):
    """
    把人填的「今天開盤前的現金」套進現金那一列。就地改，回傳同一個 item。

    只能在 plan() 之後叫得動 —— 它要用今日淨收付，而那是網頁資料，
    登入的當下還沒去查（login_only 只登入、不抓資料）。這也正是介面上那顆
    「修改」要等讀完網頁資料才亮的原因：算得出結果，就不必請人回答「你填的數字
    含不含今天的成交」，那個問題每次都要人回想今天做過什麼，答錯的代價剛好是
    一整天的淨收付。
    """
    target = round(opening + item["net"], 2)
    item["proposed"] = target
    item["formula"] = cash_formula(opening, item["net"])
    item["reset_to"] = target
    item["record_net"] = False          # calibrate 會把流水一起寫好
    item["will_write"] = not same_number(item["current"], target)
    item["note"] = (f"今日初始現金餘額 {show(opening)} + 今日淨收付 {show(item['net'])}"
                    f"（{item['net_rows']} 筆成交），從今天起重新起算")
    return item


def _cash(sheet_data, record, book, today, warnings, method=METHOD_OPENING):
    """
    現金那一列。跟持股不同的地方全部集中在這裡。

        opening  今日初始現金餘額 + 今日淨收付   累加，要有基準
        bank     銀行餘額 + 還沒交割的淨收付     快照，網頁直接給絕對值

    只算選中的那一種。曾經兩種都算、把另一種也顯示在畫面上當對帳訊號，後來拿掉了
    —— 差額講不出是哪一種原因造成的（全額交割？匯撥？有人手改過 B8？），
    對「今天要用哪一種」這個決定幫不上忙，而那個決定的依據（今天有沒有買全額交割股）
    本來就只有人知道。省下來的還有兩支查詢：用 opening 的日子完全不必碰銀行餘額。
    """
    row, col = CELL_BALANCE
    balance = sheet_data["balance"]
    cash = book["cash"]

    # 當日淨收付兩種算法都要抓：opening 拿它算餘額；bank 用不到它算餘額，但
    # 「重設基準」要它，而且它是這一輪唯一一支「每一列都帶 bhno/cseq」的現金相關
    # 資料 —— 銀行餘額與交割金額都沒有回顯帳號，身分核對全靠它擋在前面
    # （見 fetch.collect 與下面 _cash_blocked 的第一道）。
    net, rows = settlement_total(record.get("當日淨收付", []))

    proposal = _row(row, col, "cash", "cash", "balance", "現金餘額", balance, net, None)
    proposal["net"] = net
    proposal["net_rows"] = rows
    proposal["method"] = method
    proposal["record_net"] = False
    # 淨收付本身信不過的時候立起來。這一格連「重設基準」都不該讓人按 ——
    # 拿一個已知是錯的 net 去 calibrate，等於把今天的成交永久算進基準裡。
    proposal["blocked"] = False

    bank = pending = today_amount = None
    pending_items = []
    if method == METHOD_BANK:
        bank = bank_balance(record.get("銀行餘額"))
        pending_items, pending, today_amount = pending_rows(
            record.get("交割金額", []), today)
    proposal["bank"] = bank
    proposal["pending"] = pending
    proposal["pending_items"] = pending_items

    if balance is None:
        proposal["note"] = "B8 是空的或不是數字，請先填一個現金餘額"
        return proposal

    reason = _cash_blocked(record, today, method, net, bank, today_amount)
    if reason is not None:
        proposal["note"] = reason
        proposal["blocked"] = True
        # 空字串 = 擋，但不多講一句（見 _cash_blocked 裡銀行餘額那一道）。
        if reason:
            warnings.append("[現金] " + reason)
        return proposal

    if method == METHOD_OPENING and cash.get("baseline_value") is None:
        proposal["note"] = "還沒有今日初始現金餘額，這組讀到網頁資料就會自動設定"
        return proposal

    if method == METHOD_BANK:
        amounts = [amount for _label, amount in pending_items]
        proposal["proposed"] = round(bank + pending, 2)
        proposal["formula"] = cash_formula(bank, *amounts)
        proposal["note"] = _bank_note(bank, pending_items)
    else:
        proposal["proposed"] = ledger.cash_after(cash, today, net)
        proposal["formula"] = cash_formula(cash.get("baseline_value"), net)
        # 不接「（N 筆成交）」（2026/08/21 使用者要求）。這一句要回答的是
        # 「這個餘額怎麼來的」，筆數不在算式上，接在後面只是把一句已經很長的
        # 話再拉長。筆數還在 net_rows 裡，「修改今日初始現金餘額」那個視窗有列。
        proposal["note"] = _formula_note(
            "今日初始現金餘額", cash.get("baseline_value"),
            [("今日淨收付", net)], proposal["proposed"])

    # 現金這一格只要算得出來就寫，兩種算法都一樣（2026/08/21 使用者要求）。
    #
    # 股數/成本那幾格還是「數字跟現在不一樣才寫」，現金不跟這條規則，
    # 因為它寫進去的是公式（=893-238+0），不是一個數字。拿 same_number 擋，
    # 擋掉的是「今天算出來剛好也是 655」—— 但格子裡那條公式可能是昨天的
    # =900-245，數字對、明細過期，看起來卻像今天剛算的。人看 B8 不只看那個數字，
    # 還看它是怎麼來的，所以每讀一次就把今天的算式覆蓋上去。
    #
    # 前面幾道「先不動」（讀不到資料、對不起來、還沒有基準）都已經在上面 return 掉了，
    # 能走到這裡代表這個數字是算得出來、而且信得過的。
    proposal["will_write"] = proposal["proposed"] is not None
    proposal["record_net"] = True
    return proposal


def _bank_note(bank, items):
    """
    銀行餘額推算那一句：算式在前、數字在後，一個交割日一項。

        銀行餘額 + 淨收付(T+0) + 淨收付(T+1) = 893 - 238 + 0 = 655

    左邊是「這個數字是怎麼來的」，中間是「這次各是多少」，最後一個等號是算出來的
    餘額 —— 就是要寫進 B8 的那個數字，不必自己心算才知道這串加起來是多少。
    攤成一天一項是使用者 2026/08/21 要求的 —— 原本寫成「銀行餘額 893 + 還沒交割的
    -238（1 天的成交還沒扣款）」，合計看不出是哪幾天湊的，對不上時只能回網頁
    一列一列數。
    """
    if not items:
        return _formula_note("銀行餘額", bank, [], None,
                             "（沒有還沒交割的成交）")
    labelled = [(f"淨收付({label})", amount) for label, amount in items]
    total = round(bank + sum(amount for _label, amount in items), 2)
    return _formula_note("銀行餘額", bank, labelled, total)


def _formula_note(base_label, base, items, total, tail=""):
    """
    「這個數字是怎麼來的」那一句：名字在前、數字在中、算出來的餘額在最後。

        銀行餘額 + 淨收付(T+0) + 淨收付(T+1) = 893 - 238 + 0 = 655
        今日初始現金餘額 + 今日淨收付 = 893 - 238 = 655

    兩種算法共用同一個形狀（2026/08/21 使用者要求）。今日初始那一種原本只寫
    「今日淨收付 -238（1 筆成交）」—— 只講得出「今天動了多少」，講不出
    「從多少開始、算完是多少」，而那两個數字才是要拿去跟 B8 對的。

    中間那段數字跟寫進 B8 的公式是同一串（見 util.cash_formula），只差在這裡
    加了空白與千分位逗點。畫面跟 Excel 對不起來的時候，兩邊講的是同一句話。

    total 由呼叫方給，不在這裡加 —— 要顯示的正是真正要寫進 B8 的那個數字，
    自己再加一次等於另外算一遍，算法一改就會跟實際寫進去的值分家。
    """
    if not items:
        return f"{base_label} = {show(base)}{tail}"
    names = " + ".join(label for label, _amount in items)
    numbers = show(base) + "".join(
        f" {'-' if amount < 0 else '+'} {show(abs(amount))}"
        for _label, amount in items)
    return f"{base_label} + {names} = {numbers} = {show(total)}{tail}"


def _cash_blocked(record, today, method, net, bank, today_amount):
    """
    現金這一格該不該整個擋住不寫。要擋就回一句給人看的話，不擋回 None。
    回空字串也是擋，只是不多講那一句 —— 用在「別的地方已經講過了」的情況。

    兩種算法各有各的破口，但形狀一樣：某一項資料其實是「查不到」，
    而查不到跟「就是 0」在數字上長得一模一樣。分不出來的時候一律不寫 ——
    少寫一次人看得到，寫錯一次沒有人會發現。
    """
    if "當日淨收付" not in record:
        return "今日淨收付沒有讀到，現金這格先不動"

    if method == METHOD_BANK:
        if bank is None:
            # 擋住但不說話（2026/08/21 使用者要求）。這一道只剩兩種觸發方式，
            # 兩種都不需要這句話：查詢真的失敗時，fetch 已經把具體原因寫進
            # record["problems"]（見 fetch.bank_problem），畫面上那句講得比這句清楚；
            # 剩下的情況是讀完才用點名字換算法（測試用入口，.env 關得掉），
            # 那是自己剛手動改的，不必再被提醒一次。
            return ""
        if "交割金額" not in record:
            return "交割金額查詢沒有讀到，算不出還有哪幾天沒交割，現金這格先不動"
        # 交叉檢查：未實現損益說今天有成交，交割金額查詢卻說今天沒有錢要交割。
        # 今天成交的要 T+2 才扣，所以那筆錢一定還掛在這支查詢上；掛不上就是這份
        # 資料不對（收盤結帳後這支還準不準，沒有人在那個時段查過）。少算的正好是
        # 今天成交的那一整筆，而畫面上不會有任何徵兆，所以整格不寫。
        if traded_today(record.get("未實現損益", []), today) and not today_amount:
            return ("未實現損益顯示今天有成交，交割金額查詢裡今天卻沒有金額"
                    "（收盤結帳後可能查不到），現金這格先不動")
        return None

    # 交叉檢查：未實現損益說今天有成交，淨收付卻是 0，兩邊矛盾。
    # 最可能的原因是收盤結帳後查不到當日資料了，這時候記 0 會讓餘額默默漏掉今天。
    # 只擋現金這一格，持股照樣可以寫。
    if net == 0 and traded_today(record.get("未實現損益", []), today):
        return ("未實現損益顯示今天有成交、淨收付卻是 0"
                "（收盤結帳後可能查不到當日資料），現金這格先不動")
    return None


# 現金的初始化每天都會發生一次，值常常跟昨天一樣（昨天收盤多少，今天就從多少開始），
# 所以歷程上會出現一排「893 → 893」。說明寫清楚它是什麼，那一列才讀得懂。
OPENING_NOTE = "今日初始現金餘額（今天第一次登入時的 B8）"


def initialize(sheet_data, book, sheet_name, today, at):
    """
    現金每天重新起算：只要基準日還不是今天，就把 B8 收成今天的起點。
    介面在登入成功後跑一次。股數/成本不需要初始化 —— 程式每次都直接以
    網頁值覆蓋，沒有「先接管」這一步。

    這裡刻意不需要網頁資料。今天在網頁上成交了什麼，等按「讀取」時
    再往上加 —— 所以現金基準直接取 B8，今天的流水先記 0，之後 commit 會用
    真正的淨收付覆蓋掉那個 0。這也是為什麼不必再問「B8 含不含今天的淨收付」：
    登入的當下它一定還沒含。

    一天也只設一次（基準日已經是今天就跳過）。今天的淨收付要是已經寫進
    B8 了，再重設一次基準會讓它被加第二次，而且畫面上不會有任何徵兆。

    跳過之後那一格不是就沒有出口了 —— 當天想改餘額由人明講，走 apply_cash_reset
    （介面上現金那一條底下的「今日初始現金餘額　[修改]」）。
    """
    balance = sheet_data["balance"]
    cash = book["cash"]
    if balance is None or cash.get("baseline_date") == today.isoformat():
        return []

    row, col = CELL_BALANCE
    item = _row(row, col, "cash", "cash", "balance", "現金餘額", balance, None, None)
    was = cash.get("last_written")
    ledger.calibrate(cash, balance, today, 0.0, False, at)
    return [_event(at, sheet_name, item, "adopt", was, balance, OPENING_NOTE)]


def commit(proposals, book, sheet_name, today, at):
    """寫入成功之後更新紀錄檔，回傳要追加到歷程的項目。"""
    events = []

    for item in proposals:
        if not item["will_write"]:
            continue
        events.append(_event(at, sheet_name, item, "program",
                             item["current"], item["proposed"], item["note"]))

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


def _event(at, sheet_name, item, by, old, new, note):
    return {
        "at": at, "sheet": sheet_name, "cell": item["cell"], "label": item["label"],
        "by": by, "old": old, "new": new, "note": note,
    }
