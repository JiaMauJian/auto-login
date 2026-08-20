"""
唯讀偵察腳本：登入後把「未實現損益」「當日淨收付」「昨日淨收付」「銀行餘額」的原始回應
抓下來存檔，什麼都不寫。

這支程式不會碰 Excel、不會下單、不會改網站上的任何東西，只做兩件事：
把原始 JSON 存成檔案，以及印出一份摘要。目的是在動手寫 update_excel.py 之前，
先用真實資料回答幾個「光看網頁原始碼看不出來」的問題：

1. 同一檔股票會不會出現多列？（未實現損益的資料是以「股票+交易別」為單位，
   現股與融資會各自成列；零股據使用者說會跟現股合併，但要實測確認。）
   Excel 一檔股票只有一列，多列就得決定怎麼合併，成交均價還得用股數加權。
2. 回應裡有沒有回顯帳號？多帳號時所有分頁共用同一個 JSESSIONID，
   後登入的會把前一個的 server session 頂掉。如果回應裡帶了帳號，
   就能在寫 Excel 前核對「這份資料真的是這個帳號的」；如果沒有，
   就只能靠「登入完立刻抓」的時序來保證，得知道自己在賭什麼。
3. 收盤結帳（每日約 17~20 點）之後查「當日」還拿不拿得到資料？
   拿不到的話，那個時段跑更新會靜靜地把淨收付當成 0，餘額不會扣。
4. （2026/08/19 新增，為了現金餘額的第二種算法，見 docs/現金餘額兩種算法.md）
   queryBankBalance 的參數怎麼帶、回應裡哪個欄位是銀行餘額？
   queryType=6（前一日）回的是不是真的前一個交易日？（用 transDateQuery 回的
   lastTransDate 交叉比對，不必等到週一才驗得出來。）

因為第 2 點，這裡刻意採用「每登入完一組帳號就立刻抓完它的資料」的順序，
而不是全部登入完再回頭抓 —— 後者前面幾個帳號的 session 可能早就被頂掉了。

抓資料的方式是在已登入的分頁裡呼叫網站自己的 AJAX（POST /tbb/MainController），
沿用頁面現成的 B64_XOR_Encode / XOR_KEY 與 cookie，不做畫面解析。

輸出在 偵察資料\ 資料夾（已加進 .gitignore）。裡面是你的帳務資料，不要隨手外流。
"""

import json
import sys
import traceback
from datetime import datetime, timedelta

from playwright.sync_api import Error as PlaywrightError, TimeoutError as PlaywrightTimeoutError
from playwright.sync_api import sync_playwright

from login import (
    app_dir,
    configure_browsers_path,
    do_login,
    load_accounts,
    open_context,
    pause,
    wait_until_finished,
)
from util import to_num

# 抓資料前先把分頁導到這裡。這頁一定載入了 common.js（B64_XOR_Encode / XOR_KEY 在裡面），
# 而且它本來就是「未實現損益」，你可以直接用眼睛跟腳本印出來的數字對照。
ACCOUNT_PAGE = "https://www.tbbstock.com.tw/tbb/account/layoutRWD.jsp?type=4"

# 銀行餘額查詢。這頁只有四個欄位，值由頁面自己打一次 queryBankBalance 填進去，
# 所以它同時是「AJAX 的答案對不對」的對照組 —— 畫面上那個數字是網站自己算的。
BANK_PAGE = "https://www.tbbstock.com.tw/tbb/account/layoutRWD.jsp?type=12"

OUTPUT_DIR_NAME = "偵察資料"

# 交易別代碼對照，跟網頁上的 OrderMeta.trade 一致。
TRADE_NAMES = {
    "0": "現股", "1": "代辦融資", "2": "代辦融券", "3": "融資", "4": "融券",
    "5": "借券", "6": "金融商品借券", "7": "當沖融資", "8": "當沖融券",
    "9": "自動當沖", "A": "現沖賣",
}

# 在頁面裡呼叫網站自己的查詢 API。用 fetch 而不是 $.ajax，是因為頁面原本那支帶了
# async:false（同步阻塞），在自動化裡等於把整個分頁鎖住。
FETCH_JS = """
async ({ cmd, paramInfo }) => {
    if (typeof B64_XOR_Encode !== 'function' || typeof XOR_KEY === 'undefined') {
        return { error: '這個頁面找不到 B64_XOR_Encode / XOR_KEY，common.js 可能沒載入' };
    }
    const body = new URLSearchParams();
    body.set('CMD', cmd);
    body.set('Param', B64_XOR_Encode(JSON.stringify(paramInfo), XOR_KEY));
    const resp = await fetch('/tbb/MainController?timestamp=' + Date.now(), {
        method: 'POST',
        headers: { 'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8' },
        body: body.toString(),
        credentials: 'same-origin',
    });
    return { status: resp.status, text: await resp.text() };
}
"""

# 只讀身分相關的欄位。sessionStorage 裡的 login_info 含有密碼與身分證字號，
# 絕對不要 dump 出來寫進檔案，這裡只記錄有哪些 key 存在。
SESSION_JS = """
() => ({
    branch_id: sessionStorage.getItem('branch_id'),
    cust_id: sessionStorage.getItem('cust_id'),
    account: sessionStorage.getItem('account'),
    systemid: sessionStorage.getItem('systemid'),
    ratest: sessionStorage.getItem('ratest'),
    keys: Object.keys(sessionStorage),
})
"""


# 銀行餘額頁的四個欄位。值在 <span id="x"><input value="..." disabled></span> 裡面，
# 所以要讀 input 的 value，textContent 是空的。
BANK_JS = """
() => {
    const val = (id) => {
        const box = document.getElementById(id);
        if (!box) return '';
        const input = box.querySelector('input');
        return ((input ? input.value : box.textContent) || '').trim();
    };
    return {
        date: val('tDate'), time: val('tTimes'),
        account: val('tAccount'), balance: val('tBalance'),
    };
}
"""


def account_codes(items):
    """
    把資料裡回顯的帳號挑出來，回傳集合（格式 1112-0108640）。

    2026/08/17 偵察發現：回應的每一列都帶 bhno（分公司）與 cseq（客戶號），
    連 stkdat / matdat 明細層也有。這是多帳號時能「硬核對這份資料屬於誰」的關鍵 ——
    共用 JSESSIONID 導致 session 被頂掉時，不必只靠「登入完立刻抓」的時序來賭。
    """
    found = set()

    def walk(node):
        if isinstance(node, dict):
            bhno, cseq = node.get("bhno"), node.get("cseq")
            if bhno and cseq:
                found.add(f"{bhno}-{cseq}")
            for value in node.values():
                walk(value)
        elif isinstance(node, list):
            for value in node:
                walk(value)

    walk(items)
    return found


def query(page, cmd, param_info):
    """呼叫一次 MainController，回傳 (解析後的 dict, 原始文字)。失敗時 dict 為 None。"""
    result = page.evaluate(FETCH_JS, {"cmd": cmd, "paramInfo": param_info})

    if "error" in result:
        return None, result["error"]

    raw = result.get("text", "")
    if result.get("status") != 200:
        return None, f"HTTP {result.get('status')}\n{raw}"

    try:
        return json.loads(raw), raw
    except json.JSONDecodeError:
        # 回應不是 JSON，最常見的原因是 session 過期被導去登入頁。
        return None, raw


def summarize_pnl(data, expect_acc=None):
    """未實現損益：把重點欄位列出來，並標記出「同一檔股票佔了多列」的情況。"""
    lines = []
    arrays = data.get("arrays") or []
    lines.append(f"retcode={data.get('retcode')} retmsg={data.get('retmsg')} 筆數={len(arrays)}")
    lines.append(f"回應最上層欄位: {sorted(data.keys())}")
    lines.extend(check_account(data, expect_acc))

    if arrays and not (arrays[0].get("stkno") or arrays[0].get("stkna")):
        lines.append("（第一筆的 stkno 與 stkna 都是空的 = 網頁定義的「目前沒有資料」）")
        return lines

    seen = {}
    for i, item in enumerate(arrays):
        stkno = item.get("stkno", "")
        trade = item.get("trade", "")
        seen.setdefault(stkno, []).append(TRADE_NAMES.get(trade, trade))
        # 網頁自己會濾掉當沖與成本股數 0 的列，這裡標記出來但不隱藏。
        skipped = ""
        if trade in ("9", "A"):
            skipped = "  <- 網頁會濾掉(當沖)"
        elif str(item.get("costqtyn", "")).strip() == "0":
            skipped = "  <- 網頁會濾掉(成本股數0)"
        lines.append(
            f"  [{i}] {stkno} {item.get('stkna', '')} 交易別={TRADE_NAMES.get(trade, trade)}"
            f" 成交股數={item.get('costqtyn')} 成交均價={item.get('priceavgn')}"
            f" 投資成本={item.get('costsum')} 現價={item.get('pricemkt')}"
            f" 損益={item.get('makeasum')} 報酬率={item.get('makeaper')}"
            f" 明細筆數={len(item.get('stkdat') or [])}{skipped}"
        )

    dups = {no: kinds for no, kinds in seen.items() if len(kinds) > 1 and no}
    if dups:
        lines.append("!! 同一檔股票佔了多列，Excel 只有一列，必須決定怎麼合併：")
        for no, kinds in dups.items():
            lines.append(f"   {no} 出現 {len(kinds)} 次：{'、'.join(kinds)}")
    else:
        lines.append("同一檔股票都只有一列（本次資料沒有需要合併的情況）")

    lines.append(f"欄位一覽（第一筆）: {sorted(arrays[0].keys())}" if arrays else "無資料")
    return lines


def check_account(data, expect_acc):
    """核對回應裡回顯的帳號跟預期的是否一致。寫 Excel 前的最後一道防線。"""
    codes = account_codes(data)
    if not codes:
        return ["回顯帳號: 這次的資料裡找不到 bhno/cseq（沒有資料時屬正常）"]
    line = f"回顯帳號: {'、'.join(sorted(codes))}"
    if expect_acc and codes != {expect_acc}:
        return [line + f"  !! 與預期的 {expect_acc} 不符，這份資料可能不是這個帳號的"]
    return [line + ("  （與登入帳號相符）" if expect_acc else "")]


def date_fields(rows):
    """
    把每一列裡「長得像日期」的欄位挑出來，回傳 {欄位名: 出現過的值}。

    這是拿來回答「這份資料到底是哪一天的」用的 —— 金額對不上不代表錯（昨天沒成交
    的話當日與昨日都是 0，看起來一模一樣），資料自己帶的日期才是唯一的證據。

    不寫死欄位名稱是因為還不知道有哪些：已知 cdate 是交割日、未實現損益那邊有 tdate，
    但淨收付這邊有幾個日期欄位沒人看過。凡是 8 位數字的值都撈起來，多撈幾個沒有壞處。
    """
    found = {}
    for row in rows:
        for key, value in row.items():
            text = str(value or "").strip()
            if len(text) == 8 and text.isdigit() and text.startswith("20"):
                found.setdefault(key, set()).add(text)
    return found


def summarize_settlement(data, expect_acc=None):
    """當日淨收付：照網頁的算法把 payn 加總，好跟畫面上的「淨收付合計」對答案。"""
    lines = []
    arrays = data.get("arrays") or []
    lines.append(f"retcode={data.get('retcode')} retmsg={data.get('retmsg')} 筆數={len(arrays)}")
    lines.append(f"回應最上層欄位: {sorted(data.keys())}")
    lines.extend(check_account(data, expect_acc))

    buy = sell = 0.0
    rows = 0
    for item in arrays:
        for mat in item.get("matdat") or []:
            # 網頁會跳過成交股數 0 的列，加總要跟著跳，否則對不上畫面。
            if str(mat.get("qty", "")).strip() == "0":
                continue
            rows += 1
            payn = to_num(mat.get("payn"))
            if payn < 0:
                buy += payn
            else:
                sell += payn
            lines.append(
                f"  {item.get('stkno', '')} {mat.get('stkna', '')}"
                f" {TRADE_NAMES.get(mat.get('trade'), mat.get('trade'))}"
                f" {mat.get('bs')}{mat.get('reason') or ''}"
                f" 股數={mat.get('qty')} 單價={mat.get('price')} 價金={mat.get('priceqty')}"
                f" 手續費={mat.get('fee')} 交易稅={mat.get('tax')} 淨收付={mat.get('payn')}"
                f" 委託書號={mat.get('ordno')}"
            )

    lines.append(f"成交明細 {rows} 筆")
    lines.append(f"應付金額 = {buy:,.0f}")
    lines.append(f"應收金額 = {sell:,.0f}")
    lines.append(f"淨收付合計 = {buy + sell:,.0f}   <- 這個數字要跟網頁畫面一致")

    # 這份資料是哪一天的。「當日」與「昨日」兩支查詢的差別只在一個參數，
    # 這幾行就是判斷那個參數到底有沒有生效的唯一證據（金額看不出來）。
    mats = [mat for item in arrays for mat in (item.get("matdat") or [])]
    dates = date_fields(mats)
    if dates:
        lines.append("資料帶的日期欄位（用來確認這是哪一天的資料）：")
        for key, values in sorted(dates.items()):
            lines.append(f"   {key} = {'、'.join(sorted(values))}")
    elif rows:
        lines.append("!! 這幾列裡找不到任何日期欄位，無法確認資料屬於哪一天")
    return lines


# queryBankBalance 的 paramInfo 還沒看過，所以一次試幾種寫法，誰通了就知道是哪一種。
# 順序是「最像其他查詢的」排前面。全部失敗也不算白跑 —— 下面還會去讀畫面上的數字。
BANK_PARAM_GUESSES = [
    ("branchId+custId（跟其他查詢同款）", lambda bid, cid: {"branchId": "1" + bid, "custId": cid}),
    ("branchId 不加 1", lambda bid, cid: {"branchId": bid, "custId": cid}),
    ("多帶 his/queryType", lambda bid, cid: {
        "branchId": "1" + bid, "custId": cid, "his": "y", "queryType": "1",
    }),
]


# 往前抓幾個日曆天。T+2 最多只欠三個交易日，但中間可能夾週末與連假，
# 所以用日曆天多抓一點，寧可多幾列也不要漏。180 天是網站自己的上限，離得很遠。
LOOKBACK_DAYS = 10


def range_query(page, bid, cid, start, end):
    """指定一段日期區間查淨收付。start/end 是 datetime.date。回傳 (資料, 摘要文字)。"""
    lines = ["", f"--- 指定區間 {start:%Y/%m/%d} ~ {end:%Y/%m/%d}（方法二要用的那一支）---"]
    data, raw = query(page, "queryInstantAccount_new", {
        "branchId": "1" + bid, "custId": cid,
        "his": "y", "queryType": "0",
        "startDate": f"{start:%Y/%m/%d}", "endDate": f"{end:%Y/%m/%d}",
        "range": "stksum,stkdat,matsum,matdat", "stock_no": "",
    })
    if data is None:
        lines.append(f"查詢失敗：{str(raw)[:300]}")
        return None, lines
    lines.extend(summarize_settlement(data, None))
    return data, lines


def unsettled_rows(data, today):
    """
    交割日還沒到的成交明細。回傳 (合計, 每一列的說明, 沒有 cdate 的筆數)。

    這就是方法二的核心：銀行餘額只含已經交割的錢，所以要補的正好是「cdate 比今天晚」
    的那幾筆。用 cdate 而不是「今天+昨天」，是因為週一與連假後不只欠兩天
    （週四、週五兩天的成交都還沒交割），而 cdate 是每一筆自己帶著的，不必推算。

    cdate 讀不到的列要單獨數出來 —— 那種列沒有辦法判斷該不該加，
    不能默默當成「已交割」跳過（那等於少算錢）。
    """
    stamp = today.strftime("%Y%m%d")
    total, rows, unknown = 0.0, [], 0

    for item in data.get("arrays") or []:
        for mat in item.get("matdat") or []:
            if str(mat.get("qty") or "").strip() == "0":
                continue
            cdate = str(mat.get("cdate") or "").strip()
            payn = to_num(mat.get("payn"))
            if not cdate:
                unknown += 1
                continue
            mark = "要補" if cdate > stamp else "已交割，不補"
            rows.append(
                f"   {item.get('stkno', '')} {mat.get('stkna', '')}"
                f" 成交日={mat.get('tdate')} 交割日={cdate} 淨收付={payn:,.0f}  <- {mark}"
            )
            if cdate > stamp:
                total += payn

    return total, rows, unknown


def query_trans_date(page, bid, cid):
    """
    問伺服器「上一個交易日」是哪一天。回傳 (YYYYMMDD 或 None, 摘要文字)。

    這支是對帳單頁選「指定日期」時自己會打的（CMD=transDateQuery），回應裡的
    lastTransDate 就是它拿來當預設起始日的那一天。

    為什麼值得多打這一支：「前一日」到底是不是「前一個交易日」，本來要等到週一
    才驗得出來（平常日的前一日就是昨天，看不出差別）。有了 lastTransDate 就不必等 ——
    直接拿它跟 queryType:'6' 回來的資料日期對一下就知道。連假前後也是同一招。

    注意 branchId 這裡**不加 '1'**，跟其他查詢不一樣。這是照頁面自己的寫法抄的，
    不是筆誤 —— 改成一樣反而可能查不到。
    """
    lines = ["", "--- 上一個交易日（transDateQuery）---"]
    data, raw = query(page, "transDateQuery", {"branchId": bid, "custId": cid})

    if data is None:
        lines.append(f"查詢失敗：{str(raw)[:300]}")
        return None, lines

    lines.append(f"retcode={data.get('retcode')} retmsg={data.get('retmsg')}")
    lines.append("回應內容: " + json.dumps(data, ensure_ascii=False)[:800])

    last = str(data.get("lastTransDate") or "").strip()
    if not last:
        lines.append("!! 回應裡沒有 lastTransDate，這條交叉檢查的路走不通")
        return None, lines

    lines.append(f"上一個交易日 = {last}   <- 「昨日淨收付」該是這一天的資料")
    return last, lines


def recon_bank(page, bid, cid, acc_code):
    """
    銀行餘額查詢（type=12）。回傳 (摘要文字, 要存檔的原始資料)。

    做兩件事，缺一不可：

        AJAX    試 queryBankBalance 的參數寫法，找出能用的那一種
        讀畫面  開那一頁、等它自己查完，把畫面上的數字讀下來

    讀畫面不只是備援，它是**對照組**：畫面上那個數字是網站自己算出來的，
    AJAX 回應裡哪一個欄位才是「銀行餘額」，要靠它才認得出來（回應裡通常不只一個數字）。

    注意這個函式會把分頁導去 type=12，所以它一定要排在其他查詢之後。
    """
    lines = ["", "--- 銀行餘額（type=12, queryBankBalance）---"]
    dumps = {}

    for label, build in BANK_PARAM_GUESSES:
        param_info = build(bid, cid)
        data, raw = query(page, "queryBankBalance", param_info)
        lines.append(f"送出參數（{label}）: {json.dumps(param_info, ensure_ascii=False)}")
        if data is None:
            lines.append(f"  失敗：{str(raw)[:300]}")
            continue
        lines.append(f"  retcode={data.get('retcode')} retmsg={data.get('retmsg')}")
        lines.append(f"  回應最上層欄位: {sorted(data.keys())}")
        # 整份回應原樣印出來。這是唯一一次能看到欄位名稱的機會，而它只有幾個數字，
        # 不像庫存那樣長 —— 直接攤開比事後翻檔案快。
        lines.append("  回應內容: " + json.dumps(data, ensure_ascii=False)[:1500])
        lines.extend("  " + line for line in check_account(data, acc_code))
        dumps[f"銀行餘額_{label}"] = data
        if data.get("retcode") == "000000":
            lines.append("  ^ 這一組通了")
            break

    lines.append("")
    lines.append(f"改讀畫面上的數字（{BANK_PAGE}）：")
    try:
        page.goto(BANK_PAGE)
        # 頁面自己會打一次查詢再把值填進去，所以不能只等 load —— 要等那個欄位真的有值。
        page.wait_for_function(
            "() => { const i = document.querySelector('#tBalance input');"
            " return i && i.value.trim() !== ''; }",
            timeout=20000,
        )
        shown = page.evaluate(BANK_JS)
    except (PlaywrightError, PlaywrightTimeoutError) as exc:
        lines.append(f"  讀不到：{exc}")
        return lines, dumps

    lines.append(f"  查詢日期={shown['date']} 查詢時間={shown['time']}")
    lines.append(f"  銀行帳號={shown['account']}  銀行餘額={shown['balance']}")
    dumps["銀行餘額_畫面"] = shown

    # 銀行帳號裡看得到客戶號的話，多帳號時就能核對「這個餘額真的是這個人的」。
    # 銀行餘額是單一數字，抄錯人不會有任何徵兆，所以這道檢查值得為它多寫幾行。
    if cid and cid.lstrip("0") and cid.lstrip("0") in shown["account"]:
        lines.append(f"  銀行帳號裡含著客戶號 {cid} —— 可以拿來核對身分")
    else:
        lines.append(f"  !! 銀行帳號裡看不到客戶號 {cid}，這條核對身分的路走不通，要另想辦法")

    return lines, dumps


def method_two(dumps, span, today):
    """
    把方法二整條算完印出來：銀行餘額 + 還沒交割的成交淨收付。純顯示，什麼都不寫。

    偵察報告最後有這一段，是因為前面那些 retcode 與欄位名對不對，人是看不出來的；
    「算出來的餘額跟我 Excel 上那個數字一不一樣」才看得出來。
    """
    lines = ["", "--- 方法二試算（只印出來，不寫任何東西）---"]

    ajax = dumps.get("銀行餘額_branchId+custId（跟其他查詢同款）")
    shown = dumps.get("銀行餘額_畫面") or {}
    balance = None

    if isinstance(ajax, dict):
        rows = ajax.get("data") or []
        if len(rows) != 1:
            lines.append(f"!! 銀行帳戶有 {len(rows)} 筆，不知道該用哪一個，方法二在這裡要擋住")
        if rows:
            # Amount 的單位是分。這是整條路上最容易錯又最不會被發現的一步 ——
            # 弄錯只會生出一個 100 倍大、看起來很像真的數字，所以這裡跟畫面對一次答案。
            balance = to_num(rows[0].get("Amount")) / 100
            lines.append(f"AJAX 的 Amount = {rows[0].get('Amount')} -> {balance:,.2f}")
            lines.append(f"畫面上的銀行餘額 = {shown.get('balance')}")
            if to_num(shown.get("balance")) != balance:
                lines.append("!! 兩邊對不起來，Amount 的單位跟想的不一樣")

    if balance is None or span is None:
        lines.append("資料不齊，算不出來")
        return lines

    total, rows, unknown = unsettled_rows(span, today)
    lines.append(f"還沒交割的成交（交割日比 {today:%Y/%m/%d} 晚）：")
    lines.extend(rows or ["   （這段區間內沒有任何成交）"])
    if unknown:
        lines.append(f"!! 有 {unknown} 筆讀不到交割日，這種列不能默默跳過（會少算錢）")
    lines.append(f"合計要補的金額 = {total:,.0f}")
    lines.append(f"方法二算出來的現金餘額 = {balance:,.2f} + {total:,.0f} = {balance + total:,.2f}")
    lines.append("   ^ 這個數字要跟你 Excel 上的現金餘額對得起來")
    return lines


def recon_one(page):
    """在已登入的分頁上抓兩份資料。回傳 (帳號代碼, 摘要文字列表, 要存檔的原始資料)。"""
    lines = []
    dumps = {}

    page.goto(ACCOUNT_PAGE)
    # 這頁自己也會跑一次查詢，等它安靜下來再動作，避免跟它搶。
    # 等不到也無所謂（頁面可能一直有背景請求），照樣往下查我們自己的。
    try:
        page.wait_for_load_state("networkidle", timeout=15000)
    except PlaywrightTimeoutError:
        lines.append("（頁面沒有進入 networkidle，不影響查詢，繼續。）")

    session = page.evaluate(SESSION_JS)
    bid = session.get("branch_id")
    cid = session.get("cust_id")

    if not bid or not cid:
        lines.append("sessionStorage 裡沒有 branch_id / cust_id —— 這個帳號很可能沒有登入成功。")
        lines.append(f"目前網址: {page.url}")
        return None, lines, dumps

    acc_code = f"1{bid}-{cid}"
    lines.append(f"帳號代碼 = {acc_code}   帳戶名 = {session.get('account')}")
    lines.append(f"sessionStorage 的 key: {session.get('keys')}")
    lines.append(f"目前網址: {page.url}")

    queries = [
        (
            "未實現損益",
            "queryInstantAccount_new",
            {"branchId": "1" + bid, "custId": cid, "range": "stksum,stkdat", "stock_no": ""},
            summarize_pnl,
        ),
        (
            "當日淨收付",
            "queryInstantAccount_new",
            {
                "branchId": "1" + bid, "custId": cid,
                "his": "y", "queryType": "1",
                "range": "stksum,stkdat,matsum,matdat", "stock_no": "",
            },
            summarize_settlement,
        ),
        (
            # 2026/08/19 從對帳單查詢頁（type=2）挖到的。下拉選單的 value：
            #   1=當日  6=前一日  2=近三日  4=本月  5=上月  0=指定日期
            # 而那頁的 JS 確實把它填進 paramInfo 的 queryType（旁邊還有 startDate/endDate，
            # 只有選「指定日期」時才有值）。所以昨日淨收付不必另找一支 API，
            # 也不必自己算交易日 —— 前一日由伺服器決定，連假不會錯開。
            #
            # 欄位刻意跟頁面帶得一模一樣（含兩個空的日期），少帶欄位時伺服器怎麼解讀
            # 是沒人知道的事，而這支查詢錯了的代價是整天的淨收付默默算成 0。
            "昨日淨收付",
            "queryInstantAccount_new",
            {
                "branchId": "1" + bid, "custId": cid,
                "his": "y", "queryType": "6",
                "startDate": "", "endDate": "",
                "range": "stksum,stkdat,matsum,matdat", "stock_no": "",
            },
            summarize_settlement,
        ),
    ]

    last_trans, trans_lines = query_trans_date(page, bid, cid)
    lines.extend(trans_lines)

    for title, cmd, param_info, summarize in queries:
        lines.append("")
        lines.append(f"--- {title} ---")
        lines.append(f"送出參數: {json.dumps(param_info, ensure_ascii=False)}")

        data, raw = query(page, cmd, param_info)
        if data is None:
            lines.append(f"查詢失敗：{raw[:500]}")
            dumps[title] = raw
            continue

        dumps[title] = data
        try:
            lines.extend(summarize(data, acc_code))
        except Exception:
            # 摘要只是輔助，原始 JSON 已經存下來了，不該因為摘要出錯就中斷整個偵察。
            lines.append("整理摘要時出錯（原始 JSON 已存檔，可直接看檔案）：")
            lines.append(traceback.format_exc())

    today = datetime.now().date()
    span, span_lines = range_query(
        page, bid, cid, today - timedelta(days=LOOKBACK_DAYS), today)
    lines.extend(span_lines)
    if span is not None:
        dumps["指定區間"] = span

    # 一定排在最後：它會把分頁導去 type=12，上面那幾個查詢就沒地方跑了。
    bank_lines, bank_dumps = recon_bank(page, bid, cid, acc_code)
    lines.extend(bank_lines)
    dumps.update(bank_dumps)

    lines.extend(method_two(dumps, span, today))

    return acc_code, lines, dumps


def pick_accounts(accounts, argv):
    """
    依命令列參數挑要偵察哪幾組帳號。

    不帶參數 = 全部；帶數字 = 只做第幾組（從 1 開始），例如 `python recon.py 1`。
    這樣要先單獨試一組時就不必去動 .env，免得改壞了帳密設定。
    """
    numbered = list(enumerate(accounts, start=1))
    if len(argv) < 2:
        return numbered

    try:
        which = int(argv[1])
    except ValueError:
        print(f"參數要是數字（第幾組帳號），收到的是: {argv[1]}")
        sys.exit(1)

    if not 1 <= which <= len(accounts):
        print(f".env 裡目前有 {len(accounts)} 組帳號，第 {which} 組不存在。")
        sys.exit(1)

    return [numbered[which - 1]]


def main():
    accounts = load_accounts()
    if not accounts:
        print(f"找不到帳號設定。請在 {app_dir()} 放一個 .env 檔（可複製 .env.example）。")
        sys.exit(1)

    total = len(accounts)
    selected = pick_accounts(accounts, sys.argv)
    if len(selected) < total:
        print(f".env 裡有 {total} 組帳號，這次只偵察第 {selected[0][0]} 組。")

    out_dir = app_dir() / OUTPUT_DIR_NAME
    out_dir.mkdir(exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M")

    configure_browsers_path()

    report = [
        f"偵察時間: {datetime.now():%Y/%m/%d %H:%M:%S}",
        f"本次偵察組數: {len(selected)}（.env 裡共 {total} 組）",
        "注意：收盤結帳（每日約 17~20 點）前後的「當日淨收付」可能不一樣，記下這份是幾點跑的。",
    ]

    with sync_playwright() as p:
        context, browser = open_context(p)
        spare_page = context.pages[0] if context.pages else None

        for index, account in selected:
            tbb_id = account["id"]
            report.append("")
            report.append("=" * 70)
            report.append(f"第 {index} 組帳號")
            report.append("=" * 70)

            try:
                page = do_login(context, tbb_id, account["password"], spare_page)
                spare_page = None
                # 關鍵：登入完就立刻抓。等下一個帳號登入後，這個的 session 可能已經被頂掉。
                acc_code, lines, dumps = recon_one(page)
            except PlaywrightTimeoutError:
                report.append("登入逾時，找不到欄位，網站版面可能已變更。")
                continue
            except PlaywrightError as exc:
                report.append(f"瀏覽器操作失敗：{exc}")
                continue

            report.extend(lines)

            # 檔名用帳號代碼（1112-0108640），不用身分證字號。
            for title, payload in dumps.items():
                name = f"{stamp}_{acc_code or f'第{index}組'}_{title}.json"
                path = out_dir / name
                if isinstance(payload, str):
                    path = path.with_suffix(".txt")
                    path.write_text(payload, encoding="utf-8")
                else:
                    path.write_text(
                        json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8"
                    )
                report.append(f"已存檔: {path.name}")

        report_path = out_dir / f"{stamp}_摘要.txt"
        report_path.write_text("\n".join(report), encoding="utf-8")

        print()
        print("\n".join(report))
        print()
        print("=" * 70)
        print(f"原始資料與摘要都在: {out_dir}")
        print("這些檔案是你的帳務資料，不要隨手外流（已加進 .gitignore）。")
        print("=" * 70)
        print("瀏覽器留著不關，你可以直接對照畫面上的數字。")

        wait_until_finished(context)

        try:
            context.close()
            if browser is not None:
                browser.close()
        except PlaywrightError:
            pass


if __name__ == "__main__":
    try:
        main()
    except SystemExit as exc:
        if exc.code:
            pause("按 Enter 關閉視窗...")
        raise
    except Exception:
        traceback.print_exc()
        pause("發生未預期的錯誤，按 Enter 關閉視窗...")
        sys.exit(1)
