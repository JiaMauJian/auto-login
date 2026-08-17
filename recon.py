"""
唯讀偵察腳本：登入後把「未實現損益」與「當日淨收付」的原始回應抓下來存檔，什麼都不寫。

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

因為第 2 點，這裡刻意採用「每登入完一組帳號就立刻抓完它的資料」的順序，
而不是全部登入完再回頭抓 —— 後者前面幾個帳號的 session 可能早就被頂掉了。

抓資料的方式是在已登入的分頁裡呼叫網站自己的 AJAX（POST /tbb/MainController），
沿用頁面現成的 B64_XOR_Encode / XOR_KEY 與 cookie，不做畫面解析。

輸出在 偵察資料\ 資料夾（已加進 .gitignore）。裡面是你的帳務資料，不要隨手外流。
"""

import json
import sys
import traceback
from datetime import datetime

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
    ]

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
