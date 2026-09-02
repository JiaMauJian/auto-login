"""
唯讀偵察腳本：登入後查一次「委託查詢」（CMD=queryOrder），把今天的委託單
列出來，什麼都不寫、不下單、不改單、不取消單。

跟 recon.py 同樣的定位與同一套呼叫方式（page.evaluate 呼叫頁面現成的
B64_XOR_Encode，不在 Python 端重做那層加密，見 docs 與記憶
「下單 Param 是 RSA 不是 XOR」），只是查的目標換成委託單狀態，用來驗證
半自動下單流程送出後，程式能不能自己讀到「這張單後來怎麼樣了」。

queryOrder 的 paramInfo 是從「委託查詢」頁（order/layoutRWD.jsp?type=2）
自己的 renderTable() 挖出來的，兩個容易猜錯的地方特別記一下：

    branchId 不加 '1' 前綴（跟 query610／queryBankBalance 不一樣）
    鍵名是 cust_id，不是其他頁常見的 custId

委託結果怎麼判讀，也是照同一頁 fakeJSON() 裡的邏輯抄的（不是用猜的）：
errcode 不是 "00000000" 就是失敗，訊息在 errmsg 裡；成功的話看 matqty
是否等於 orgqty（完全成交／部分成交），或看 act 是新單/改價/減量/刪單。

輸出在 偵察資料\ 資料夾（已加進 .gitignore）。
"""

import json
import sys
from datetime import datetime

from playwright.sync_api import Error as PlaywrightError, TimeoutError as PlaywrightTimeoutError
from playwright.sync_api import sync_playwright

from login import app_dir, configure_browsers_path, do_login, load_accounts, open_context, pause, wait_until_finished
from recon import OUTPUT_DIR_NAME, check_account, pick_accounts, query

# 委託查詢頁。一定要先導到這裡：它跟其他 account/ 底下的頁面一樣載了
# common.js（B64_XOR_Encode 在裡面），而且是這支 CMD 原本的來源頁。
ORDER_PAGE = "https://www.tbbstock.com.tw/tbb/order/layoutRWD.jsp?type=2"

# 預約查詢頁——**不要跟 recon.ACCOUNT_PAGE（account/layoutRWD.jsp?type=4，
# 未實現損益）搞混**，兩個都是 type=4，但目錄不一樣（order/ vs account/）、
# 內容完全無關，純粹是網站自己的頁碼剛好撞了。
RESERVE_PAGE = "https://www.tbbstock.com.tw/tbb/order/layoutRWD.jsp?type=4"

SESSION_JS = """
() => ({
    branch_id: sessionStorage.getItem('branch_id'),
    cust_id: sessionStorage.getItem('cust_id'),
    account: sessionStorage.getItem('account'),
})
"""

APCODE_NAMES = {"1": "整股", "2": "盤後", "3": "盤後零股", "4": "興櫃", "5": "盤中零股"}
BUYSELL_NAMES = {"B": "買進", "S": "賣出"}
BS_FLAG_NAMES = {"F": "FOK全部成交否則取消", "R": "ROD當日有效", "I": "IOC立即成交否則取消"}
ACT_NAMES = {"C": "刪單成功", "R": "改價成功", "M": "減量成功", "O": "委託成功"}

# 「交易別」要拿 apcode 與 trade 兩碼**合起來**查，不是各查各的：委託確認視窗
# 自己的 type{} 就是這樣寫的（var aType = orderObj.apcode + orderObj.trade，見
# 偵察資料\委託確認視窗\orderConfirmRWD.html）。單看 trade 的 "0" 會以為那一碼
# 自己就是現股，其實整張表是兩碼一組，換一種盤別就換一組碼。
TRADE_NAMES = {
    "10": "現股", "13": "融資", "14": "融券", "1A": "現沖賣",
    "20": "現股", "23": "融資", "24": "融券",
    "30": "現股", "40": "現股", "50": "現股",
}

# 委託價格的種類，抄同一份 orderConfirmRWD.html 的 priceFlag{}。程式自己送出去的
# 一律是限價（order_fill.py 的 #priceRadio 固定 "0"），但人自己在網站上下的單會
# 查回同一張表，市價單的 odprice 是 0，直接印那個 0 會看起來像「委託價 0 元」。
PRICE_FLAG_NAMES = {"0": "限價", "1": "平盤價", "2": "跌停價", "3": "漲停價", "4": "市價"}


def describe_trade(row):
    """網站「交易別」那一欄：盤別（整股／盤後…）加上現股／融資／融券。"""
    apcode = str(row.get("apcode") or "")
    kind = APCODE_NAMES.get(apcode, apcode)
    trade = TRADE_NAMES.get(apcode + str(row.get("trade") or ""), "")
    return f"{kind} {trade}".strip()


def describe_outcome(row):
    """
    這張委託單現在的結果，照「委託查詢」頁 fakeJSON() 裡的判讀邏輯抄的，
    不是用猜的——errcode 不是 "00000000" 就直接顯示 errmsg（失敗原因），
    成功的話才看成交量、act 決定要顯示哪一句。
    """
    errcode = str(row.get("errcode") or "")
    if errcode != "00000000":
        return row.get("errmsg") or f"失敗（errcode={errcode}）"

    orgqty = int(row.get("orgqty") or 0)
    matqty = int(row.get("matqty") or 0)
    if orgqty and orgqty == matqty:
        return "完全成交"
    if matqty > 0 and orgqty > matqty:
        return "部分成交"
    return ACT_NAMES.get(row.get("act"), f"未知動作（act={row.get('act')}）")


def summarize_orders(data, expect_acc=None):
    """把一次 queryOrder 回應整理成人看得懂的幾行。"""
    lines = []
    rows = data.get("ack") or []
    lines.append(f"retcode={data.get('retcode')} retmsg={data.get('retmsg')} 筆數={len(rows)}")
    lines.extend(check_account(data, expect_acc))

    if not rows:
        lines.append("（今天沒有委託單）")
        return lines

    for i, row in enumerate(rows):
        apcode = APCODE_NAMES.get(row.get("apcode"), row.get("apcode"))
        buysell = BUYSELL_NAMES.get(row.get("buysell"), row.get("buysell"))
        bs_flag = BS_FLAG_NAMES.get(row.get("bs_flag"), row.get("bs_flag"))
        lines.append(
            f"  [{i}] {row.get('stockno')} {buysell} {apcode} 委託價={row.get('odprice')}"
            f" 條件={bs_flag} 原委託={row.get('orgqty')} 成交={row.get('matqty')}"
            f" 取消={row.get('celqty')} 委託書號={row.get('ordno')}"
            f" ordstatus={row.get('ordstatus')}（1預約單/2盤中單，跟成不成交無關）"
            f" -> {describe_outcome(row)}"
        )
    return lines


def query_orders(page, bid, cid):
    """呼叫一次 queryOrder，回傳 (解析後的 dict 或 None, 摘要文字列表)。"""
    lines = ["", "--- 委託查詢（type=2, queryOrder）---"]
    param_info = {
        "branchId": bid,   # 不加 '1' 前綴，跟其他查詢不一樣
        "cust_id": cid,    # 注意鍵名是 cust_id
        "stock_no": "",
        "apcode": "0",
        "market": "0",
        "qry_type": "0",
    }
    lines.append(f"送出參數: {json.dumps(param_info, ensure_ascii=False)}")

    data, raw = query(page, "queryOrder", param_info)
    if data is None:
        lines.append(f"查詢失敗：{str(raw)[:500]}")
        return None, lines

    lines.extend(summarize_orders(data, None))
    return data, lines


def recon_one(page):
    """在已登入的分頁上查一次委託單。回傳 (帳號代碼, 摘要文字, 要存檔的原始資料)。"""
    page.goto(ORDER_PAGE)
    try:
        page.wait_for_load_state("networkidle", timeout=15000)
    except PlaywrightTimeoutError:
        pass

    session = page.evaluate(SESSION_JS)
    bid, cid = session.get("branch_id"), session.get("cust_id")

    lines = []
    if not bid or not cid:
        lines.append("sessionStorage 裡沒有 branch_id / cust_id —— 這個帳號很可能沒有登入成功。")
        return None, lines, {}

    acc_code = f"1{bid}-{cid}"
    lines.append(f"帳號代碼 = {acc_code}   帳戶名 = {session.get('account')}")

    data, order_lines = query_orders(page, bid, cid)
    lines.extend(order_lines)

    dumps = {"委託查詢": data} if data is not None else {}
    return acc_code, lines, dumps


def main():
    accounts = load_accounts()
    if not accounts:
        print(f"找不到帳號設定。請在 {app_dir()} 放一個 .env 檔（可複製 .env.example）。")
        sys.exit(1)

    selected = pick_accounts(accounts, sys.argv)

    out_dir = app_dir() / OUTPUT_DIR_NAME
    out_dir.mkdir(exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M")

    configure_browsers_path()

    report = [f"偵察時間: {datetime.now():%Y/%m/%d %H:%M:%S}（委託查詢）"]

    with sync_playwright() as p:
        context, browser = open_context(p)
        spare_page = context.pages[0] if context.pages else None

        for index, account in selected:
            report.append("")
            report.append("=" * 70)
            report.append(f"第 {index} 組帳號")
            report.append("=" * 70)

            try:
                page = do_login(context, account["id"], account["password"], spare_page)
                spare_page = None
                acc_code, lines, dumps = recon_one(page)
            except PlaywrightTimeoutError:
                report.append("登入逾時，找不到欄位，網站版面可能已變更。")
                continue
            except PlaywrightError as exc:
                report.append(f"瀏覽器操作失敗：{exc}")
                continue

            report.extend(lines)
            for title, payload in dumps.items():
                name = f"{stamp}_{acc_code or f'第{index}組'}_{title}.json"
                (out_dir / name).write_text(
                    json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8"
                )
                report.append(f"已存檔: {name}")

        report_path = out_dir / f"{stamp}_委託查詢摘要.txt"
        report_path.write_text("\n".join(report), encoding="utf-8")

        print("\n".join(report))
        print()
        print(f"原始資料與摘要都在: {out_dir}")

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
