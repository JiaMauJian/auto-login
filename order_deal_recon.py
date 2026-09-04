"""
唯讀偵察腳本：登入後查一次「成交查詢」（CMD=queryDealOrder），把今天的成交
彙總列出來，什麼都不寫、不下單、不改單、不取消單。

跟 order_recon.py（委託查詢，queryOrder）同樣的定位與同一套呼叫方式，查的
目標換成「這筆委託到底成交了沒、成交多少」——用來驗證能不能拿 queryDealOrder
取代「比對持股有沒有變」去判斷一輪委託有沒有真的成交，讓買賣股票／出清股票／
全持股交易這三個作業對這件事的判斷方式一致（見對話討論）。

queryDealOrder 的 paramInfo 是從「成交查詢」頁（order/layoutRWD.jsp?type=3）
自己的原始碼裡 renderTable() 挖出來的，不是猜的：

    var Param = JSON.stringify({
        "branchId": brokerId,   // sessionStorage 的 branch_id，不加 '1' 前綴
        "cust_id": custId,      // 注意鍵名是 cust_id，跟 queryOrder 一樣
        "stock_no": "",         // 空字串＝全部股票；renderDetail() 查單一檔時才會帶代號
        "apcode": "0",          // 全部盤別
        "market": "0",          // 全部市場
        "qry_type": "1",        // 0=成交明細（一筆成交一列）　1=成交彙總（一個 ordno 一列）
    });

qry_type=1（彙總）回應的形狀（rtn.mat 陣列的一列）：

    {"market","buysell","match_qty","ordno","stock_no","apcode","trade",
     "mat_date","payment","match_time","avg_price","source"}

這正是要判斷「這個 ordno 成交了沒、成交多少股」最直接的答案——一個 ordno
一列、match_qty 是總成交股數，不必像頁面上「查詢單一檔明細」（qry_type=0，
renderDetail()）那樣還要在 Python 端自己過濾 stock_no+ordno。

qry_type=0（明細）回應欄位不太一樣（用的是 pay_price／price 不是
payment／avg_price，多一個 mkt_seq_num），這支腳本兩種都查一次存檔，方便
對照，但真正要拿來判斷成交與否的是 qry_type=1 那份。

輸出在 偵察資料\\ 資料夾（已加進 .gitignore）。
"""

import json
import sys
from datetime import datetime

from playwright.sync_api import Error as PlaywrightError, TimeoutError as PlaywrightTimeoutError
from playwright.sync_api import sync_playwright

from login import app_dir, configure_browsers_path, do_login, load_accounts, open_context, pause, wait_until_finished
from order_recon import APCODE_NAMES, BS_FLAG_NAMES, BUYSELL_NAMES, SESSION_JS
from recon import OUTPUT_DIR_NAME, check_account, pick_accounts, query

# 成交查詢頁。一定要先導到這裡：跟其他 order/ 底下的頁面一樣載了 common.js
# （B64_XOR_Encode 在裡面），而且是這支 CMD 原本的來源頁。
DEAL_PAGE = "https://www.tbbstock.com.tw/tbb/order/layoutRWD.jsp?type=3"


def describe_summary_row(row):
    """彙總（qry_type=1）一列 -> 一行摘要文字。"""
    apcode = APCODE_NAMES.get(row.get("apcode"), row.get("apcode"))
    buysell = BUYSELL_NAMES.get(row.get("buysell"), row.get("buysell"))
    return (
        f"  {row.get('stock_no')} {buysell} {apcode} 委託書號={row.get('ordno')}"
        f" 成交股數={row.get('match_qty')} 成交均價={row.get('avg_price')}"
        f" 收付金額={row.get('payment')} 成交日={row.get('mat_date')}"
        f" 市場={row.get('market')} 來源={row.get('source')}"
    )


def describe_detail_row(row):
    """明細（qry_type=0）一列 -> 一行摘要文字。"""
    apcode = APCODE_NAMES.get(row.get("apcode"), row.get("apcode"))
    buysell = BUYSELL_NAMES.get(row.get("buysell"), row.get("buysell"))
    bs_flag = BS_FLAG_NAMES.get(row.get("bs_flag"), row.get("bs_flag"))
    return (
        f"  {row.get('stock_no')} {buysell} {apcode} 條件={bs_flag}"
        f" 委託書號={row.get('ordno')} 成交股數={row.get('match_qty')}"
        f" 成交價={row.get('price')} 應收付={row.get('pay_price')}"
        f" 成交時間={row.get('mat_date')} {row.get('match_time')}"
        f" 市場序號={row.get('mkt_seq_num')}"
    )


def query_deal_order(page, bid, cid, qry_type, stock_no=""):
    """
    呼叫一次 queryDealOrder，回傳 (data 或 None, 摘要文字列表)。

    qry_type：'1' 彙總（一個 ordno 一列）、'0' 明細（一筆成交一列）。
    stock_no：空字串是全部股票，明細那條路頁面自己查單一檔時才會帶代號
    （見 renderDetail()），這裡兩種都用空字串查全部，比對起來比較方便。
    """
    label = "成交彙總" if qry_type == "1" else "成交明細"
    lines = ["", f"--- 成交查詢（type=3, queryDealOrder, qry_type={qry_type} {label}）---"]
    param_info = {
        "branchId": bid,     # 不加 '1' 前綴，跟 queryOrder 一樣
        "cust_id": cid,      # 注意鍵名是 cust_id
        "stock_no": stock_no,
        "apcode": "0",
        "market": "0",
        "qry_type": qry_type,
    }
    lines.append(f"送出參數: {json.dumps(param_info, ensure_ascii=False)}")

    data, raw = query(page, "queryDealOrder", param_info)
    if data is None:
        lines.append(f"查詢失敗：{str(raw)[:500]}")
        return None, lines

    lines.append(f"retcode={data.get('retcode')} retmsg={data.get('retmsg')}")
    lines.append(f"回應最上層欄位: {sorted(data.keys())}")
    lines.extend(check_account(data, None))

    rows = data.get("mat") or []
    lines.append(f"筆數={len(rows)}")
    if not rows:
        lines.append("（今天沒有成交）")
        return data, lines

    describe = describe_summary_row if qry_type == "1" else describe_detail_row
    for row in rows:
        lines.append(describe(row))
    return data, lines


def recon_one(page):
    """在已登入的分頁上查一次成交彙總與明細。回傳 (帳號代碼, 摘要文字, 要存檔的原始資料)。"""
    page.goto(DEAL_PAGE)
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

    dumps = {}
    summary_data, summary_lines = query_deal_order(page, bid, cid, "1")
    lines.extend(summary_lines)
    if summary_data is not None:
        dumps["成交彙總"] = summary_data

    detail_data, detail_lines = query_deal_order(page, bid, cid, "0")
    lines.extend(detail_lines)
    if detail_data is not None:
        dumps["成交明細"] = detail_data

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

    report = [f"偵察時間: {datetime.now():%Y/%m/%d %H:%M:%S}（成交查詢）"]

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

        report_path = out_dir / f"{stamp}_成交查詢摘要.txt"
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
