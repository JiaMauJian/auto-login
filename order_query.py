"""
掛單查詢的正式版：`queryOrder`（委託查詢，`order/layoutRWD.jsp?type=2`）。

跟 `order_recon.py` 的關係，就是 `fetch.py` 跟 `recon.py` 的關係——偵察腳本
留著當偵察腳本（存原始 JSON、印摘要、有自己的 main()），這裡只做掛單分頁要
用的事：在一個已經登入的分頁上查一次，把回應整理成畫面直接畫得出來的列。

判讀方式不是猜的：`errcode` 不是 "00000000" 就是失敗、成功才看 `matqty` 與
`orgqty` 決定完全／部分成交，這一份是 2026/08/28 從「委託查詢」頁自己的
`fakeJSON()` 抄出來、寫進 `order_recon.describe_outcome` 的，這裡直接沿用
同一支函式——同一套判讀只能有一份，兩邊各寫一次遲早會分岔。

paramInfo 有兩個容易猜錯的地方（見 `order_recon.py` 開頭）：`branchId`
**不加** '1' 前綴（跟 `query610`／`queryBankBalance` 不一樣），而且鍵名是
`cust_id` 不是 `custId`。這兩條錯了都不會報錯，只會查不到東西。

**一組登入完就立刻查那一組**，不要先把全部帳號登入完再回頭查（見
`fetch.ensure_logged_in` 的警告）：整個瀏覽器只有一組 cookie，全部登入完之後
它是最後一組的，回頭查前面幾組會拿到最後那一組的資料。
"""

from playwright.sync_api import Error as PlaywrightError, TimeoutError as PlaywrightTimeoutError

from fetch import account_code
from order_recon import (APCODE_NAMES, BS_FLAG_NAMES, BUYSELL_NAMES, ORDER_PAGE,
                         describe_outcome)
from recon import account_codes, query
from util import to_num

# 委託查詢頁「可以打 AJAX 了」的條件：加密參數用的 common.js 全域函式真的載好了。
# 刻意不等 networkidle——這個網站的頁面有背景請求一直在跑，等它安靜是在等一件我們
# 不在乎的事（fetch._open_account_page 量過：0.90 秒 vs 0.33 秒）。
PAGE_READY_JS = """
() => typeof B64_XOR_Encode === 'function' && typeof XOR_KEY !== 'undefined'
"""


def _int(value):
    """網站回的數量欄位都是字串，空字串、None、非數字一律當 0。"""
    try:
        return int(str(value or "0").strip() or 0)
    except ValueError:
        return 0


def normalize(row, sheet):
    """
    把 `queryOrder` 的一列整理成掛單分頁直接畫得出來的欄位。

    「未成交」是自己算的（原委託 − 成交 − 取消）：網站回的欄位裡沒有這個數字，
    但那正是「現在還掛在外面多少」，也是取消掛單真正會動到的量（9.7 第 3 步）。

    `open` 是「這一列還掛在外面、取消得掉」：委託本身沒失敗，而且還有沒成交也
    沒被取消的量。第 3 步的三顆取消按鈕就是照這個旗標挑要送哪幾筆，這一步先
    只拿它決定畫面上要不要淡化顯示。
    """
    org, mat, cel = _int(row.get("orgqty")), _int(row.get("matqty")), _int(row.get("celqty"))
    left = max(org - mat - cel, 0)
    ok = str(row.get("errcode") or "") == "00000000"
    side = (row.get("buysell") or "").strip()
    return {
        "sheet": sheet,
        "code": (row.get("stockno") or "").strip(),
        "side": side,
        "side_text": BUYSELL_NAMES.get(side, side),
        "kind_text": APCODE_NAMES.get(row.get("apcode"), row.get("apcode") or ""),
        "flag_text": BS_FLAG_NAMES.get(row.get("bs_flag"), row.get("bs_flag") or ""),
        "price": to_num(row.get("odprice"), None),
        "qty": org,
        "matched": mat,
        "cancelled": cel,
        "left": left,
        "status": describe_outcome(row),
        "ordno": (row.get("ordno") or "").strip(),
        "open": ok and left > 0,
        # 取消掛單要用到的欄位還沒全部確認（那支 API 還沒偵察過），先把原始那一列
        # 整份留著，免得第 3 步發現少帶了什麼又要回頭改這裡的形狀。
        "raw": row,
    }


def query_orders(page, session, sheet):
    """
    在一個已登入的分頁上查一次委託單，回傳整理過的列（一列一張委託單）。

    先導到「委託查詢」頁再查：那是這支 CMD 原本的來源頁，也是 order_recon.py
    2026/08/28 實際驗過的那條路，不要省。

    回的是**今天所有的委託**，不只還掛著的那幾筆——掛單分頁同時是「自動送出到底
    送了什麼」的唯一驗證面（見 docs/介面規劃.md 10.1），已經成交或已經取消的也
    要看得到。哪幾筆還掛在外面看 `open`。
    """
    bid, cid = (session or {}).get("branch_id"), (session or {}).get("cust_id")
    if not bid or not cid:
        raise RuntimeError(f"{sheet}：沒有登入成功（sessionStorage 沒有帳號資料）")

    page.goto(ORDER_PAGE, wait_until="domcontentloaded")
    try:
        page.wait_for_function(PAGE_READY_JS, timeout=15000)
    except (PlaywrightError, PlaywrightTimeoutError):
        pass

    data, raw = query(page, "queryOrder", {
        "branchId": bid,     # 不加 '1' 前綴，跟其他查詢不一樣
        "cust_id": cid,      # 鍵名是 cust_id，不是 custId
        "stock_no": "",
        "apcode": "0",
        "market": "0",
        "qry_type": "0",
    })
    if data is None:
        raise RuntimeError(f"{sheet}：委託查詢失敗　{str(raw)[:200]}")
    if data.get("retcode") != "000000":
        raise RuntimeError(f"{sheet}：委託查詢回應異常　{data.get('retcode')} {data.get('retmsg')}")

    # 每一列都帶 bhno/cseq，跟這個分頁登入的身分對不上就是 session 被別人頂掉了
    # （同 fetch.collect 那道核對）。這一頁是「自動送出到底送了什麼」的驗證面，
    # 把別人的委託掛在這個人名下顯示，比查不到還糟。
    expect = account_code(session)
    codes = account_codes(data)
    if codes and expect and codes != {expect}:
        raise RuntimeError(
            f"{sheet}：查到的委託屬於 {'、'.join(sorted(codes))}，與登入的 {expect} 不符"
            f"（session 可能被其他帳號頂掉）")

    return [normalize(row, sheet) for row in (data.get("ack") or [])]
