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
from order_recon import (BS_FLAG_NAMES, BUYSELL_NAMES, ORDER_PAGE, PRICE_FLAG_NAMES,
                         describe_outcome, describe_trade)
from recon import account_codes, query
from util import show, to_num

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


def _date(text):
    """網站的 yyyymmdd → yyyy/mm/dd。不是 8 位數字就原樣回去，不要自己補。"""
    value = str(text or "").strip()
    if len(value) == 8 and value.isdigit():
        return f"{value[:4]}/{value[4:6]}/{value[6:8]}"
    return value


def _time(text):
    """
    網站的 hhmmssSSS → hh:mm:ss。

    毫秒不印：網站畫面上有（09:00:43.783），但那三位數是用來分辨「同一秒送出的
    兩筆誰先誰後」的，人在這一頁核對的是「這張單是幾點下的」，多三位只是變窄。
    """
    value = str(text or "").strip()
    if len(value) >= 6 and value.isdigit():
        return f"{value[:2]}:{value[2:4]}:{value[4:6]}"
    return value


def normalize(row, sheet):
    """
    把 `queryOrder` 的一列整理成掛單分頁直接畫得出來的欄位。

    欄位是照網站「委託查詢」那張表一欄一欄對過來的（2026/08/31 使用者要求兩邊
    對得起來，才好拿程式的表跟網頁的表互相核對），只有最前面的「帳戶」是這裡
    多的一欄——網站一次只看得到登入的那一個人，程式是所有帳戶攤在同一張表上，
    不標名字就分不出誰是誰。

    「有效數量」是自己算的（原委託 − 已成交 − 已取消）：網站回的欄位裡沒有這個
    數字（`celable` 不是數量，那筆 IOC 失敗單的 orgqty 是 1、celable 卻是 2），
    但那正是「現在還掛在外面多少」，也是取消掛單真正會動到的量（9.7 第 3 步）。

    `open` 是「這一列還掛在外面、取消得掉」，用的是網站自己的判斷：`celable`
    等於 `'1'` 才畫得出那一列的勾選框（見「委託查詢」頁 renderTable 裡的
    celBox，2026/08/31 偵察確認）。**不要改回自己算「原委託 − 已成交 − 已取消
    > 0」**——那個算式目前結論一樣，但它是推的，`celable` 是網站給的答案，而
    取消掛單那三顆按鈕就是照這個旗標挑要送哪幾筆（10.3 第一點）。順便擋掉
    errcode 失敗的那幾列：那種列本來就沒有勾選框。

    這個旗標不分委託單／預約單，兩種都可能是 `True`——「取消全部買/賣/掛單」
    三顆按鈕本來就該兩種都能取消（見 docs/自動下單與半自動下單規劃.pptx.txt
    81-85 行，2026/09/02 使用者確認不是分開兩顆），差別只在 `ordstatus` 決定
    要走哪一支取消函式（`order_cancel.py` 或 `order_cancel_reservation.py`），
    不是要不要讓它出現在清單上。
    """
    org, mat, cel = _int(row.get("orgqty")), _int(row.get("matqty")), _int(row.get("celqty"))
    left = max(org - mat - cel, 0)
    ok = str(row.get("errcode") or "") == "00000000"
    cancellable = str(row.get("celable") or "").strip() == "1"
    side = (row.get("buysell") or "").strip()
    ordstatus = str(row.get("ordstatus") or "").strip()
    # 預約單（ordstatus=='1'）在這支 CMD 的回應裡 ordno 是空的，真正的識別
    # 欄位是 preordno（預約書號，網站「預約查詢」頁「委託書號」欄印的就是這個）
    # ——兩個欄位一直都在，只是依 ordstatus 條件性地一個有值一個是空字串
    # （2026/09/02 對照真實回應 偵察資料\20260828_1055_..._委託查詢.json
    # 與預約查詢頁原始碼確認）。取消掛單認的就是這個欄位，不分型態都要對得上。
    ordno = (row.get("preordno") if ordstatus == "1" else row.get("ordno")) or ""
    # 市價／漲跌停單的 odprice 是 0，那一欄直接印 0 會看起來像「委託價 0 元」，
    # 所以非限價的改印價格種類（網站自己也是這樣分的，見 PRICE_FLAG_NAMES）。
    price_flag = str(row.get("priceflag") or "0")
    price = to_num(row.get("odprice"), None)
    return {
        "sheet": sheet,
        "ordered_at": f"{_date(row.get('orddate'))} {_time(row.get('ordtime'))}".strip(),
        "work_date": _date(row.get("workdate")),
        "ordno": ordno.strip(),
        "ordstatus": ordstatus,
        "code": (row.get("stockno") or "").strip(),
        "trade_text": describe_trade(row),
        "side": side,
        "side_text": BUYSELL_NAMES.get(side, side),
        "flag_text": BS_FLAG_NAMES.get(row.get("bs_flag"), row.get("bs_flag") or ""),
        "price": price,
        "price_text": (show(price) if price is not None else "") if price_flag == "0"
                      else PRICE_FLAG_NAMES.get(price_flag, price_flag),
        "qty": org,
        "matched": mat,
        "cancelled": cel,
        "left": left,
        "status": describe_outcome(row),
        "open": ok and cancellable,
        # 原始那一列整份留著：取消掛單認的是委託書號（ordno），但真的出事要對帳
        # 的時候，errcode/celable/ordstatus 這些沒進欄位的值就在這裡面。
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
