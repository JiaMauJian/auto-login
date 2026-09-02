"""
取消預約單的正式版：操作「預約查詢」頁（`order/layoutRWD.jsp?type=4`）。

跟 `order_cancel.py` 的關係，就是 `order_recon.py` 跟 `order_query.py` 彼此的
關係——同一個網站機制、不同頁面，各自一支檔案。委託單刪不掉預約單
（`order_cancel.py` 那一頁的 `orderMode` 寫死 `'2'`），這一支就是補那一塊：
「取消全部買/賣/掛單」三顆按鈕本來就該兩種單都涵蓋（見
docs/自動下單與半自動下單規劃.pptx.txt 81-85 行，原始規劃沒有分開兩顆），
不是新功能，是把漏掉的一半補上。

2026/09/02 用 `order_cancel_reservation_recon.py` 對真帳號、真的一筆預約單
（P0638918）偵察過兩輪，這支裡面每一條選擇器都是照那次看到的東西寫的：

- 逐列的「刪除」鈕是 `.delRow`，點下去**不會送出任何 AJAX**——只是把
  `orderObj` 塞進 `parent.orderArray`、`parent.mod='3'`，再 `layer.open(...)`
  開視窗，安全等級跟委託查詢頁的「終止委託單」開視窗一樣。
- 開出來的 iframe 是 `orderConfirmRWD.html`，**跟 order_cancel.py／
  order_fill.py 用的是同一份**，`#submit`／`#cancel`／`#result0` 選擇器直接
  沿用（偵察資料\委託確認視窗\orderConfirmRWD.html 已存檔核對過原始碼）。
  那份原始碼裡兩件事值得記住：`mod=='3'`（刪單）時它自己會把 `chgqty` 覆寫成
  `'0'` 才送出去（不是我們要填對的值，晚填也沒關係）；送出去的 CMD 固定是
  `modifyOrder`，`orderMode` 直接讀 `orderObj.ordstatus`——`.delRow` 那段
  固定填 `'1'`，這就是委託單跟預約單在同一支 CMD 底下分岔的地方。
- 這一頁真正的識別欄位是 **`preordno`（預約書號）**，不是 `ordno`——那個
  欄位在 queryOrder 回應裡對預約單是空字串（對照
  偵察資料\20260828_1055_..._委託查詢.json 確認：同一支 CMD 的回應兩個欄位
  一直都在，只是依 ordstatus 條件性地一個有值一個是空的）。`order_query.py`
  的 `normalize()` 已經改成看 ordstatus 選欄位，這裡收到的 `ordnos` 就是
  那個修正後的值，等於是 `preordno`。
- 這頁也有勾選＋`#openConfirm`「終止委託單」的批次路，但它開的是**另一份**
  `orderConfirm.html`（沒有 RWD），完全沒偵察過，跟 `.delRow` 不是同一個
  iframe。**這支刻意不碰它**，一律走 `.delRow`、一筆一次——預約單量體通常
  很小（一個帳戶頂多幾筆），換取不用再猜一份沒看過的確認視窗。
"""

from playwright.sync_api import Error as PlaywrightError, TimeoutError as PlaywrightTimeoutError

from order_cancel import close_dialog
from order_fill import CONFIRM_IFRAME_SELECTOR, OrderMaybeSubmitted
from order_query import PAGE_READY_JS
from order_recon import RESERVE_PAGE

# 刪單確認視窗裡那顆真的會送出去的按鈕。整支檔案只有 cancel_orders() 會點它，
# 而且點之前一定先跑完 _verify_dialog()——跟 order_cancel.py 同一條界線。
SUBMIT_BUTTON = "#submit"

# 頁面自己那張表：一個 <tbody id="bar<i>"> 一筆預約單。欄位位置抄自頁面
# fakeJSON() 畫表格那段（td:eq(3) 是 preordno、td:eq(4) 股票、td:eq(6) 買賣別），
# 跟 order_cancel.DUMP_ROWS_JS 同一套邏輯，只是這裡沒有 checkbox 可讀，改讀
# .delRow 按鈕在不在——按鈕只在 celable=='1' 時才畫得出來。
DUMP_ROWS_JS = """
() => [...document.querySelectorAll('#qOrderTable tbody[id^="bar"]')].map((tb) => {
    const tds = tb.querySelectorAll('tr td');
    const text = (i) => tds[i] ? tds[i].innerText.trim().replace(/\\s+/g, ' ') : '';
    return {
        bar: tb.id,
        cancellable: !!tb.querySelector('.delRow'),
        ordno: text(3),   // 這一欄印的其實是 preordno（預約書號）
        code: text(4).split(' ')[0],
        side: text(6),
    };
})
"""

# .delRow 點下去之後，orderConfirmRWD.html 讀的是 parent.orderArray——核對
# 送出去前的最後一道防線，跟 order_cancel._verify_dialog 同一個目的，只是
# 這裡固定一次一筆。
DUMP_ORDER_ARRAY_JS = """
() => (typeof orderArray === 'undefined') ? null : JSON.parse(JSON.stringify(orderArray))
"""


def _open_page(page):
    """導到預約查詢頁，等它把表格畫出來。回傳頁面自己那張表的每一列。"""
    page.goto(RESERVE_PAGE, wait_until="domcontentloaded")
    try:
        page.wait_for_function(PAGE_READY_JS, timeout=15000)
    except (PlaywrightError, PlaywrightTimeoutError):
        pass

    try:
        page.wait_for_function(
            "() => document.querySelectorAll('#qOrderTable tbody[id^=\"bar\"]').length > 0",
            timeout=15000)
    except (PlaywrightError, PlaywrightTimeoutError):
        pass

    return page.evaluate(DUMP_ROWS_JS)


def _verify_dialog(page, expected_ordno, sheet, session):
    """
    按確認之前的最後一道核對，跟 order_cancel._verify_dialog 同樣三件事，只是
    這裡一次只處理一筆：視窗上剛好 1 筆、預約書號對得上、客戶帳號是這個分頁
    登入的人。對不上就丟 RuntimeError：呼叫端會去按「取消」關掉視窗，這一筆
    不送。
    """
    order_array = page.evaluate(DUMP_ORDER_ARRAY_JS)
    if not isinstance(order_array, list) or len(order_array) != 1:
        raise RuntimeError(
            f"{sheet}：刪單確認視窗上讀到 "
            f"{0 if not isinstance(order_array, list) else len(order_array)} 筆，"
            f"預約單一次只處理一筆，先不要送出。")

    item = order_array[0]
    got = str(item.get("ordno") or "").strip()
    if got != expected_ordno:
        raise RuntimeError(
            f"{sheet}：刪單確認視窗上是 {got or '（空白）'}，"
            f"跟要刪的 {expected_ordno} 對不起來，沒有送出。")

    cid = str((session or {}).get("cust_id") or "").strip()
    got_cid = str(item.get("custId") or "").strip()
    if cid and got_cid and got_cid != cid:
        raise RuntimeError(
            f"{sheet}：刪單確認視窗上的委託屬於 {got_cid}，與登入的 {cid} 不符"
            f"（session 可能被其他帳號頂掉），沒有送出。")

    return got


def _read_result0(page, timeout_ms):
    """只有一筆，讀 #result0 就夠，不用 order_cancel._read_results 那套逐列迴圈。"""
    frame = page.frame_locator(CONFIRM_IFRAME_SELECTOR)
    text = ""
    for _ in range(max(timeout_ms // 200, 1)):
        try:
            text = (frame.locator("#result0").inner_text() or "").strip()
        except PlaywrightError:
            text = ""
        if text:
            break
        page.wait_for_timeout(200)
    return text


def cancel_orders(page, session, sheet, ordnos, timeout_ms=20000):
    """
    在一個已登入的分頁上，把 `ordnos`（預約書號）這幾筆預約單刪掉，回傳一列一筆
    的結果，形狀跟 `order_cancel.cancel_orders` 一模一樣：

        {"results": [{"ordno": "P0638918", "code": "2002", "side": "賣出",
                      "ok": True, "message": "刪單成功"}, ...],
         "missing": [...], "locked": [...]}

    一次一筆：開視窗 → 核對 → 送出 → 關視窗，不是勾多筆一次送（這頁能勾多筆的
    那條路開的是另一份沒偵察過的確認視窗，見檔頭說明）。**每一筆都重新導頁重讀
    一次表格**——上一筆刪成功後 orderConfirmRWD.html 自己會呼叫
    `parent.renderTable()` 把表重畫，`bar{i}` 的索引會跟著位移，沿用舊的 `bar`
    id 會點錯列。

    例外的兩種意思跟 order_cancel.cancel_orders 一樣：
      - `RuntimeError`：發生在按下確認之前，這一筆沒送出，重試安全。
      - `OrderMaybeSubmitted`：確認已經按下去了，這一筆多半已經送到券商，不要
        自動重來，回報給人、重查掛單。
    """
    wanted = [str(o).strip() for o in ordnos if str(o).strip()]
    if not wanted:
        raise RuntimeError(f"{sheet}：沒有指定要刪哪一筆。")

    results, missing, locked = [], [], []
    for ordno in wanted:
        rows = _open_page(page)
        row = next((r for r in rows if r["ordno"] == ordno), None)
        if row is None:
            missing.append(ordno)
            continue
        if not row["cancellable"]:
            locked.append(ordno)
            continue

        page.locator(f"#{row['bar']} .delRow").click()
        page.locator(".layui-layer-title", has_text="刪單確認").wait_for(state="visible", timeout=10000)

        try:
            sent = _verify_dialog(page, ordno, sheet, session)
        except RuntimeError:
            close_dialog(page)   # 這一筆沒送，把視窗收乾淨再把例外丟出去
            raise

        frame = page.frame_locator(CONFIRM_IFRAME_SELECTOR)
        frame.locator(SUBMIT_BUTTON).click()
        # ↑ 過了這一行就沒有回頭路了：以下任何失敗都是 OrderMaybeSubmitted。

        text = _read_result0(page, timeout_ms)
        close_dialog(page)

        if not text:
            raise OrderMaybeSubmitted(
                f"{sheet}：已經按下刪單確認視窗的「確認」（預約書號 {sent}），"
                f"但畫面上沒出現結果。這一筆可能已經送出去了，不要再按一次——"
                f"請重查掛單，用查回來的結果為準。")

        results.append({
            "ordno": ordno,
            "code": row["code"],
            "side": row["side"],
            "ok": "刪單成功" in text,
            "message": text,
        })

    return {"results": results, "missing": missing, "locked": locked}
