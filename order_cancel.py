"""
取消掛單的正式版：操作「委託查詢」頁（`order/layoutRWD.jsp?type=2`）真的把單刪掉。

跟 `order_cancel_recon.py` 的關係，就是 `order_fill.py` 跟 `recon_order_form.py`
的關係——偵察腳本留著當偵察腳本（唯讀、不按確認、印報告），這裡做的是掛單分頁
那三顆按鈕真正要跑的事。整套規格與每一條為什麼，寫在 docs/介面規劃.md 10.3。

**為什麼不自己打那支 API**：刪單的 CMD 是 `modifyOrder`，送出前那一步是
`getCert()`（TWCA 憑證簽章，在瀏覽器裡做），而且送出去的欄位有一半是頁面自己從
那一列委託湊出來的（`sales`、`cust_data`、`signcode`、`plainText`）。跟 `newOrder`
卡在同一個地方，所以走頁面，不走 AJAX。

2026/08/31 10:06 用 `order_cancel_recon.py` 對真帳號實地看過一遍，這支裡面每一條
選擇器與判斷都是照那次看到的東西寫的，不是猜的：

- 逐列的勾選框是 `#qOrderTable` 底下的 `input[name=chkCel]`，`celable == '1'`
  的列才畫得出來（畫不出來就是刪不掉）。
- **一律不按全選。** 全選是 `$('input:checkbox').not(this).prop('checked', ...)`，
  設的是整份文件；而這一頁跑 `freezeTable` 會 clone 一份影子表格出來，`#openConfirm`
  收的又是整份文件的 `$('input[name=chkCel]:checked')`——按全選的結果是同一張單被
  送兩次刪單（08/31 實測：一張 J0845，視窗上出現兩列）。
- checkbox 的 `value` 是 `queryOrder` 回應陣列的**索引**，不是委託書號，而且那個
  陣列是新的在前面——多一筆新委託，全部的索引就往後移。所以這支從頭到尾用
  **委託書號**認單，先讀頁面自己那張表，再決定要勾哪幾個框。
- 這一頁只畫 `ordstatus == '2'` 的列：**預約單刪不掉**，送出去的 `orderMode`
  也是頁面寫死的 `'2'`。
"""

from playwright.sync_api import Error as PlaywrightError, TimeoutError as PlaywrightTimeoutError

from order_fill import CONFIRM_IFRAME_SELECTOR, OrderMaybeSubmitted
from order_query import PAGE_READY_JS
from order_recon import ORDER_PAGE

# 刪單確認視窗裡那顆真的會送出去的按鈕。整支檔案只有 confirm_cancel() 會點它，
# 而且點之前一定先跑完 _verify_dialog()——改這支的人請維持這條界線。
SUBMIT_BUTTON = "#submit"

# 頁面自己那張表：一個 <tbody id="bar<i>"> 一筆委託。欄位位置抄自頁面
# `#openConfirm` 的 handler（td:eq(3) 委託書號、td:eq(4) 股票、td:eq(6) 買賣別、
# td:eq(13) 有效數量），不是自己數出來的。沒有勾選框就是這一列刪不掉。
DUMP_ROWS_JS = """
() => [...document.querySelectorAll('#qOrderTable tbody[id^="bar"]')].map((tb) => {
    const tds = tb.querySelectorAll('tr td');
    const text = (i) => tds[i] ? tds[i].innerText.trim().replace(/\\s+/g, ' ') : '';
    const box = tb.querySelector('input[name=chkCel]');
    return {
        bar: tb.id,
        value: box ? box.value : null,
        cancellable: !!box,
        ordno: text(3),
        code: text(4).split(' ')[0],
        side: text(6),
        left: text(13),
    };
})
"""

# 只勾指定的那幾列，而且**先把整份文件的勾選清乾淨**——影子表格那一份也在裡面。
# 直接設 checked 不用 click()：`#openConfirm` 只讀 `:checked`，那幾個框身上沒有
# 自己的 click handler，用 click() 反而多一次 toggle 的機會。
CHECK_ROWS_JS = """
(bars) => {
    const all = [...document.querySelectorAll('input[name=chkCel]')];
    all.forEach((b) => { b.checked = false; });
    let hit = 0;
    bars.forEach((id) => {
        const box = document.querySelector('#qOrderTable #' + id + ' input[name=chkCel]');
        if (box) { box.checked = true; hit++; }
    });
    return {hit, checked: all.filter((b) => b.checked).length};
}
"""

# 按下確認會送出去的那一批。頁面自己湊的，跟我們勾了什麼中間隔了一層——
# 核對就是核對這個（10.3 第七點）。
DUMP_ORDER_ARRAY_JS = """
() => (typeof orderArray === 'undefined') ? null : JSON.parse(JSON.stringify(orderArray))
"""


def _open_page(page):
    """導到委託查詢頁，等它把表格畫出來。回傳頁面自己那張表的每一列。"""
    page.goto(ORDER_PAGE, wait_until="domcontentloaded")
    try:
        page.wait_for_function(PAGE_READY_JS, timeout=15000)
    except (PlaywrightError, PlaywrightTimeoutError):
        pass

    # 表格是 renderTable() 打完 queryOrder 才畫出來的，goto 回來的當下還是空的。
    # 等不到不代表壞掉——今天完全沒有委託的帳戶本來就一列都沒有。
    try:
        page.wait_for_function(
            "() => document.querySelectorAll('#qOrderTable tbody[id^=\"bar\"]').length > 0",
            timeout=15000)
    except (PlaywrightError, PlaywrightTimeoutError):
        pass

    return page.evaluate(DUMP_ROWS_JS)


def _verify_dialog(page, expected, sheet, session):
    """
    按確認之前的最後一道核對：視窗上那一批，跟我們要刪的那一批是不是同一批。

    要對三件事，缺一不可：
      1. 筆數一樣（重複被送兩次就死在這裡，見 10.3 第七點）；
      2. 委託書號的集合一樣（勾錯列、頁面在我們讀完之後又重畫了）；
      3. 每一筆的客戶帳號就是這個分頁登入的人（整個瀏覽器只有一組 cookie，
         換人是換 cookie——這是「會不會刪到別人的單」的最後一道防線）。

    對不上就丟 RuntimeError：呼叫端會去按「取消」關掉視窗，一筆都不送。
    """
    order_array = page.evaluate(DUMP_ORDER_ARRAY_JS)
    if not isinstance(order_array, list) or not order_array:
        raise RuntimeError(f"{sheet}：刪單確認視窗上讀不到任何委託，先不要送出。")

    got = [str(item.get("ordno") or "").strip() for item in order_array]
    if len(got) != len(expected) or sorted(got) != sorted(expected):
        raise RuntimeError(
            f"{sheet}：刪單確認視窗上是 {len(got)} 筆（{'、'.join(got)}），"
            f"跟要刪的 {len(expected)} 筆（{'、'.join(expected)}）對不起來，一筆都沒有送出。")

    cid = str((session or {}).get("cust_id") or "").strip()
    wrong = {str(item.get("custId") or "").strip() for item in order_array} - {cid}
    if cid and wrong:
        raise RuntimeError(
            f"{sheet}：刪單確認視窗上的委託屬於 {'、'.join(sorted(wrong))}，"
            f"與登入的 {cid} 不符（session 可能被其他帳號頂掉），一筆都沒有送出。")

    return got


def _read_results(page, count, timeout_ms):
    """
    逐列讀 `#result<i>`。**不是只讀 `#result0`**——一批多筆的話，每一筆有自己的
    那一格（`order_fill.confirm_order` 只讀第一格，直接沿用會把「第一筆成功」
    當成「全部成功」，見 10.3 第四點）。

    回傳一個 list，長度就是 count；還沒出現文字的那幾格是空字串。
    """
    frame = page.frame_locator(CONFIRM_IFRAME_SELECTOR)
    texts = [""] * count
    for _ in range(max(timeout_ms // 200, 1)):
        for i in range(count):
            if texts[i]:
                continue
            try:
                texts[i] = (frame.locator(f"#result{i}").inner_text() or "").strip()
            except PlaywrightError:
                texts[i] = ""
        if all(texts):
            break
        page.wait_for_timeout(200)
    return texts


def cancel_orders(page, session, sheet, ordnos, timeout_ms=20000):
    """
    在一個已登入的分頁上，把 `ordnos` 這幾筆委託刪掉，回傳一列一筆的結果：

        [{"ordno": "J0845", "code": "2007", "side": "買進",
          "ok": True, "message": "刪單成功"}, ...]

    `ordnos` 是委託書號，不是索引——頁面那個 `value` 會隨著新委託整批位移。
    要刪的單如果已經不在頁面上（成交了、剛剛被別人刪了），不算失敗：它會出現在
    回傳的 `missing` 裡，其餘照樣送出。

    **例外的兩種意思差很多**：
      - `RuntimeError`：發生在按下確認**之前**（登入不對、按鈕不見了、核對不過），
        一筆都沒有送出去，重試是安全的。
      - `OrderMaybeSubmitted`：確認**已經按下去了**，那一批多半已經送到券商，
        只是沒等到結果。這種絕對不能自動重來一遍——回報給人，然後重查掛單。
        （沿用 order_fill 那一顆例外，是因為兩邊的意思一模一樣，同一件事不該有
        兩種名字；差別只在這裡是一整批，訊息要講清楚是哪幾筆。）
    """
    wanted = [str(o).strip() for o in ordnos if str(o).strip()]
    if not wanted:
        raise RuntimeError(f"{sheet}：沒有指定要刪哪一筆。")

    rows = _open_page(page)
    by_ordno = {row["ordno"]: row for row in rows}

    targets, missing, locked = [], [], []
    for ordno in wanted:
        row = by_ordno.get(ordno)
        if row is None:
            missing.append(ordno)
        elif not row["cancellable"]:
            # 網站沒畫勾選框就是這一列刪不掉（celable != '1'）——通常是已經成交
            # 或已經刪過了。硬去勾也沒有那個框，據實回報比較有用。
            locked.append(ordno)
        else:
            targets.append(row)

    if not targets:
        return {"results": [], "missing": missing, "locked": locked}

    checked = page.evaluate(CHECK_ROWS_JS, [row["bar"] for row in targets])
    if checked["hit"] != len(targets) or checked["checked"] != len(targets):
        raise RuntimeError(
            f"{sheet}：要勾 {len(targets)} 列，實際勾到 {checked['hit']} 列、"
            f"整頁被勾起來的有 {checked['checked']} 個，先不要送出。")

    # 沒有任何一列刪得掉的時候，頁面會把這顆按鈕整個 remove() 掉，不是變灰
    # （renderTable 最後那段），所以這裡先確認它還在。
    if page.locator("#openConfirm").count() == 0:
        raise RuntimeError(f"{sheet}：頁面上沒有「終止委託單」這顆按鈕，這個帳戶沒有刪得掉的委託。")

    page.click("#openConfirm")
    # 等的是 layer 自己畫的標題列，不是 iframe 裡的 <h3>——理由跟
    # order_fill.fill_order 那段一樣（同樣四個字在兩個地方各出現一次）。
    page.locator(".layui-layer-title", has_text="刪單確認").wait_for(state="visible", timeout=10000)

    expected = [row["ordno"] for row in targets]
    try:
        sent = _verify_dialog(page, expected, sheet, session)
    except RuntimeError:
        close_dialog(page)   # 一筆都沒送，把視窗收乾淨再把例外丟出去
        raise

    frame = page.frame_locator(CONFIRM_IFRAME_SELECTOR)
    frame.locator(SUBMIT_BUTTON).click()
    # ↑ 過了這一行就沒有回頭路了：以下任何失敗都是 OrderMaybeSubmitted。

    texts = _read_results(page, len(sent), timeout_ms)
    close_dialog(page)

    if not any(texts):
        raise OrderMaybeSubmitted(
            f"{sheet}：已經按下刪單確認視窗的「確認」（{len(sent)} 筆：{'、'.join(sent)}），"
            f"但畫面上一格結果都沒出現。這幾筆可能已經送出去了，不要再按一次——"
            f"請重查掛單，用查回來的結果為準。")

    results = []
    for row, ordno, text in zip(targets, sent, texts):
        results.append({
            "ordno": ordno,
            "code": row["code"],
            "side": row["side"],
            # 成功的字是頁面自己寫的「刪單成功」（orderConfirmRWD.html 的 mod=='3'
            # 分支）；失敗直接印券商回的 retmsg，那句話本來就是給人看的。
            "ok": "刪單成功" in text,
            "message": text or "（沒等到結果，請重查掛單確認）",
        })
    return {"results": results, "missing": missing, "locked": locked}


def close_dialog(page):
    """
    關掉刪單確認視窗。送出之後那顆的字會從「取消」變成「關閉」，同一個 `#cancel`。

    關不掉不當成失敗往上丟：這一步跟「單有沒有刪掉」無關，但視窗留著會擋住下一個
    帳戶（整個瀏覽器只有一組 cookie，視窗還開著就換 cookie，那個視窗後來送出去的
    會是新身分——見 10.3 第六點），所以呼叫端要在每一條路上都叫到它。
    """
    try:
        page.frame_locator(CONFIRM_IFRAME_SELECTOR).locator("#cancel").click()
        page.wait_for_timeout(300)
    except PlaywrightError:
        pass


def describe(result):
    """
    一筆結果 -> 一行字。

    刻意不加 ✔／✘ 那種記號：成不成功那句話（「刪單成功」或券商回的失敗原因）
    本來就寫在 message 裡，而那幾個符號在 cp950 編不出來——畫面上沒事，寫進
    crash.log 或 print 出來的時候會炸掉，換來的只是一個記號。
    """
    return f"{result['code']} {result['side']}（{result['ordno']}）　{result['message']}"
