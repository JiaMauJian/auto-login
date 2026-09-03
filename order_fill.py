"""
半自動下單：登入後操作真實的「單筆委託」表單，填好股票／買賣別／數量／
價格／委託別，按「確認下單」開出委託確認視窗——**到這裡就停下來**，不會
按視窗上的「確認」，那一步留給人自己決定要不要真的送出。2026/08/28 已經
實測跑通全程，委託確認視窗上的股票／買賣別／張數／價格／委託條件／
預估價金都正確帶出來過。

跟 order_recon.py 不一樣：那支完全唯讀；這支會操作真實表單（雖然還沒有
真的送出委託）。新股票／新帳號第一次測試還是建議用小單位、不容易成交的
價格（例如原本用過的：跌停價買進、IOC，或反過來賣出用漲停價），降低萬一
按下確認真的送出去的影響。

股票代號的選取沒有走畫面上的下拉選單（select2 的視覺互動要點開、打字、
等 AJAX、點結果，這幾步在瀏覽器自動化裡容易因為時間點沒抓好而失敗）。
改成直接呼叫 select2 本來就會打的同一支 AJAX（Select2Servlet），拿到結果
後用 jQuery 組一個 Option 塞進 select2、觸發它自己的 select2:select 事件
（順序要排在原生 change 之前，見 SELECT_STOCK_JS 的說明）——這是頁面
本來就有的機制，只是繞過視覺互動，跟 recon.py 呼叫 B64_XOR_Encode 那支
現成函式同一個做法（見 CLAUDE.md、記憶「下單 API 細節」）。實際踩過的坑
（select2 內部資料不同步、Select2Servlet 的 id 不是股票代號、.tab1 一動
就清空股票、確認視窗標題要去哪裡找）都寫在對應函式的說明跟記憶
「order_fill.py 半自動下單已跑通」裡，之後改版遇到類似症狀先查那邊。

用法：
    python order_fill.py <第幾組帳號> <股票代號> <數量> <價格> [bs_flag] [--odd]

    數量的單位跟著 --odd 走：沒帶就是整張、數量填「張」；帶了 --odd 是零股、
    數量填「股」（1~999）。--odd 放在哪個位置都可以。
    bs_flag 預設 I（IOC，立即成交否則取消，安全，不會掛著等成交）。
    真正盤前要用的 R（ROD，當日有效）會真的掛在那邊等撮合，測試時要留意。

例：
    python order_fill.py 1 1714 1 14.9 I
    python order_fill.py 1 1714 350 14.9 R --odd

下一步：接到「下單」分頁的執行預覽，讓多帳戶依序跑（一個一個開委託確認
視窗，還是人自己按確認），還沒做，也還沒驗證過「換帳戶 cookie 不重登」
這招用在下單頁面上一樣行得通（目前只驗證過查詢，見 order_recon.py）。
"""

import sys

from playwright.sync_api import Error as PlaywrightError, TimeoutError as PlaywrightTimeoutError
from playwright.sync_api import sync_playwright

from login import app_dir, configure_browsers_path, do_login, load_accounts, open_context, pause
from orders import SHARES_PER_LOT

ORDER_ENTRY_PAGE = "https://www.tbbstock.com.tw/tbb/order/layoutRWD.jsp?type=0"

# 交易盤別（.tab1）。整股的 "1" 是 2026/08/28 實測跑通的；零股的 "5" 是
# 2026/09/01 09:18 盤中跑 recon_order_form.py 倒出來的（報告在
# 偵察資料60901_0918_下單表單選項.txt）。這個下拉選單一共四個選項：
#
#     '1' = 整股      '2' = 盤後      '5' = 盤中零股      '3' = 盤後零股
#
# 零股取 '5'（盤中零股）不是 '3'（盤後零股）：'5' 對整股的 '1'，是同一段連續
# 交易時間裡的另一半，規劃文件的零股流程（買賣股票照試算送、出清零股「全部掛
# 賣單、20 秒後取消」）講的都是這一段。'3' 是 13:40~14:30 那場盤後集合競價，
# 對到的是 '2'（盤後），整張那半邊也沒在用——真的要做盤後那一場的話，是多一個
# 「時機」設定，不是把這個值改掉。
#
# 同一份報告確認的另外兩件（都跟程式原本的假設一致，所以沒有東西要改）：
#   - 零股的數量欄旁邊寫的是「股」（整股寫「張」），見 _check_qty
#   - 零股的委託別只剩 'R'（ROD），交易別只剩 '0'（現股）——整股才有 IOC/FOK
TAB1_LOT = "1"
TAB1_ODD = "5"


class OrderMaybeSubmitted(RuntimeError):
    """
    「確認」已經按下去了，但沒辦法確定委託實際上有沒有送出去、送出去的結果
    是什麼——跟一般 RuntimeError（送出前就確定失敗，重試安全）不一樣，呼叫端
    看到這個例外類型絕對不能自動當成「按下一筆重試同一筆」的訊號，得先讓人
    自己去網站查證，見 confirm_order 的說明。
    """

# 直接呼叫 select2 本來會打的那支 AJAX，繞過視覺互動（見檔案開頭說明）。
# type=public 是從頁面的 select2() 設定原樣抄過來的。
LOOKUP_STOCK_JS = """
(code) => new Promise((resolve) => {
    $.ajax({
        url: '../Select2Servlet',
        data: { search: code, type: 'public' },
        success: (data) => resolve({ok: true, data}),
        error: (jq, status) => resolve({ok: false, status, text: jq.responseText}),
    });
})
"""

# 把查到的結果塞進 select2，觸發它原本 change 時會做的那一串（帶出股名、
# 平盤/漲跌停價）。option.id/option.text 是 Select2Servlet 回應本來就有的
# 欄位名稱（select2 的 processResults 直接把原始回應當結果用，見
# order/layoutRWD.jsp?type=0 的 select2_input 設定）。
#
# 2026/08/28 實測分兩輪才抓到正確順序：
#   第一輪只 append + trigger('change')：底層 <select> 的 value 設對了，
#     但 select2 自己畫出來的框還是顯示 placeholder——select2 另外維護
#     一份「目前選了什麼」的內部資料，只跟著它自己的 select2:select
#     事件走，不會因為原生 change 事件就重畫。
#   第二輪把 select2:select 補在 change 後面：框上的字對了，但頁面本身
#     「查股票資訊」那段（填 #stkName、平盤/漲跌停價）完全沒有跑——
#     因為頁面自己的 change handler 裡會呼叫 $(this).select2('data')，
#     這段是在 change 事件當下就執行的，那時候 select2 的內部資料還沒
#     被 select2:select 同步過，讀到的是空的，handler 中途出錯就整段
#     沒繼續往下跑。
# 正確順序是「先讓 select2 自己的資料同步，change 事件觸發時頁面才讀得到
# 正確的 select2('data')」，所以 select2:select 要排在 change 前面。
SELECT_STOCK_JS = """
(item) => {
    const option = new Option(item.text, item.id, true, true);
    $('#stockId').append(option);
    $('#stockId').trigger({ type: 'select2:select', params: { data: item } });
    $('#stockId').trigger('change');
}
"""


def lookup_stock(page, code):
    """查一次 Select2Servlet，回傳原始回應（不管成不成功都印出來方便對照）。"""
    result = page.evaluate(LOOKUP_STOCK_JS, code)
    if not result.get("ok"):
        raise RuntimeError(f"Select2Servlet 查詢失敗：{result}")
    return result["data"]


def select_stock(page, code):
    """
    選股票。查到多筆的話，挑「代號」完全等於 code 的那一筆。

    2026/08/28 實測發現 Select2Servlet 回應的 id 不是股票代號，是內部的
    數字 ID（例如 1714 和桐的 id 是 1360），跟頁面本身 change handler
    判斷股號的方式一樣——不能比對 id，要從 text 裡「代號 名稱」的第一段
    取代號來比對，比對 id 永遠對不上，會一路掉到「用第一筆」那個備援。
    """
    matches = lookup_stock(page, code)
    print(f"[select_stock] Select2Servlet 回應：{matches}")
    if not matches:
        raise RuntimeError(f"查無股票代號 {code}，Select2Servlet 沒有回傳任何結果。")

    exact = [m for m in matches
             if str(m.get("text", "")).strip().split(" ")[0].upper() == code.upper()]
    item = exact[0] if exact else matches[0]
    if not exact:
        print(f"[select_stock] 沒有完全對上代號的結果，改用第一筆：{item}")

    page.evaluate(SELECT_STOCK_JS, item)
    # #stkName 從 "--" 變成真正的股名，代表 change handler 真的跑完了
    # （包含它自己那支查股價的 AJAX），不能只看 select2 觸發完就当作選好了。
    page.wait_for_function(
        "() => document.getElementById('stkName') && document.getElementById('stkName').textContent.trim() !== '--'",
        timeout=10000,
    )
    stk_name = page.locator("#stkName").inner_text().strip()

    # 「確認下單」讀的是 $("#stockId").select2('data')，不是單純的 DOM
    # value——這裡直接照它的讀法驗一次，比只看 #stkName 更接近真正會
    # 被拿去用的那份資料，選錯/沒同步在這裡就會現形，不用等按下去才發現。
    select2_data = page.evaluate("() => $('#stockId').select2('data')")
    print(f"[select_stock] select2('data') = {select2_data}")
    if not select2_data or str(select2_data[0].get("text", "")).split(" ")[0].upper() != code.upper():
        raise RuntimeError(
            f"select2 內部資料跟預期的代號 {code} 對不起來：{select2_data}，"
            f"「確認下單」按下去讀到的會是錯的股票，先不要繼續。")

    print(f"[select_stock] 已選定：{item.get('id')} {stk_name}")
    return stk_name


def open_order_form(page, *, odd=False):
    """
    交易盤別（整股／零股）／交易別（現股）先設好，一定要排在選股票之前。

    2026/08/28 實測發現：頁面自己有一段 `$('.tab1').on('change', function(){
    $('#stockId').val('').trigger('change'); ... })`，只要 .tab1 這個下拉選單
    的 change 事件被觸發，就會把已經選好的股票清空——原本 fill_order()
    把這兩個下拉選單的設定放在 select_stock() 之後，結果就是股票選好了、
    緊接著被這個 handler 清掉，等按「確認下單」的時候 #stkName 又變回
    "--"，跳出「請輸入正確股號」，跟股票選錯是同一種症狀但成因完全不同。

    odd=True 是零股那一段（同一個試算數字的另一半，見 orders.split_lots）。
    """
    tab1 = TAB1_ODD if odd else TAB1_LOT
    page.select_option(".tab1", tab1)

    # 選完讀回來核對一次。select_option 選一個「不存在」的值會自己丟例外，但
    # 「選項在、這個時段不能選」會怎樣沒人試過——2026/09/01 那份報告是 09:18
    # 盤中倒的，四個盤別當時都切得過去，**盤前（09:00 以前）切不切得過去還
    # 沒人看過**。真的沒切過去的話下面填的數量會用整股的單位送出去，多問這一
    # 句就擋掉了，成本是一次 evaluate。
    actual = page.eval_on_selector(".tab1", "el => el.value")
    if actual != tab1:
        raise RuntimeError(
            f"交易盤別設成 {tab1!r} 沒有生效，頁面現在停在 {actual!r}"
            f"（{'零股' if odd else '整股'}這個選項這個時段可能不能選）。這一筆不送。")

    page.select_option("#tradeType", "0")   # 現股


def _check_qty(qty, *, odd):
    """
    數量欄填錯單位不會報錯，只會送出差 1000 倍的委託，所以填之前先量一下。

    整張填「張」、零股填「股」（不到一張的量本來就寫不成張），兩者共用
    plan_* 那一列的 "lots" 欄位（見 orders.plan_trade_orders）——正因為同一個
    欄位名裝著兩種單位，把「這一列的 unit 跟裡面的數字對不對得起來」在送出去
    之前問一次才有意義。

    B1 如果查出零股的數量欄其實不收「股」，要改的是這個檢查跟呼叫端算出來的
    數字，不是在這裡偷偷除以 1000。
    """
    try:
        value = int(str(qty).strip())
    except ValueError:
        raise RuntimeError(f"數量看不懂：{qty!r}，這一筆不送。") from None
    if value <= 0:
        raise RuntimeError(f"數量要是正整數，收到 {value}，這一筆不送。")
    if odd and value >= SHARES_PER_LOT:
        raise RuntimeError(
            f"零股的數量要是 1~{SHARES_PER_LOT - 1} 股，收到 {value}——這個數字"
            f"看起來是「張」不是「股」（差 {SHARES_PER_LOT} 倍），這一筆不送。")
    return value


def _check_bs_radio(page, side, timeout_ms=8000):
    """
    選買進/賣出。2026/09/03 多輪測試撞過 page.check() 卡滿 30 秒逾時，
    debug log 顯示先是「element is not stable」（版面還在位移），最後變成
    「element was detached from the DOM」——推測是選完股票後頁面自己還有
    別的 AJAX 在重畫這附近的區塊（跟 select_stock() 裡 select2 資料同步
    要等一輪的成因類似），但這裡抓不到一個像 #stkName 那樣明確的「查完
    了」訊號可以等。做法比照 select2：優先讓 Playwright 走一般的 click
    （讓頁面自己可能綁的 click handler 正常跑一次），真的卡住逾時才退回
    直接用 JS 設值＋補發事件，不用管視覺穩不穩定、也不管節點被整個換掉
    幾次，因為 JS 是重新查一次目前的 DOM，不是抱著舊的節點參考硬打。
    """
    try:
        page.check(f"#order{side}", timeout=timeout_ms)
        return
    except PlaywrightTimeoutError:
        print(f"[fill_order] #order{side} 一般 click 逾時（頁面可能還在重畫），改用 JS 直接設值。")

    page.evaluate(
        """(side) => {
            const el = document.getElementById('order' + side);
            if (!el) throw new Error('#order' + side + ' 不存在');
            el.checked = true;
            el.dispatchEvent(new Event('input', {bubbles: true}));
            el.dispatchEvent(new Event('change', {bubbles: true}));
            el.dispatchEvent(new Event('click', {bubbles: true}));
        }""",
        side,
    )


def fill_order(page, *, side, qty, price, bs_flag="I", odd=False):
    """
    填好限價單的其餘欄位、按「確認下單」開出委託確認視窗，不按裡面的
    「確認」——那一步留給人。side 是 'B' 或 'S'，對到 #orderB / #orderS。

    qty 的單位跟著 odd 走：整張填「張」、零股填「股」（2026/09/01 的偵察報告
    確認過，見 TAB1_ODD 那段）。
    odd 只影響這裡的範圍檢查，真正切盤別的是 open_order_form(odd=...)
    ——兩支要帶同一個值，只帶一邊會變成「用整股的盤別送零股的數量」。

    呼叫這支之前一定要先 open_order_form() 再 select_stock()，順序反了
    股票會被清空（見 open_order_form 的說明）。
    """
    qty = _check_qty(qty, odd=odd)
    if odd and bs_flag != "R":
        # 零股的委託別只剩 ROD（2026/09/01 偵察報告）。硬選下去 select_option
        # 會丟「找不到這個選項」，訊息看不出真正的原因，先在這裡講清楚。
        raise RuntimeError(
            f"零股的委託別只有 R（ROD-當日有效），收到 {bs_flag!r}——IOC／FOK 是"
            f"整股才有的。這一筆不送。")

    _check_bs_radio(page, side)             # 買進/賣出
    page.fill("#qty", str(qty))
    page.select_option("#priceRadio", "0")  # 限價
    page.fill("#price", str(price))
    page.select_option("#bsFlag", bs_flag)

    page.click("#openConfirm1")
    # 2026/08/28 實測分兩輪才找對地方：委託確認視窗裡「委託確認」這四個字
    # 其實出現兩處——layer.js 自己畫的標題列（在主頁面 DOM 裡，class 是
    # .layui-layer-title）跟 iframe 內容裡 orderConfirmRWD.html 自己的
    # <h3>（載入 orderConfirmRWD.html，見偵察資料\委託確認視窗\layer.js）。
    # 第一輪檢查 iframe 裡那個一直逾時（畫面明明已經正確跳出來），改成檢查
    # 主頁面上 layer.js 自己的標題列，不用猜 iframe 裡的精確文字/選擇器。
    page.locator(".layui-layer-title", has_text="委託確認").wait_for(
        state="visible", timeout=10000)
    print("[fill_order] 委託確認視窗已開啟，停在這裡——請自己到瀏覽器裡看內容，"
          "確認沒問題再按「確認」送出，或按「取消」放棄，程式不會替你按。")


# 委託確認視窗裡「確認」按鈕跟送出結果的文字都在 layer.js 開出來的 iframe
# 裡（orderConfirmRWD.html，見偵察資料\委託確認視窗\），不是主頁面 DOM——
# 跟上面 fill_order 找標題列要避開 iframe 是同一件事的另一面：這次要按的
# 按鈕就在 iframe 裡，躲不掉，只能進去找。iframe 的 id/name 是動態編號
# （layui-layer-iframe1、2...，看這個瀏覽器 session 已經開過幾個 layer），
# 沒辦法寫死，用 src 裡有 orderConfirmRWD 這個特徵去找。
CONFIRM_IFRAME_SELECTOR = "iframe[src*='orderConfirmRWD']"


def confirm_order(page, timeout_ms=15000):
    """
    按下委託確認視窗裡的「確認」，真的把委託送出去——**過這一步沒有回頭路**。
    只有 GUI 的「自動送出」開關打開時才會呼叫這裡；order_fill.py 本身的
    半自動流程（main()／fill_order）到開出確認視窗就停手，不會呼叫這支。

    這支函式的選擇器（CONFIRM_IFRAME_SELECTOR、#submit、#result0）是照
    orderConfirmRWD.html 的原始碼推出來的，還沒有實際跑過驗證——2026/08/28
    只驗證到「填單、開出確認視窗、人自己按確認」這一步（見記憶
    order-exec-sequential-wired-up），這支是新加的，第一次用請找一個影響最小
    的情境測（小單位、不容易成交的價格，或收盤後可以事後刪單的時段）。

    回傳 (ok, message)：message 是 #result0 顯示的原文（成功會是「委託成功,
    委託書編號: ...」；IOC 沒吃到價會是券商回的錯誤訊息，例如「IOC. FOK
    委託未能成交，委託失敗」——這是市場真的沒成交，不是這支函式或這一步
    壞掉，屬於正常結果，不是例外）。ok 只是簡單判斷 message 裡有沒有
    「委託成功」四個字，給呼叫端畫面用。

    **最關鍵的安全設計**：`#submit` 一旦真的被點下去，委託多半已經送出去了，
    之後不管是等結果逾時還是收視窗失敗，都不能被呼叫端當成「這一筆還沒送出、
    按下一筆可以重試」——重試等於把同一筆委託多送一次。所以點擊之後任何
    失敗都在這裡包成一句話講清楚「已經按下確認，請自己去網站查證，不要再
    按下一筆重試」，跟「送出前就失敗、按下一筆重試是安全的」是兩種完全
    不同的失敗，呼叫端分不出來的話，會有真的把同一筆送兩次的風險。
    """
    frame = page.frame_locator(CONFIRM_IFRAME_SELECTOR)
    frame.locator("#submit").click()

    result = frame.locator("#result0")
    message = ""
    for _ in range(timeout_ms // 200):
        try:
            message = (result.inner_text() or "").strip()
        except PlaywrightError:
            message = ""
        if message:
            break
        page.wait_for_timeout(200)

    if not message:
        raise OrderMaybeSubmitted(
            "已經按下「確認」，委託應該已經送出去了，但畫面上沒有等到結果文字"
            "（可能只是反應比較慢）——請自己到「委託查詢」或「預約查詢」頁確認"
            "這筆到底送出去了沒，不要直接當作失敗按「下一筆」重試，那樣可能會"
            "把同一筆委託送兩次。")

    try:
        frame.locator("#cancel").click()   # 送出成功後這顆文字變成「關閉」，見 orderConfirmRWD.html
    except PlaywrightError:
        # 結果已經拿到了，這顆按不按得下去不影響委託本身有沒有送出去；但視窗
        # 沒關掉的話，_order_dialog_closed 那道輪詢會一直看到它還開著，「下
        # 一筆」不會自動解鎖，使用者要自己去瀏覽器把這個視窗關掉才能繼續。
        pass

    return "委託成功" in message, message


def main():
    # --odd 可以放在任何位置（它不是位置參數），先濾掉再照位置讀其餘的。
    argv = [arg for arg in sys.argv[1:] if arg != "--odd"]
    odd = "--odd" in sys.argv
    if len(argv) < 4:
        print(__doc__)
        sys.exit(1)

    which = int(argv[0])
    code, qty, price = argv[1], argv[2], argv[3]
    bs_flag = argv[4] if len(argv) > 4 else "I"

    accounts = load_accounts()
    if not accounts:
        print(f"找不到帳號設定。請在 {app_dir()} 放一個 .env 檔。")
        sys.exit(1)
    if not 1 <= which <= len(accounts):
        print(f".env 裡目前有 {len(accounts)} 組帳號，第 {which} 組不存在。")
        sys.exit(1)

    account = accounts[which - 1]
    configure_browsers_path()

    with sync_playwright() as p:
        context, browser = open_context(p)
        spare_page = context.pages[0] if context.pages else None

        try:
            page = do_login(context, account["id"], account["password"], spare_page)
            page.goto(ORDER_ENTRY_PAGE)
            open_order_form(page, odd=odd)
            select_stock(page, code)
            fill_order(page, side="S", qty=qty, price=price, bs_flag=bs_flag, odd=odd)
        except PlaywrightTimeoutError as exc:
            print(f"逾時：{exc}")
        except PlaywrightError as exc:
            print(f"瀏覽器操作失敗：{exc}")
        except Exception as exc:
            # select_stock 裡的安全檢查（select2 資料對不起來）丟的是
            # RuntimeError，不是上面兩種 Playwright 例外——沒有這一道的話，
            # 例外會直接穿出 try，連下面「瀏覽器留著不關」都不會印，
            # with sync_playwright() 收尾時瀏覽器可能跟著關掉，等於失敗的
            # 那一次反而看不到現場，比正常結束更不該發生。
            print(f"發生錯誤：{exc}")

        print("瀏覽器留著不關，自己去看畫面、決定要不要送出。")
        pause("看完按 Enter 結束這支程式（不會關瀏覽器）...")


if __name__ == "__main__":
    main()
