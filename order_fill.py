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
    python order_fill.py <第幾組帳號> <股票代號> <張數> <價格> [bs_flag]

    bs_flag 預設 I（IOC，立即成交否則取消，安全，不會掛著等成交）。
    真正盤前要用的 R（ROD，當日有效）會真的掛在那邊等撮合，測試時要留意。

例：
    python order_fill.py 1 1714 1 14.9 I

下一步：接到「下單」分頁的執行預覽，讓多帳戶依序跑（一個一個開委託確認
視窗，還是人自己按確認），還沒做，也還沒驗證過「換帳戶 cookie 不重登」
這招用在下單頁面上一樣行得通（目前只驗證過查詢，見 order_recon.py）。
"""

import sys

from playwright.sync_api import Error as PlaywrightError, TimeoutError as PlaywrightTimeoutError
from playwright.sync_api import sync_playwright

from login import app_dir, configure_browsers_path, do_login, load_accounts, open_context, pause

ORDER_ENTRY_PAGE = "https://www.tbbstock.com.tw/tbb/order/layoutRWD.jsp?type=0"

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


def open_order_form(page):
    """
    交易盤別（整股）／交易別（現股）先設好，一定要排在選股票之前。

    2026/08/28 實測發現：頁面自己有一段 `$('.tab1').on('change', function(){
    $('#stockId').val('').trigger('change'); ... })`，只要 .tab1 這個下拉選單
    的 change 事件被觸發，就會把已經選好的股票清空——原本 fill_order()
    把這兩個下拉選單的設定放在 select_stock() 之後，結果就是股票選好了、
    緊接著被這個 handler 清掉，等按「確認下單」的時候 #stkName 又變回
    "--"，跳出「請輸入正確股號」，跟股票選錯是同一種症狀但成因完全不同。
    """
    page.select_option(".tab1", "1")        # 整股
    page.select_option("#tradeType", "0")   # 現股


def fill_order(page, *, side, qty, price, bs_flag="I"):
    """
    填好整股限價單的其餘欄位、按「確認下單」開出委託確認視窗，不按裡面的
    「確認」——那一步留給人。side 是 'B' 或 'S'，對到 #orderB / #orderS。

    呼叫這支之前一定要先 open_order_form() 再 select_stock()，順序反了
    股票會被清空（見 open_order_form 的說明）。
    """
    page.check(f"#order{side}")             # 買進/賣出
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


def main():
    if len(sys.argv) < 5:
        print(__doc__)
        sys.exit(1)

    which = int(sys.argv[1])
    code, qty, price = sys.argv[2], sys.argv[3], sys.argv[4]
    bs_flag = sys.argv[5] if len(sys.argv) > 5 else "I"

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
            open_order_form(page)
            select_stock(page, code)
            fill_order(page, side="S", qty=qty, price=price, bs_flag=bs_flag)
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
