"""
唯讀偵察腳本：驗證「FastQuote 的 WebSocket 是不是真的只能靠 do_login() 那個
page，另開分頁一律拿不到」這個結論，是不是把兩種不同的「另開分頁」搞混了。

fastquote.py 模組說明裡記的那次測試，用的是 `context.new_page()`——一個跟
登入完的分頁完全沒有關聯、獨立生出來的新分頁，結果 sessionStorage 是空的、
被導去首頁。但網站自己在畫面上就有一個連結：

    <a href="javascript:fastQuoteUtil.openWinURL('../FastQuote/index.jsp')">簡易看盤下單</a>

內部是 `window.open()`。瀏覽器對「同一個分頁自己呼叫 window.open() 開出來的
新視窗」有專屬規則：新視窗會複製一份 opener 那個分頁當下的 sessionStorage
（跟瀏覽器「複製分頁」sessionStorage 會延續是同一條規則），跟完全獨立、沒有
opener 關係的新分頁不是同一回事——login.py 的 wait_until_finished() 甚至已經
在處理「使用者手動點這個連結」的案例（跳出視窗卡在 about:blank，要靠持續呼叫
Playwright 才會載入），side effect 是那份說明間接證實這個彈出視窗本來就是網站
設計上會被人拿來用的東西，不是我們自己編出來的路徑。

這支腳本就是要驗證這個猜測：用登入完的那個 page 自己呼叫
`fastQuoteUtil.openWinURL(...)`，攔截跳出來的視窗，檢查它的 sessionStorage
是不是真的有 branch_id/cust_id、URL 是不是真的停在 FastQuote，以及掛上去的
WebSocket 監聽器收不收得到報價封包——如果這條路通，之後要接即時委買賣一就
不必再去煩惱「跟哪一組帳號的 page 搶」，因為這個彈出視窗本來就是另一個獨立
分頁，登入完的那個 page 完全不受影響，可以繼續做別的查詢。

安全設計：全程只做「登入 → 用頁面自己的 JS 開出這個彈出視窗 → 讀
sessionStorage → 掛被動 WebSocket 監聽器等幾秒 → 記錄」，不點擊/填寫/送出
任何下單相關的元素（跟 recon_fastquote.py 同一種態度）。
"""

import sys
import time
import traceback
from datetime import datetime

from playwright.sync_api import Error as PlaywrightError, TimeoutError as PlaywrightTimeoutError
from playwright.sync_api import sync_playwright

from fastquote import _PATCH_WEBSOCKET_JS, _decode_records
from login import (
    app_dir,
    configure_browsers_path,
    do_login,
    load_accounts,
    open_context,
    pause,
    wait_until_finished,
)

OUTPUT_DIR_NAME = "偵察資料"
WATCH_SECONDS = 15  # 掛上監聽器之後被動等幾秒，看有沒有收到報價封包

SESSION_JS = """
() => ({
    branch_id: sessionStorage.getItem('branch_id'),
    cust_id: sessionStorage.getItem('cust_id'),
    keys: Object.keys(sessionStorage),
})
"""

OPEN_POPUP_JS = "() => { fastQuoteUtil.openWinURL('../FastQuote/index.jsp'); }"


def main():
    accounts = load_accounts()
    if not accounts:
        print(f"找不到帳號設定。請在 {app_dir()} 放一個 .env 檔（可複製 .env.example）。")
        sys.exit(1)

    which = 1
    if len(sys.argv) >= 2:
        try:
            which = int(sys.argv[1])
        except ValueError:
            print(f"參數要是數字（第幾組帳號），收到的是: {sys.argv[1]}")
            sys.exit(1)
    if not 1 <= which <= len(accounts):
        print(f".env 裡目前有 {len(accounts)} 組帳號，第 {which} 組不存在。")
        sys.exit(1)
    account = accounts[which - 1]

    out_dir = app_dir() / OUTPUT_DIR_NAME
    out_dir.mkdir(exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M")

    configure_browsers_path()

    report = [
        f"偵察時間: {datetime.now():%Y/%m/%d %H:%M:%S}",
        f"第 {which} 組帳號",
        "測試：從登入完的 page 用 fastQuoteUtil.openWinURL() 開彈出視窗，",
        "看 sessionStorage 跟 WebSocket 收不收得到報價（本次不點擊/填寫任何下單元素）。",
    ]

    with sync_playwright() as p:
        context, browser = open_context(p)
        # 要在任何分頁（含之後跳出來的彈出視窗）載入前就патch好，跟 fastquote.py
        # 的 FastQuoteStream 同一招，只是這裡用 context 層級，讓等一下跳出來的
        # popup 也吃得到這段 patch（popup 不是我們自己 page.goto 開的，沒辦法
        # 用 page.add_init_script 專門對它下）。
        context.add_init_script(_PATCH_WEBSOCKET_JS)

        spare_page = context.pages[0] if context.pages else None

        try:
            page = do_login(context, account["id"], account["password"], spare_page)
        except PlaywrightTimeoutError:
            report.append("登入逾時，找不到欄位，網站版面可能已變更。")
            print("\n".join(report))
            sys.exit(1)
        except PlaywrightError as exc:
            report.append(f"瀏覽器操作失敗：{exc}")
            print("\n".join(report))
            sys.exit(1)

        report.append(f"登入完成，目前頁面: {page.url}")
        print("\n".join(report))

        try:
            with context.expect_page(timeout=10000) as popup_info:
                page.evaluate(OPEN_POPUP_JS)
            popup = popup_info.value
        except PlaywrightError as exc:
            report.append(f"沒有攔截到彈出視窗（fastQuoteUtil 可能不存在於這頁，或被瀏覽器擋掉彈窗）：{exc}")
            print(report[-1])
            wait_until_finished(context)
            sys.exit(1)

        popup.wait_for_load_state("domcontentloaded", timeout=10000)
        report.append(f"彈出視窗網址: {popup.url}")

        session = popup.evaluate(SESSION_JS)
        report.append(f"彈出視窗 sessionStorage: branch_id={session.get('branch_id')!r} "
                       f"cust_id={session.get('cust_id')!r}")
        report.append(f"彈出視窗 sessionStorage 的 key: {session.get('keys')}")

        identity_ok = bool(session.get("branch_id") and session.get("cust_id"))
        report.append("身分結論: " + ("有效（跟登入時同一份 sessionStorage）" if identity_ok
                                     else "空的（這條路一樣行不通，跟 context.new_page() 一樣的結果）"))

        quotes_seen = {}

        def on_websocket(ws):
            if "push.tbbstock.com.tw" not in ws.url:
                return
            report.append(f"偵測到 WebSocket 連線: {ws.url}")

            def on_frame(payload):
                if isinstance(payload, str):
                    return
                for record in _decode_records(bytes(payload)):
                    if record[0] == "quote":
                        _, internal_id, bid, ask, last = record
                        quotes_seen[internal_id] = (bid, ask, last)

            ws.on("framereceived", on_frame)

        popup.on("websocket", on_websocket)

        print(f"{page.url} 登入分頁沒有被動任何操作；彈出視窗被動等 {WATCH_SECONDS} 秒看有沒有報價...")
        popup.wait_for_timeout(WATCH_SECONDS * 1000)

        if quotes_seen:
            report.append(f"收到 {len(quotes_seen)} 檔報價（內部代碼 -> (委買一,委賣一,成交價)）: {quotes_seen}")
            report.append("結論: 彈出視窗這條路可行，WebSocket 收得到報價，且完全沒動到登入分頁。")
        else:
            report.append("沒有收到任何報價封包（可能是收盤後自選清單沒有預設報價，不代表連線本身不行）。")

        report_path = out_dir / f"{stamp}_fastquote_popup_摘要.txt"
        report_path.write_text("\n".join(report), encoding="utf-8")

        print("\n".join(report[4:]))
        print(f"摘要存在: {report_path}")

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
