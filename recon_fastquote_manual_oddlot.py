"""
唯讀偵察腳本：這次不是程式自己送 WebSocket 指令去測假說，是把操作權交還給
使用者——開好 FastQuote 彈出視窗之後，程式只掛被動的 WebSocket／網路／畫面
監聽器，完全不碰任何畫面元素，讓使用者自己在視窗上輸入股號、按「查詢」／
「零股查詢」、切換「整股／盤後／盤中零股／盤後零股」分頁，程式全程被動記錄
這段期間收到的每一筆 WebSocket frame、每一筆 XHR/fetch，外加畫面文字內容跟
所有 input 欄位值有變動就拍一張快照，當「螢幕上實際看到的數字」的對照組。

## 為什麼換這個方向

`recon_fastquote_channel57.py`（channel 57，零股查詢按鈕本來會送的 WS 指令）
已經測完，10 輪 5 分鐘數值不動，是死路。`recon_fastquote_oddlot_subscribe.py`
（channel 42 訂閱時代號加 `.O`）是另一個假說，還在測。這支腳本不猜測是哪個
channel／哪種格式，直接觀察真人在畫面上點「零股查詢」或切到「盤中零股」分頁
查詢的當下，程式被動記下所有東西——不管零股即時委買賣一實際上是走 WebSocket
的哪個 channel/field，還是其實是一支我們還沒注意到的 HTTP AJAX，這支腳本都
不預設答案，全部收下來，事後再比對哪一筆跟畫面上看到的數字對得上。

## 怎麼操作

跑起來、登入完、彈出 FastQuote 視窗之後，程式會印出提示然後停在原地等——
跟 login.py 的 wait_until_finished() 同一個機制：按 Enter 或直接關掉那個
彈出視窗才會結束記錄。這段期間請直接在彈出的 FastQuote 視窗上操作：在
「股號」輸入框打代號、按「查詢」，再按「零股查詢」，也可以切到「盤中零股」
分頁看即時五檔，想測幾次都可以，程式全程只記錄不介入。操作完按 Enter 結束，
程式會把這段期間的記錄整理成報告跟原始 json。

## 安全設計

全程不呼叫任何 page.evaluate 之外的操作，也不 patch WebSocket、不呼叫
`.send()`——只掛 `ws.on("framereceived")`、`page.on("request")`、
`page.on("response")` 這幾種被動監聽，加上唯讀的 `page.evaluate(...)`
（讀 `document.body.innerText` 跟所有 input 的 value，不寫入任何值）。
操作畫面完全交給使用者，程式不點擊/填寫/送出任何元素。
"""

import json
import sys
import time
import traceback
from datetime import datetime

from playwright.sync_api import Error as PlaywrightError, TimeoutError as PlaywrightTimeoutError
from playwright.sync_api import sync_playwright

from fastquote import _decode_records
from login import (
    _enter_pressed,
    app_dir,
    configure_browsers_path,
    do_login,
    load_accounts,
    open_context,
    pause,
)

OUTPUT_DIR_NAME = "偵察資料"
OPEN_POPUP_JS = "() => { fastQuoteUtil.openWinURL('../FastQuote/index.jsp'); }"
SNAPSHOT_JS = """
() => ({
    text: document.body.innerText,
    inputs: Array.from(document.querySelectorAll('input')).map(el => ({
        id: el.id, name: el.name, value: el.value
    })),
})
"""
SNAPSHOT_INTERVAL_S = 1.0  # 每隔這麼久檢查一次畫面有沒有變化（有變才記錄）


def _watch_and_wait(popup, context, get_snapshot):
    """
    跟 login.wait_until_finished() 同一個機制（按 Enter 或視窗關掉才結束、
    用 wait_for_timeout 分段等讓 Playwright 有機會處理事件），多一件事：
    每隔 SNAPSHOT_INTERVAL_S 秒檢查一次畫面內容，有變動才記一筆快照——
    不是每次都記，操作之間的空檔沒必要塞一堆一模一樣的快照。
    """
    print("開始被動記錄，可以在彈出視窗上操作了：輸入股號、按查詢／零股查詢、"
          "切換分頁都可以，想測幾次都行。", flush=True)
    print("操作完按 Enter 結束記錄（也可以直接關掉那個彈出視窗）...", flush=True)

    last_snapshot_text = None
    last_check = 0.0

    while True:
        pressed = _enter_pressed()
        if pressed is None:
            return  # 讀不到鍵盤（例如被別的程式呼叫），維持原本不等待的行為
        if pressed:
            return

        try:
            if popup.is_closed():
                print("彈出視窗已關閉。")
                return
        except PlaywrightError:
            print("彈出視窗已關閉。")
            return

        now = time.monotonic()
        if now - last_check >= SNAPSHOT_INTERVAL_S:
            last_check = now
            try:
                snap = popup.evaluate(SNAPSHOT_JS)
            except PlaywrightError:
                snap = None
            if snap is not None:
                key = json.dumps(snap, ensure_ascii=False, sort_keys=True)
                if key != last_snapshot_text:
                    last_snapshot_text = key
                    get_snapshot(snap)

        try:
            popup.wait_for_timeout(150)
        except PlaywrightError:
            print("彈出視窗已關閉。")
            return


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
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")

    configure_browsers_path()

    report = [
        f"偵察時間: {datetime.now():%Y/%m/%d %H:%M:%S}",
        f"第 {which} 組帳號",
        "手動操作側錄：使用者自己在 FastQuote 彈出視窗輸入股號／按查詢／"
        "按零股查詢／切分頁，程式只被動記錄 WebSocket + XHR/fetch + 畫面快照。",
    ]

    ws_log = []       # WebSocket frame（含已知欄位解出來的、跟完全沒解出來的）
    net_log = []      # XHR/fetch 的 request/response
    snapshot_log = []  # 畫面 innerText + input value，有變動才記一筆

    with sync_playwright() as p:
        context, browser = open_context(p)
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
            report.append(f"沒有攔截到彈出視窗：{exc}")
            print(report[-1])
            sys.exit(1)

        popup.wait_for_load_state("domcontentloaded", timeout=10000)
        report.append(f"彈出視窗網址: {popup.url}")
        print(report[-1])

        record_start = time.monotonic()

        def elapsed():
            return round(time.monotonic() - record_start, 2)

        def on_websocket(ws):
            if "push.tbbstock.com.tw" not in ws.url:
                return
            ws_log.append({"t": elapsed(), "type": "ws_open", "url": ws.url})
            print(f"[{elapsed():7.2f}s] WebSocket 連線開啟: {ws.url}")

            def on_frame(payload):
                t = elapsed()
                if isinstance(payload, str):
                    ws_log.append({"t": t, "type": "text_frame", "text": payload[:1000]})
                    return
                data = bytes(payload)
                records = _decode_records(data)
                entry = {"t": t, "type": "binary_frame", "length": len(data), "raw_hex": data.hex()}
                if records:
                    entry["decoded"] = [
                        {"kind": r[0], "internal_id": r[1], "code": r[2]} if r[0] == "code"
                        else {"kind": "quote", "internal_id": r[1], "bid": r[2], "ask": r[3], "last": r[4]}
                        for r in records
                    ]
                ws_log.append(entry)

            ws.on("framereceived", on_frame)

        popup.on("websocket", on_websocket)

        def on_request(req):
            if req.resource_type not in ("xhr", "fetch"):
                return
            net_log.append({"t": elapsed(), "type": "request", "method": req.method,
                             "url": req.url, "post_data": req.post_data})
            print(f"[{elapsed():7.2f}s] {req.resource_type} 請求: {req.method} {req.url}")

        def on_response(resp):
            req = resp.request
            if req.resource_type not in ("xhr", "fetch"):
                return
            try:
                body = resp.text()
            except PlaywrightError:
                body = None
            net_log.append({"t": elapsed(), "type": "response", "status": resp.status,
                             "url": resp.url, "body_head": body[:3000] if body else body})

        popup.on("request", on_request)
        popup.on("response", on_response)

        def on_snapshot(snap):
            entry = {"t": elapsed(), **snap}
            snapshot_log.append(entry)
            print(f"[{elapsed():7.2f}s] 畫面內容有變動，已記錄一筆快照（第 {len(snapshot_log)} 筆）。")

        _watch_and_wait(popup, context, on_snapshot)

        report.append("")
        report.append("=== 總結 ===")
        report.append(f"收到 WebSocket frame 共 {len(ws_log)} 筆"
                       f"（其中已解出已知格式 quote/code 的: "
                       f"{sum(1 for e in ws_log if e.get('decoded'))} 筆）。")
        report.append(f"收到 XHR/fetch 共 {sum(1 for e in net_log if e['type'] == 'request')} 個請求、"
                       f"{sum(1 for e in net_log if e['type'] == 'response')} 個回應。")
        report.append(f"畫面內容變動快照共 {len(snapshot_log)} 筆。")
        if net_log:
            report.append("")
            report.append("XHR/fetch 網址清單（可能藏著零股查詢的 AJAX）：")
            for url in sorted({e["url"] for e in net_log}):
                report.append(f"  {url}")

        report_path = out_dir / f"{stamp}_fastquote_manual_oddlot_摘要.txt"
        report_path.write_text("\n".join(report), encoding="utf-8")
        raw_path = out_dir / f"{stamp}_fastquote_manual_oddlot_raw.json"
        raw_path.write_text(
            json.dumps({"ws_log": ws_log, "net_log": net_log, "snapshot_log": snapshot_log},
                       ensure_ascii=False, indent=2),
            encoding="utf-8",
        )

        print("\n".join(report[3:]))
        print()
        print("=" * 70)
        print(f"摘要存在: {report_path}")
        print(f"完整原始記錄存在: {raw_path}")
        print("=" * 70)

        pause("按 Enter 關閉瀏覽器...")

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
