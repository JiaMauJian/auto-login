"""
唯讀偵察腳本：登入後開「簡易看盤下單」(FastQuote/index.jsp) 頁面，
被動側錄這段時間內瀏覽器對外送出的所有請求與 WebSocket 活動，
用來回答一個問題：報價是怎麼「自動更新」的——是不是像猜測的那樣一直打
GetStockInfo（HTTP 輪詢），還是走 WebSocket 推送，還是別的機制。

安全設計（刻意、不是疏漏）：
這支程式全程只做「登入 → page.goto 開頁面 → 掛被動監聽器 → 等待 → 記錄」，
沒有任何一行對這個分頁呼叫 .click() / .fill() / .press() / .select_option()，
所以不可能誤觸下單按鈕、委託方式、或任何買賣相關的元素（#orderB、#orderS、
#tabstash、#orderData 下拉選單……全部不會被程式碰到）。

副作用：因為不輸入任何股票代號，頁面可能一開始沒有任何預設報價可看，
側錄結果可能很空——這本身也是有用的資訊（代表沒有主動查詢就不會自動打）。
如果側錄結果太空，可以自己在瀏覽器裡手動輸入股票代號查詢（純看盤查詢，
不是下單），再觀察一次；但這支程式本身不會替你按任何東西。

輸出：側錄到的事件（時間戳、種類、URL）存成 txt，放在 偵察資料\\ 底下
（已加進 .gitignore）。
"""

import sys
import time
import traceback
from collections import Counter
from datetime import datetime

from playwright.sync_api import Error as PlaywrightError, TimeoutError as PlaywrightTimeoutError
from playwright.sync_api import sync_playwright

from login import (
    app_dir,
    configure_browsers_path,
    do_login,
    load_accounts,
    open_context,
    pause,
    wait_until_finished,
)

FASTQUOTE_URL = "https://www.tbbstock.com.tw/tbb/FastQuote/index.jsp?xScreenWidth=1900&xScreenHeight=950"
OUTPUT_DIR_NAME = "偵察資料"
WATCH_SECONDS = 45  # 被動側錄的秒數


def hex_dump(data, max_bytes=800):
    """
    hexdump -C 風格：16 bytes 一行，左邊 hex、右邊 ascii（不可印字元用 . 代替）。
    超過 max_bytes 只印前面這麼多就截斷——側錄到的封包裡有些是上萬 bytes 的
    圖表資料（StockChart/TAChart 那些訂閱），跟五檔/報價無關，全印出來只會
    把檔案灌爆、淹沒真正要找的小封包，截斷只是不逐位元組印完，筆數跟總長度
    還是照實記錄。
    """
    shown = data[:max_bytes]
    lines = []
    for i in range(0, len(shown), 16):
        chunk = shown[i:i + 16]
        hex_part = " ".join(f"{b:02x}" for b in chunk)
        ascii_part = "".join(chr(b) if 32 <= b < 127 else "." for b in chunk)
        lines.append(f"  {i:04x}  {hex_part:<47}  {ascii_part}")
    if len(data) > max_bytes:
        lines.append(f"  ...（截斷，總長 {len(data)}B）")
    return "\n".join(lines)


def watch_network(page, seconds):
    """
    掛上被動監聽器，記錄這段時間內的請求與 WebSocket 活動，
    然後只是等待，不對頁面做任何操作。

    回傳 (events, raw_frames)：events 是給人看的摘要（二進位只記大小），
    raw_frames 額外存下二進位封包的實際 bytes，拿來另外 hex dump 分析協定
    格式用——events 裡只記大小這件事本身沒有問題，但要反推五檔欄位在
    第幾個 byte，非得看到真正的內容不可。
    """
    events = []
    raw_frames = []
    start = time.monotonic()

    def on_request(request):
        events.append((time.monotonic() - start, "request", request.method, request.url))

    def on_websocket(ws):
        events.append((time.monotonic() - start, "websocket-open", "", ws.url))

        def on_frame_sent(payload):
            text = payload if isinstance(payload, str) else f"<binary {len(payload)}B>"
            events.append((time.monotonic() - start, "ws-send", "", text[:200]))
            if not isinstance(payload, str):
                raw_frames.append((time.monotonic() - start, "send", bytes(payload)))

        def on_frame_received(payload):
            text = payload if isinstance(payload, str) else f"<binary {len(payload)}B>"
            events.append((time.monotonic() - start, "ws-recv", "", text[:200]))
            if not isinstance(payload, str):
                raw_frames.append((time.monotonic() - start, "recv", bytes(payload)))

        ws.on("framesent", on_frame_sent)
        ws.on("framereceived", on_frame_received)
        ws.on("close", lambda: events.append((time.monotonic() - start, "websocket-close", "", ws.url)))

    page.on("request", on_request)
    page.on("websocket", on_websocket)

    page.wait_for_timeout(seconds * 1000)

    return events, raw_frames


def summarize(events):
    lines = [f"側錄到 {len(events)} 筆事件。"]

    ws_opens = [e for e in events if e[1] == "websocket-open"]
    if ws_opens:
        lines.append(f"偵測到 {len(ws_opens)} 個 WebSocket 連線：")
        for _, _, _, url in ws_opens:
            lines.append(f"  {url}")
        ws_frames = [e for e in events if e[1] in ("ws-send", "ws-recv")]
        lines.append(f"WebSocket 封包共 {len(ws_frames)} 筆（明細看原始側錄檔）。")
    else:
        lines.append("沒有偵測到 WebSocket 連線。")

    url_counts = Counter(url for _, kind, _, url in events if kind == "request")
    repeated = sorted(
        ((url, n) for url, n in url_counts.items() if n > 1),
        key=lambda pair: -pair[1],
    )

    if repeated:
        lines.append("以下網址在側錄期間被重複呼叫（可能就是輪詢）：")
        for url, n in repeated[:20]:
            times = [t for t, kind, _, u in events if kind == "request" and u == url]
            gaps = [round(b - a, 2) for a, b in zip(times, times[1:])]
            lines.append(f"  x{n}  間隔(秒)={gaps}")
            lines.append(f"       {url}")
    else:
        lines.append("沒有觀察到重複呼叫同一網址（可能輪詢間隔比側錄時間長，或走 WebSocket，或根本沒查任何股票）。")

    return lines


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
        f"第 {which} 組帳號，被動側錄 {WATCH_SECONDS} 秒",
        "本次全程不點擊、不輸入任何內容，只被動側錄網路流量（不會下單）。",
    ]

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

        page.goto(FASTQUOTE_URL)
        report.append(f"已開啟: {page.url}")
        report.append(f"開始側錄 {WATCH_SECONDS} 秒，請不要手動操作這個分頁（避免混進你自己的點擊/查詢）...")
        print("\n".join(report))

        events, raw_frames = watch_network(page, WATCH_SECONDS)
        tail = [""]
        tail.extend(summarize(events))

        log_path = out_dir / f"{stamp}_fastquote_raw.txt"
        log_path.write_text(
            "\n".join(f"{t:.2f}s {kind} {method} {url}" for t, kind, method, url in events),
            encoding="utf-8",
        )
        tail.append(f"完整原始側錄記錄: {log_path.name}")

        binary_lines = []
        for t, direction, data in raw_frames:
            binary_lines.append(f"{t:.2f}s {direction} {len(data)}B")
            binary_lines.append(hex_dump(data))
            binary_lines.append("")
        binary_path = out_dir / f"{stamp}_fastquote_binary.txt"
        binary_path.write_text("\n".join(binary_lines), encoding="utf-8")
        tail.append(f"二進位封包 hex dump: {binary_path.name}（共 {len(raw_frames)} 筆，單筆超過 800B 截斷）")

        report.extend(tail)
        report_path = out_dir / f"{stamp}_fastquote_摘要.txt"
        report_path.write_text("\n".join(report), encoding="utf-8")

        print("\n".join(tail))
        print(f"檔案都在: {out_dir}")

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
