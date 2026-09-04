"""
唯讀偵察腳本：直接送出跟 FastQuote 頁面「零股查詢」按鈕（`#btnQueryO`）一樣的
WebSocket 指令（`sendQUERY:57*15*代號.O`、`sendQUERY:57*31*代號.O`），不點擊
任何按鈕，重複送、記錄每一輪收到的所有 frame，藉此回答 channel 57 那兩個小
整數（2026/09/03 側錄到 2330 是 41、23299）到底是不是委買賣一，還是別的統計值。

## 背景

見 `fastquote.py` 模組說明跟 `recon_getstockinfo.py` 模組說明那幾天的側錄過程：
FastQuote 頁面自己的「零股查詢」按鈕，F12 側錄到實際送出的不是 HTTP 請求，是
這條已經開著的 WebSocket 上兩筆 `sendQUERY` 指令。回應是 26 bytes 的小封包，
結構（側錄兩筆對照出來的，只有一個時間點的樣本，格式不保證完整）：

    byte0     channel（0x39 = 57）
    byte1     子查詢代號（0x0f=15 或 0x1f=31，對到送出去的那個數字）
    byte2-3   兩筆都是 0x58 0x13，不確定是不是真的固定
    byte4-6   保留，全 0
    byte7-19  股票代號 ASCII，13 bytes 固定欄位，代號後面補 0
    byte20-  剩下的 bytes，前 4 bytes 當 little-endian 整數解出一個值
                （2330.O：query15=41、query31=23299）

這兩個數字量級不像價格（不管除不除以 100 都兜不起來），比較像數量／筆數，但只
側錄過一次、一個時間點，沒辦法確定是不是會跟著盤中撮合變動。這支腳本重複送
同一組查詢、跨時間比對這兩個數字有沒有變化，藉此判斷它是不是即時資料——如果
在盤中反覆查會變動，值得繼續深挖；如果全程不動，這條路大概率也不是委買賣一。

同時借同一條連線訂閱 channel 42（跟 `fastquote.py` 一樣，整股即時委買賣一）當
時間軸參考，方便對照「這段期間市場真的有在動」。

## 安全設計

全程只做「登入 → 開 FastQuote 彈出視窗 → 直接送 WebSocket 文字指令（跟正式
`fastquote.FastQuoteStream.subscribe()` 送 `addSUBSCRIBEX` 同一招，只是換了
指令內容）→ 被動記錄」，不點擊/填寫/送出任何下單相關元素（跟
`recon_fastquote_subscribe.py` 同一種態度）。
"""

import json
import sys
import time
import traceback
from datetime import datetime

from playwright.sync_api import Error as PlaywrightError, TimeoutError as PlaywrightTimeoutError
from playwright.sync_api import sync_playwright

from fastquote import _PATCH_WEBSOCKET_JS, _SEND_JS, _decode_records
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
OPEN_POPUP_JS = "() => { fastQuoteUtil.openWinURL('../FastQuote/index.jsp'); }"

DEFAULT_CODES = ["2330", "0050"]
SUBSCRIBE_TIMEOUT_S = 6
ROUNDS = 10
ROUND_INTERVAL_SECONDS = 30   # 每輪間隔，總長約 ROUNDS*ROUND_INTERVAL_SECONDS 秒
PER_QUERY_WAIT_MS = 2500      # 送出一組 57*15/57*31 之後，等這麼久收頻框再換下一檔


def _decode_channel57(data):
    """
    嘗試用 2026/09/03 側錄推出來的格式解 channel 57 的回應。data[0] 不是 0x39
    (57) 或長度不夠就回 None——格式是猜出來的，不強行套用在看起來不像的封包上，
    套不上的一律留在呼叫端當 unknown 存原始 hex，不會憑空丟掉資料。
    """
    if len(data) < 20 or data[0] != 0x39:
        return None
    subquery = data[1]
    code_field = data[7:20]
    code = code_field.split(b"\x00", 1)[0].decode("ascii", "ignore")
    tail = data[20:]
    value = int.from_bytes(tail[:4], "little") if len(tail) >= 4 else None
    return {
        "subquery": subquery, "code": code,
        "value": value, "tail_hex": tail.hex(), "raw_hex": data.hex(),
    }


def main():
    accounts = load_accounts()
    if not accounts:
        print(f"找不到帳號設定。請在 {app_dir()} 放一個 .env 檔（可複製 .env.example）。")
        sys.exit(1)

    which = 1
    codes = DEFAULT_CODES
    if len(sys.argv) >= 2:
        try:
            which = int(sys.argv[1])
        except ValueError:
            print(f"第一個參數要是數字（第幾組帳號），收到的是: {sys.argv[1]}")
            sys.exit(1)
    if not 1 <= which <= len(accounts):
        print(f".env 裡目前有 {len(accounts)} 組帳號，第 {which} 組不存在。")
        sys.exit(1)
    if len(sys.argv) >= 3:
        codes = [c.strip() for c in sys.argv[2].split(",") if c.strip()]
    account = accounts[which - 1]

    out_dir = app_dir() / OUTPUT_DIR_NAME
    out_dir.mkdir(exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")

    configure_browsers_path()

    report = [
        f"偵察時間: {datetime.now():%Y/%m/%d %H:%M:%S}",
        f"第 {which} 組帳號，測試股票代號: {'、'.join(codes)}",
        f"共 {ROUNDS} 輪，每輪間隔 {ROUND_INTERVAL_SECONDS} 秒"
        f"（總長約 {ROUNDS * ROUND_INTERVAL_SECONDS // 60} 分鐘）。",
        "直接送 sendQUERY:57*15/31*代號.O（跟按「零股查詢」按鈕一樣的指令），"
        "不點擊任何畫面元素，記錄每一輪收到的回應，順便訂閱 channel 42 當參考。",
    ]

    raw_log = []  # 每一筆解出來的紀錄都存，最後存成 json 方便回頭重新分析

    with sync_playwright() as p:
        context, browser = open_context(p)
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
            report.append(f"沒有攔截到彈出視窗：{exc}")
            print(report[-1])
            wait_until_finished(context)
            sys.exit(1)

        popup.wait_for_load_state("domcontentloaded", timeout=10000)
        report.append(f"彈出視窗網址: {popup.url}")
        print(report[-1])

        start = None
        codes_by_id = {}

        def on_websocket(ws):
            nonlocal start
            if "push.tbbstock.com.tw" not in ws.url:
                return
            start = time.monotonic()

            def on_frame(payload):
                if isinstance(payload, str):
                    return
                t = (time.monotonic() - start) if start else -1
                data = bytes(payload)

                ch57 = _decode_channel57(data)
                if ch57 is not None:
                    raw_log.append({"t": round(t, 2), "type": "channel57", **ch57})
                    return

                records = _decode_records(data)
                if records:
                    for rec in records:
                        if rec[0] == "code":
                            _, internal_id, rcode = rec
                            codes_by_id[internal_id] = rcode
                            raw_log.append({"t": round(t, 2), "type": "code",
                                             "internal_id": internal_id, "code": rcode})
                        else:
                            _, internal_id, bid, ask, last = rec
                            raw_log.append({"t": round(t, 2), "type": "quote",
                                             "internal_id": internal_id,
                                             "code": codes_by_id.get(internal_id),
                                             "bid": bid, "ask": ask, "last": last})
                    return

                # 兩種已知格式都套不上，留原始 hex（截斷到 200 字避免大包塞爆
                # 摘要），不憑空丟掉。
                raw_log.append({"t": round(t, 2), "type": "unknown",
                                 "length": len(data), "raw_hex_head": data.hex()[:200]})

            ws.on("framereceived", on_frame)

        popup.on("websocket", on_websocket)
        popup.wait_for_timeout(1000)  # 給 WebSocket 一點時間連上

        def send(cmd):
            deadline = time.monotonic() + SUBSCRIBE_TIMEOUT_S
            while True:
                try:
                    result = popup.evaluate(_SEND_JS, cmd)
                except PlaywrightError:
                    return False
                if result == "sent":
                    return True
                if time.monotonic() >= deadline:
                    return False
                popup.wait_for_timeout(200)

        # channel 42 當時間軸參考，一次訂閱完，之後每輪不用重送。
        send("addSUBSCRIBEX:42*" + "*".join(codes) + "*")

        for round_no in range(1, ROUNDS + 1):
            round_time = datetime.now()
            report.append("")
            report.append(f"=== 第 {round_no}/{ROUNDS} 輪 {round_time:%H:%M:%S} ===")

            for code in codes:
                ok15 = send(f"sendQUERY:57*15*{code}.O")
                ok31 = send(f"sendQUERY:57*31*{code}.O")
                report.append(f"{code}: 送出 57*15({'成功' if ok15 else '失敗'})"
                               f"／57*31({'成功' if ok31 else '失敗'})")
                popup.wait_for_timeout(PER_QUERY_WAIT_MS)

            if round_no < ROUNDS:
                remaining_ms = ROUND_INTERVAL_SECONDS * 1000 - len(codes) * PER_QUERY_WAIT_MS
                popup.wait_for_timeout(max(0, remaining_ms))

        # 整理 channel 57 的值有沒有隨時間變動——這是判斷它是不是即時資料的關鍵證據。
        report.append("")
        report.append("=== 總結：channel 57 的值有沒有隨時間變動 ===")
        by_key = {}
        for entry in raw_log:
            if entry["type"] != "channel57":
                continue
            key = (entry["code"], entry["subquery"])
            by_key.setdefault(key, []).append((entry["t"], entry["value"]))
        if not by_key:
            report.append("完全沒有收到任何解得出來的 channel 57 回應——"
                           "可能格式跟猜的不一樣，或這條指令盤中行為不同，"
                           "去 raw json 的 unknown 那些筆看有沒有線索。")
        for (code, subquery), points in sorted(by_key.items()):
            values = sorted({v for _, v in points if v is not None})
            report.append(f"{code} query{subquery}: 收到 {len(points)} 次，"
                           f"出現過的值={values}"
                           + ("  <- 有變動，值得繼續深挖" if len(values) > 1 else "  （全程沒變）"))

        unknown_count = sum(1 for e in raw_log if e["type"] == "unknown")
        report.append(f"另外收到 {unknown_count} 筆兩種已知格式都套不上的 frame，明細在 raw json。")

        report_path = out_dir / f"{stamp}_fastquote_channel57_摘要.txt"
        report_path.write_text("\n".join(report), encoding="utf-8")
        raw_path = out_dir / f"{stamp}_fastquote_channel57_raw.json"
        raw_path.write_text(json.dumps(raw_log, ensure_ascii=False, indent=2), encoding="utf-8")

        print("\n".join(report[4:]))
        print()
        print("=" * 70)
        print(f"摘要存在: {report_path}")
        print(f"完整原始記錄存在: {raw_path}")
        print("=" * 70)

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
