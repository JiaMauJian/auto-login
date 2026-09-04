"""
唯讀偵察腳本：測一個全新、跟 channel 57 無關的假說——channel 42（已驗證過的
整股即時委買賣一推播）本身，訂閱時代號加上 `.O` 後綴，會不會回傳零股自己的
即時委買賣一。

## 為什麼不繼續查 channel 57

`recon_fastquote_channel57.py`（2026/09/04 09:13 側錄）已經把「零股查詢」按鈕
（`#btnQueryO`）實際送出的 `sendQUERY:57*15*代號.O`／`57*31*代號.O` 連續側錄
10 輪、跨 5 分鐘：兩個回應數字全程沒有變動過，不像即時委買賣一，這條路已經
測過、測完了，這支腳本不重覆。

## 新假說哪裡來的

`.O` 後綴本身在這個網站是有明確意義的（`GetStockInfo` 的 `dataObj=代號.O`、
channel 57 查詢的 `代號.O`，兩處都已經證實 `.O` = 零股）。但這兩個地方都是
「查一次、回一包」的 request/response 語意，不是 push。真正驗證過的 push
語意只有 channel 42 的 `addSUBSCRIBEX:42*代號*`（`fastquote.py`
`FastQuoteStream.subscribe()` 正式在用的那條路，委買一/委賣一/成交價已經
跟畫面對過、正確），而這條路至今只送過**沒有 `.O` 的裸代號**。這支腳本要
測的就是這個從沒試過的組合：`addSUBSCRIBEX:42*代號.O*`——如果伺服器把
`.O` 後綴也認在 channel 42 的訂閱層，同一條已知格式的推播應該會生出一個
**獨立的內部代碼**，帶著它自己會跟著撮合變動的委買賣一。

同時保留一組裸代號（不加 `.O`）當對照組，跟正式 `fastquote.py` 一樣照常訂閱，
用來確認整條 pipe 全程都是活的——如果 `.O` 那組什麼都沒收到，至少能分辨是
「伺服器不認得 `.O` 這個訂閱」還是「WebSocket 這次根本沒連上」。

## 順便留一手：field id 不預設只有 0x02/0x06

`fastquote._decode_records()` 目前只解兩種 field id（0x02 委買賣一/成交價、
0x06 股票代號），因為目前只驗證過整股需要這兩種。如果零股走同一個錨點格式
但用了目前沒解過的第三種 field id，直接呼叫 `_decode_records()` 會直接把
那筆資料吃掉、當作「沒有已知紀錄」跳過，等於白側錄。這支腳本另外用同一個
錨點 regex（`fastquote._RECORD_RE`）跑一次不限定 field id 的掃描，把
0x02/0x06 以外的 field id 也記下來（field id + 後面 16 bytes 的 hex），
留著手動分析用。

## 安全設計

全程只做「登入 → 開 FastQuote 彈出視窗 → 送兩組 `addSUBSCRIBEX:42*...*`
訂閱指令（跟 `fastquote.FastQuoteStream.subscribe()` 正式程式碼同一招，只是
其中一組代號多加 `.O`）→ 被動記錄」，不點擊/填寫/送出任何下單相關元素。
"""

import json
import sys
import time
import traceback
from datetime import datetime

from playwright.sync_api import Error as PlaywrightError, TimeoutError as PlaywrightTimeoutError
from playwright.sync_api import sync_playwright

from fastquote import _PATCH_WEBSOCKET_JS, _RECORD_RE, _SEND_JS, _decode_records
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


def _scan_all_fields(data):
    """
    跟 fastquote._decode_records() 用同一個錨點 regex，但不限定 field id 是
    0x02 或 0x06——回傳這個 frame 裡每一筆錨點命中的 (internal_id, field_id,
    後面 16 bytes 的 hex)，藉此抓「零股走同一種錨點格式、但用了目前沒解過的
    field id」這種還沒預期到的情況。field id 是 0x02/0x06 的也一併回傳，方便
    跟 _decode_records() 的結果交叉核對。
    """
    out = []
    for m in _RECORD_RE.finditer(data):
        field = m.group("field")[0]
        internal_id = int.from_bytes(m.group("iid"), "little")
        start = m.end()
        out.append((internal_id, field, data[start:start + 16].hex()))
    return out


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
    odd_codes = [f"{c}.O" for c in codes]

    out_dir = app_dir() / OUTPUT_DIR_NAME
    out_dir.mkdir(exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")

    configure_browsers_path()

    report = [
        f"偵察時間: {datetime.now():%Y/%m/%d %H:%M:%S}",
        f"第 {which} 組帳號，測試股票代號: {'、'.join(codes)}",
        f"對照組（裸代號，跟正式 fastquote.py 一樣）: {'、'.join(codes)}",
        f"實驗組（channel 42 訂閱時代號加 .O，這條組合從沒試過）: {'、'.join(odd_codes)}",
        f"共 {ROUNDS} 輪，每輪間隔 {ROUND_INTERVAL_SECONDS} 秒"
        f"（總長約 {ROUNDS * ROUND_INTERVAL_SECONDS // 60} 分鐘）。",
    ]

    codes_by_id = {}       # 內部代碼 -> 代號字串（可能是 "2330" 也可能是 "2330.O"）
    quotes_by_id = {}      # 內部代碼 -> {"bid","ask","last","t"}
    new_field_log = []     # 0x02/0x06 以外的 field id，留著手動分析
    unknown_frame_count = 0
    raw_log = []

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

        def on_websocket(ws):
            nonlocal start
            if "push.tbbstock.com.tw" not in ws.url:
                return
            start = time.monotonic()

            def on_frame(payload):
                nonlocal unknown_frame_count
                if isinstance(payload, str):
                    return
                t = (time.monotonic() - start) if start else -1
                data = bytes(payload)

                for internal_id, field, tail_hex in _scan_all_fields(data):
                    if field not in (0x02, 0x06):
                        new_field_log.append({"t": round(t, 2), "internal_id": internal_id,
                                               "field": hex(field), "tail_hex": tail_hex})

                records = _decode_records(data)
                if not records:
                    unknown_frame_count += 1
                    raw_log.append({"t": round(t, 2), "type": "unknown",
                                     "length": len(data), "raw_hex_head": data.hex()[:200]})
                    return

                for rec in records:
                    if rec[0] == "code":
                        _, internal_id, rcode = rec
                        codes_by_id[internal_id] = rcode
                        raw_log.append({"t": round(t, 2), "type": "code",
                                         "internal_id": internal_id, "code": rcode})
                    else:
                        _, internal_id, bid, ask, last = rec
                        quotes_by_id[internal_id] = {"bid": bid, "ask": ask, "last": last, "t": round(t, 2)}
                        raw_log.append({"t": round(t, 2), "type": "quote",
                                         "internal_id": internal_id,
                                         "code": codes_by_id.get(internal_id),
                                         "bid": bid, "ask": ask, "last": last})

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

        ok_plain = send("addSUBSCRIBEX:42*" + "*".join(codes) + "*")
        report.append(f"對照組訂閱（裸代號）: {'成功' if ok_plain else '失敗'}")
        popup.wait_for_timeout(500)
        ok_odd = send("addSUBSCRIBEX:42*" + "*".join(odd_codes) + "*")
        report.append(f"實驗組訂閱（.O 後綴）: {'成功' if ok_odd else '失敗'}")
        print(report[-2])
        print(report[-1])

        def snapshot_line(label, code):
            internal_id = next((iid for iid, c in codes_by_id.items() if c == code), None)
            if internal_id is None:
                return f"{label} {code}: 還沒收到代號對照 (0x06)，代表伺服器還沒認過這個訂閱代號"
            quote = quotes_by_id.get(internal_id)
            if quote is None:
                return f"{label} {code}: 有代號對照 (internal_id={internal_id})，但還沒收到報價 (0x02)"
            return (f"{label} {code}: internal_id={internal_id} "
                    f"bid={quote['bid']} ask={quote['ask']} last={quote['last']} "
                    f"(收到於 t={quote['t']}s)")

        for round_no in range(1, ROUNDS + 1):
            round_time = datetime.now()
            report.append("")
            report.append(f"=== 第 {round_no}/{ROUNDS} 輪 {round_time:%H:%M:%S} ===")
            for code, odd_code in zip(codes, odd_codes):
                report.append(snapshot_line("對照組", code))
                report.append(snapshot_line("實驗組", odd_code))
            print("\n".join(report[-1 - 2 * len(codes):]))

            if round_no < ROUNDS:
                popup.wait_for_timeout(ROUND_INTERVAL_SECONDS * 1000)

        # 結論判定：對照組（裸代號）本來就驗證過會動，這裡只是再次確認 pipe
        # 是活的；重點是實驗組（.O）有沒有收到代號對照、有沒有報價、報價會
        # 不會動、跟同時間對照組的委買賣一是不是不一樣（不一樣才是「這是獨立
        # 的零股委買賣一」最直接的證據，只是「有收到報價」還不足以證明）。
        report.append("")
        report.append("=== 總結 ===")
        for code, odd_code in zip(codes, odd_codes):
            plain_id = next((iid for iid, c in codes_by_id.items() if c == code), None)
            odd_id = next((iid for iid, c in codes_by_id.items() if c == odd_code), None)
            report.append(f"{code}: 對照組{'有' if plain_id is not None else '沒有'}收到代號對照"
                           f"，實驗組(.O){'有' if odd_id is not None else '沒有'}收到代號對照")
            if plain_id is not None and odd_id is not None:
                pq, oq = quotes_by_id.get(plain_id), quotes_by_id.get(odd_id)
                if pq and oq:
                    same = pq["bid"] == oq["bid"] and pq["ask"] == oq["ask"]
                    report.append(
                        f"  對照組 bid={pq['bid']} ask={pq['ask']} / "
                        f"實驗組 bid={oq['bid']} ask={oq['ask']} "
                        + ("<- 完全相同，不能排除只是同一份整股資料回顯，不算證實"
                           if same else
                           "<- 不一樣！這是獨立零股委買賣一的直接證據，值得接上正式程式碼")
                    )
                elif oq:
                    report.append(f"  實驗組收到報價但對照組沒有：bid={oq['bid']} ask={oq['ask']}")
                elif pq:
                    report.append("  只有對照組收到報價，實驗組完全沒有——.O 這個訂閱代號目前看起來沒用")
            elif odd_id is None:
                report.append("  實驗組完全沒有被伺服器認過，這個假說到此為止")

        if new_field_log:
            report.append("")
            report.append(f"另外側錄到 {len(new_field_log)} 筆用了 0x02/0x06 以外 field id 的紀錄，"
                           "可能是還沒解過的欄位（漲跌／五檔／零股？），明細見 raw json 的 new_field_log。")
        report.append(f"完全解不出任何已知或未知欄位的 frame 共 {unknown_frame_count} 筆。")

        report_path = out_dir / f"{stamp}_fastquote_oddlot_subscribe_摘要.txt"
        report_path.write_text("\n".join(report), encoding="utf-8")
        raw_path = out_dir / f"{stamp}_fastquote_oddlot_subscribe_raw.json"
        raw_path.write_text(json.dumps({"raw_log": raw_log, "new_field_log": new_field_log},
                                        ensure_ascii=False, indent=2), encoding="utf-8")

        print("\n".join(report[5:]))
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
