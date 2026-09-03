"""
唯讀偵察腳本：直接重現 ui_order._order_quotes_job／ui_order_exec._order_fill_job
查「即時委買賣一」失敗的那條路——開 FastQuote 彈出視窗、訂閱指定股票代號、
等推播——但把 fastquote.FastQuoteStream 內部每一步都攤開記錄下來（送出訂閱
的時間點、每一筆收到的 frame 幾點幾分收到、解出哪些內部代碼/股票代號/報價），
用來回答「5512 查不到委買賣」是卡在哪一步：

    a) subscribe() 有沒有送出去成功（WebSocket 有沒有準時進入 OPEN）
    b) 送出去之後，有沒有收到任何跟 5512 有關的 frame（不管等多久）
    c) 收到的 frame 裡，代號對照（field 0x06）有沒有出現過 5512
    d) 報價（field 0x02）有沒有出現過對應 5512 內部代碼的那筆

跟 recon_fastquote.py／recon_fastquote_popup.py 一樣的安全設計：全程只做
「登入 → 開彈出視窗 → 送訂閱指令（跟正式程式碼一樣，本來就是唯讀的報價訂閱，
不是下單）→ 被動記錄 → 結束」，不點擊/填寫/送出任何下單相關元素。
"""

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
SUBSCRIBE_TIMEOUT_S = 6      # 送訂閱指令最多重試這麼久等 WebSocket 進 OPEN
WATCH_SECONDS = 20           # 送出訂閱之後被動等這麼久，看有沒有收到跟目標代號有關的 frame


def main():
    accounts = load_accounts()
    if not accounts:
        print(f"找不到帳號設定。請在 {app_dir()} 放一個 .env 檔（可複製 .env.example）。")
        sys.exit(1)

    which = 1
    code = "5512"
    if len(sys.argv) >= 2:
        try:
            which = int(sys.argv[1])
        except ValueError:
            print(f"第一個參數要是數字（第幾組帳號），收到的是: {sys.argv[1]}")
            sys.exit(1)
    if len(sys.argv) >= 3:
        code = sys.argv[2]
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
        f"第 {which} 組帳號，目標股票代號: {code}",
        "測試：跟正式程式碼一樣開 FastQuote 彈出視窗、訂閱這一檔，"
        "但把每一步的時間點跟收到的每一筆 frame 都記下來（本次不點擊/填寫任何下單元素）。",
    ]
    frame_log = []
    start = None

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

        codes_by_id = {}
        target_id = None
        got_code_record = False
        got_quote_record = False

        def on_websocket(ws):
            nonlocal start
            if "push.tbbstock.com.tw" not in ws.url:
                return
            start = time.monotonic()
            report.append(f"[{0.0:6.2f}s] WebSocket 開啟: {ws.url}")

            def on_frame(payload):
                nonlocal got_code_record, got_quote_record, target_id
                if isinstance(payload, str):
                    return
                t = (time.monotonic() - start) if start else -1
                records = _decode_records(bytes(payload))
                if not records:
                    frame_log.append(f"[{t:6.2f}s] frame {len(payload)}B，沒解出任何已知欄位")
                    return
                for rec in records:
                    if rec[0] == "code":
                        _, internal_id, rcode = rec
                        codes_by_id[internal_id] = rcode
                        hit = " <== 目標代號！" if rcode == code else ""
                        frame_log.append(f"[{t:6.2f}s] code   internal_id={internal_id} code={rcode}{hit}")
                        if rcode == code:
                            got_code_record = True
                            target_id = internal_id
                    else:
                        _, internal_id, bid, ask, last = rec
                        rcode = codes_by_id.get(internal_id, f"(未知internal_id={internal_id})")
                        hit = " <== 目標代號！" if internal_id == target_id else ""
                        frame_log.append(
                            f"[{t:6.2f}s] quote  internal_id={internal_id} code={rcode} "
                            f"bid={bid} ask={ask} last={last}{hit}")
                        if internal_id == target_id:
                            got_quote_record = True

            ws.on("framereceived", on_frame)

        popup.on("websocket", on_websocket)

        # 給 WebSocket 一點時間真的連上（跟 fastquote.FastQuoteStream.subscribe
        # 一樣：剛開完彈出視窗連線通常還在 handshake，讀不到 OPEN 就送不出訂閱）。
        popup.wait_for_timeout(500)

        cmd = f"addSUBSCRIBEX:42*{code}*"
        deadline = time.monotonic() + SUBSCRIBE_TIMEOUT_S
        attempts = 0
        sent = False
        sent_at = None
        while True:
            attempts += 1
            try:
                result = popup.evaluate(_SEND_JS, cmd)
            except PlaywrightError as exc:
                report.append(f"送訂閱指令時 evaluate 直接丟例外：{exc}")
                break
            if result == "sent":
                sent = True
                sent_at = (time.monotonic() - start) if start else None
                if sent_at is not None:
                    report.append(f"訂閱指令送出成功（第 {attempts} 次嘗試）："
                                   f"{cmd!r}，距 WebSocket 開啟 {sent_at:.2f}s")
                else:
                    report.append(f"訂閱指令送出成功（第 {attempts} 次嘗試，"
                                   f"WebSocket 尚未偵測到開啟）：{cmd!r}")
                break
            if time.monotonic() >= deadline:
                report.append(f"訂閱指令一直送不出去，{SUBSCRIBE_TIMEOUT_S}s 內重試 {attempts} 次都失敗，"
                               f"最後一次回應：{result!r}")
                break
            popup.wait_for_timeout(200)
        print("\n".join(report[-2:]))

        report.append(f"送出訂閱後被動等 {WATCH_SECONDS} 秒，記錄所有收到的 frame...")
        print(report[-1])
        popup.wait_for_timeout(WATCH_SECONDS * 1000)

        report.append("")
        report.append(f"訂閱是否送出成功: {sent}")
        report.append(f"這段期間總共收到 {len(frame_log)} 筆解出來的紀錄。")
        report.append(f"有沒有收到 {code} 的代號對照 (0x06): {got_code_record}"
                       + (f"（internal_id={target_id}）" if target_id is not None else ""))
        report.append(f"有沒有收到 {code} 的報價 (0x02): {got_quote_record}")
        if codes_by_id:
            report.append(f"這段期間總共看到 {len(codes_by_id)} 檔股票的代號對照: "
                           + ", ".join(f"{v}(id={k})" for k, v in codes_by_id.items()))
        else:
            report.append("這段期間完全沒有收到任何代號對照 (0x06) 紀錄——WebSocket 可能根本沒有推播。")

        print("\n".join(report[-5:]))

        log_path = out_dir / f"{stamp}_fastquote_subscribe_{code}.txt"
        log_path.write_text("\n".join(report) + "\n\n---- frame 明細 ----\n" + "\n".join(frame_log),
                             encoding="utf-8")
        print(f"完整記錄存在: {log_path}")

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
