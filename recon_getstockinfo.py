"""
唯讀偵察腳本：驗證 `GetStockInfo` 這支 HTTP 端點在真正的盤中零股時段
（09:00~13:30）到不到得了即時零股委買賣一，以及 `dataObj` 加不加 `.O` 後綴
查到的到底是整股還是零股。

## 背景

2026/09/03 收盤後（約 16:20）用 curl 側錄過
`https://www.tbbstock.com.tw/tbb/GetStockInfo?aURL=<url-encode>http://pushex.
syspower.com.tw/Quote/mainservlet?compress=gzip&longStockName=true&type=
quote2&stockId=</url-encode>&dataObj=<股票代號[.O]>`（不需要登入也查得到），
拿 2330、0050 兩檔都測過兩種 dataObj：

    dataObj=代號       matchTime=14:30:00（盤後零股單一撮合收盤時間）、成交量遠小於整股
    dataObj=代號.O     matchTime=13:30:00（整股收盤時間）、成交量對得上整股規模

當時看資料猜的方向是「不加 .O = 零股、加 .O = 整股」——但兩組都只是「今天最後
一筆」的收盤快照，樣本也只有兩檔股票，這個方向**不可靠，不要照這個猜測寫程式**。

同一天稍晚在 FastQuote 頁面（`FastQuote/index.jsp`）親自點了畫面上正牌的
「零股查詢」按鈕（`#btnQueryO`），F12 側錄到它送出的是完全不同的東西：
`sendQUERY:57*15*2330.O`、`sendQUERY:57*31*2330.O`——**channel 57，代號帶
`.O` 後綴**，回應是兩個 26 bytes 的小封包，解出代號原樣回顯（證實 `.O` 這個
後綴慣例在這裡是「零股」沒錯，跟一開始的直覺一致），後面帶兩個小整數
（41、23299），但看起來不像價格，比較像某種統計數字，還沒確定是什麼。
緊接著頁面又送了 `addSUBSCRIBE:46*1048861`／`addSUBSCRIBE:7*1048861`，收到
一包 26.3 kB、gzip 壓縮的大封包（開頭 `1f8b0800` 是 gzip 魔術數字）——比對
2026/08/28 舊側錄，同一組指令序列在查**整股**（0050，代號沒帶 `.O`）時也會
發生，判斷 channel 46/7 是「載入圖表資料」這種通用流程，跟零股委買賣一無關。

也就是說：`GetStockInfo` 那組 curl 猜測，跟 FastQuote 頁面自己「零股查詢」按鈕
實際走的路（WebSocket channel 57），**是兩個完全不同、互相沒有印證關係的系統**
——`GetStockInfo` 打的是外部廠商 `pushex.syspower.com.tw`，channel 57 走的是
網站自己的 `wss://push.tbbstock.com.tw` 推播。兩條路都還沒真正確認查到的是不是
零股即時委買賣一，這支腳本先只驗證 `GetStockInfo` 這條路，channel 57 那條路
的委買賣一本體還沒側錄到，之後可能要另外寫一支腳本或延伸這支腳本去挖。

## 驗證方式

登入後，對同一批股票代號在盤中反覆查詢這支端點（有 `.O`、沒有 `.O` 兩種都打），
每隔 POLL_INTERVAL_SECONDS 秒查一次、共 WATCH_ROUNDS 輪，記錄每一輪兩種查詢的
matchTime／Market／TradeShares／BidPrice1／AskPrice1，藉此回答：

    1. 兩組（加 .O／不加 .O）裡哪一組的 matchTime／委買賣一會跟同一時刻
       fastquote.FastQuoteStream（channel 42，已驗證過的整股即時委買賣一）對得
       起來——對得起來的那組就是整股，另一組才有可能是零股。方向不預設，兩組
       都要看，不要只看其中一組就下結論。
    2. 另一組（跟 fastquote 對不起來的那組）的 matchTime 會不會在盤中持續跟著
       時間推進、委買賣一是不是有意義的報價（量體、價格跟整股那組有沒有差異）
       ——如果這組在盤中完全不動（還是停在舊時間點），代表這支端點在盤中查不到
       零股即時報價，這條路走不通，得回頭挖 FastQuote 頁面「零股查詢」按鈕走的
       WebSocket channel 57（見上面「背景」段落最後一次更新）。

## 安全設計

全程只做「登入 → HTTP GET 查詢 GetStockInfo → 開 FastQuoteStream 被動訂閱 →
記錄」，不點擊/填寫/送出任何下單相關元素（跟 recon_fastquote.py 同一種態度）。
GetStockInfo 本身收盤後測過不需要登入也查得到，這裡仍然借已登入 page 的
`request` context 打，只是為了跟 fastquote 那條路共用同一個瀏覽器、行為盡量
貼近正式流程，不代表確認了這支端點需要登入。
"""

import json
import sys
import traceback
from datetime import datetime

from playwright.sync_api import Error as PlaywrightError, TimeoutError as PlaywrightTimeoutError
from playwright.sync_api import sync_playwright

from fastquote import FastQuoteStream
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

GETSTOCKINFO_URL = "https://www.tbbstock.com.tw/tbb/GetStockInfo"
# 內層網址本身就帶著 stockId= 這個空參數，真正的股票代號（與要不要加 .O）
# 是外層的 dataObj 帶的——2026/09/03 curl 測出來的行為，不是憑空猜的格式。
INNER_URL = "http://pushex.syspower.com.tw/Quote/mainservlet?compress=gzip&longStockName=true&type=quote2&stockId="

DEFAULT_CODES = ["2330", "0050"]

WATCH_ROUNDS = 12
POLL_INTERVAL_SECONDS = 30  # 12 輪 * 30 秒 = 6 分鐘，盤中零股是不是持續變動看這段夠不夠


def fetch_stock_info(request_ctx, code, with_o_suffix):
    """
    查一次 GetStockInfo。with_o_suffix=True 查 dataObj=代號.O，
    False 查 dataObj=代號（裸代號）。哪一組對應零股/整股還沒確認，見模組說明。
    回傳 (解析後的 dict 或 None, 原始文字)。
    """
    data_obj = f"{code}.O" if with_o_suffix else code
    try:
        resp = request_ctx.get(GETSTOCKINFO_URL, params={"aURL": INNER_URL, "dataObj": data_obj})
    except PlaywrightError as exc:
        return None, f"請求失敗：{exc}"

    text = resp.text()
    if resp.status != 200:
        return None, f"HTTP {resp.status}\n{text[:300]}"
    try:
        rows = json.loads(text)
    except json.JSONDecodeError:
        return None, text[:300]
    if not rows:
        return None, "回應是空陣列"
    return rows[0], text


def format_row(row):
    if row is None:
        return "查詢失敗"
    return (
        f"matchTime={row.get('matchTime')} Market={row.get('Market')} "
        f"TradeShares={row.get('TradeShares')} "
        f"BidPrice1={row.get('BidPrice1')} AskPrice1={row.get('AskPrice1')} "
        f"SalePrice={row.get('SalePrice')}"
    )


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
    stamp = datetime.now().strftime("%Y%m%d_%H%M")

    configure_browsers_path()

    report = [
        f"偵察時間: {datetime.now():%Y/%m/%d %H:%M:%S}",
        f"第 {which} 組帳號，測試股票代號: {'、'.join(codes)}",
        f"共 {WATCH_ROUNDS} 輪，每輪間隔 {POLL_INTERVAL_SECONDS} 秒"
        f"（總長約 {WATCH_ROUNDS * POLL_INTERVAL_SECONDS // 60} 分鐘）。",
        "目的：確認 GetStockInfo 的 dataObj 加/不加 .O 在盤中會不會持續變動，"
        "並跟 fastquote channel 42（整股即時委買賣一）交叉比對。",
    ]

    raw_log = []  # 每一輪、每一種查詢的完整記錄，最後存成 jsonl

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

        stream = None
        try:
            stream = FastQuoteStream(page)
            stream.subscribe(codes)
            report.append("FastQuoteStream 訂閱成功，開始比對整股即時委買賣一。")
        except (PlaywrightError, PlaywrightTimeoutError) as exc:
            report.append(f"FastQuoteStream 開不起來（不影響 GetStockInfo 那組測試，繼續）：{exc}")
        print(report[-1])

        for round_no in range(1, WATCH_ROUNDS + 1):
            round_time = datetime.now()
            report.append("")
            report.append(f"=== 第 {round_no}/{WATCH_ROUNDS} 輪 {round_time:%H:%M:%S} ===")

            for code in codes:
                plain_row, plain_raw = fetch_stock_info(page.request, code, with_o_suffix=False)
                o_row, o_raw = fetch_stock_info(page.request, code, with_o_suffix=True)
                fq = stream.latest(code) if stream else None

                report.append(f"{code}  不加.O: {format_row(plain_row)}")
                report.append(f"{code}  加.O  : {format_row(o_row)}")
                report.append(f"{code}  fastquote channel42 即時(已驗證過的整股): {fq}")
                # 方向不預設：兩組都個別跟 fastquote 對一次，哪一組對得起來，
                # 哪一組才是整股——不要假設答案一定是「加 .O」。
                for label, row in (("不加.O", plain_row), ("加.O", o_row)):
                    if row is not None and fq is not None:
                        bid_ok = abs(float(row.get("BidPrice1") or 0) - fq["bid"]) < 0.01
                        ask_ok = abs(float(row.get("AskPrice1") or 0) - fq["ask"]) < 0.01
                        report.append(
                            f"{code}  {label} 跟 fastquote(整股) 對照: "
                            f"{'一致 -> 這組是整股' if bid_ok and ask_ok else '不一致'}"
                        )

                raw_log.append({
                    "round": round_no, "time": round_time.isoformat(), "code": code,
                    "plain_variant": plain_row, "o_variant": o_row, "fastquote": fq,
                })

            if round_no < WATCH_ROUNDS:
                # 用 page.wait_for_timeout 分段等，不是 time.sleep：讓 Playwright
                # 同步 API 有機會處理 WebSocket 收到的 frame，fastquote 那份對照
                # 資料才會跟著更新（跟 fastquote.py wait_for() 同一個理由）。
                page.wait_for_timeout(POLL_INTERVAL_SECONDS * 1000)

        # 逐輪比較同一個 code+variant 的 matchTime 有沒有變動過，這是判斷「這條路
        # 在盤中到底有沒有跟著即時撮合更新」最直接的證據。
        report.append("")
        report.append("=== 總結：matchTime 有沒有在這段時間內變動過 ===")
        for code in codes:
            for label, key in (("不加.O", "plain_variant"), ("加.O", "o_variant")):
                times = {
                    entry[key].get("matchTime")
                    for entry in raw_log
                    if entry["code"] == code and entry[key] is not None
                }
                times.discard(None)
                if not times:
                    report.append(f"{code} {label}: 全部查詢失敗，沒有資料")
                elif len(times) == 1:
                    report.append(f"{code} {label}: matchTime 全程沒變 = {next(iter(times))}"
                                 f"（可能盤中沒動靜，或這條路根本不會跟著即時撮合更新）")
                else:
                    report.append(f"{code} {label}: matchTime 有變動，共出現 {sorted(times)}"
                                 f"  <- 有跟著時間推進，這是這條路可行的證據")

        if stream is not None:
            stream.close()

        report_path = out_dir / f"{stamp}_getstockinfo_摘要.txt"
        report_path.write_text("\n".join(report), encoding="utf-8")
        raw_path = out_dir / f"{stamp}_getstockinfo_raw.json"
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
