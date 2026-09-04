"""
即時委買賣一：打券商網站的 `GetStockInfo`，一次查一檔，整股零股都查得到。

給盤中追價（`orders.chase_price`）跟下單分頁的「查詢委買賣」按鈕共用同一份
資料來源，取代原本走 WebSocket 的 `fastquote.FastQuoteStream`（那支留著沒刪，
見它的模組說明）。

## 為什麼從 WebSocket 換成這條

`fastquote.py` 那條路要「登入 → 用 `fastQuoteUtil.openWinURL()` 開彈出視窗 →
在頁面 script 跑之前 patch `window.WebSocket` 攔下連線 → 送訂閱指令 → 解一份
側錄反推出來的二進位格式」，而且 2026/09/02 盤中實測踩過 `expect_page()` 等不到
新視窗、留下孤兒視窗的坑。這條路只要一個 HTTP GET，**不用登入、不用瀏覽器、
不用 cookie**（2026/09/04 用 `urllib` 裸打驗證過，連 header 都只帶 User-Agent）。

更關鍵的是零股：`fastquote` 的 channel 42 只有整股，零股那本簿子它拿不到
（channel 57 那條路 2026/09/04 側錄過，值全程不動、代號也解不出來，是死路）。
這裡 `odd=True` 就查得到，而且是真的另一本簿子——同一時刻 2454 整股
4395/4400、零股 4385/4390，差兩檔（見 偵察資料\\20260904_0939_零股委買賣_
GetStockInfo_摘要.txt）。

## 這是一個兩層的代理，回來的東西是 CSV 轉的

    我們 ──https──> www.tbbstock.com.tw/tbb/GetStockInfo   （券商自己的 servlet）
                            │ http
                            ↓
                  pushex.syspower.com.tw/Quote/mainservlet  （外部行情商）

`aURL` 是要轉發去的內層網址，`dataObj` 是股票代號。內層回的是 **CSV**（第一行
欄位名、之後一行一檔，`compress=gzip` 只決定壓不壓，內容一樣），外層負責轉發、
解 gzip、把 CSV 轉成 JSON 陣列再用 https 回我們——所以 JSON 那些 key 就是 CSV
第一行的欄位名，不是券商自己取的。

## 為什麼刻意不做批次查詢

外層其實吃得下逗號分隔的多檔（實測送 20 檔、20 筆全回，沒有斷），但**只有整股
可以**，而且混查會靜靜出錯，這三條是 2026/09/04 一條一條打出來的：

    dataObj=2330,2317,2454,...    整股批次 OK
    dataObj=2330.O,2454.O         回 0 筆——零股一次只能一檔
    dataObj=2330,2330.O           回 1 筆，而且 Market=0

最後那條是地雷：`.O` 被無聲吃掉、降級成整股，不報錯、不回空，就給你一筆整股
資料。誰要是想「整股零股一次撈完」，會拿到一半是整股的東西還以為是零股。
呼叫端本來就是一次一檔（追價一筆一檔、查詢委買賣一檔一個 GET），沒有批次的
需求，所以這支函式**只吃一檔**——不開這個門，就不用防門後面那顆地雷。

（順帶一提 `.O` 是外層自己認的，不是內層：`stockId=2330.O` 直接打內層回 0 筆，
只有經過外層的 `dataObj` 才有效，而且認得很死——小寫 `.o`、尾巴多一個逗號都
是 0 筆。所以下面組 `dataObj` 的時候不要「順手」改大小寫或補分隔符。）
"""

import json
import urllib.parse
import urllib.request

_URL = "https://www.tbbstock.com.tw/tbb/GetStockInfo"

# 外層要轉發去的內層網址。`stockId=` 本來就是空的，真正的代號是外層的 dataObj
# 帶的——這是網站自己頁面在用的形狀，不是我們拼出來的，不要「順手」把代號填
# 進這個 stockId。
_INNER = ("http://pushex.syspower.com.tw/Quote/mainservlet?compress=gzip"
          "&longStockName=true&type=quote2&stockId=")

# 網站頁面自己打的時候是瀏覽器發的請求，這裡至少帶一個像瀏覽器的 UA。實測不帶
# 也查得到，帶著只是不要看起來太像機器。
_HEADERS = {"User-Agent": "Mozilla/5.0"}


def quote(code, odd=False, timeout=10):
    """
    查一檔的委買一／委賣一／成交價，回 `{"bid","ask","last"}`。

    回傳形狀刻意跟 `fastquote.FastQuoteStream.latest()` 一模一樣，兩個呼叫端
    （`ui_order_exec` 的追價、`ui_order.fetch_order_quotes`）不必改讀法。

    `odd=True` 查盤中零股那本簿子（`dataObj=代號.O`），量的單位是**股**；
    `odd=False` 查整股，量的單位是**張**。這裡只回價格不回量，兩者沒有差別，
    但將來要是有人加了「量」進來，這個單位差異要記得。

    查得到但沒有資料回 **None**——盤中零股 09:10 第一次撮合之前、收盤後、
    代號不存在都是這一種，是正常情況不是錯誤。連不上／逾時／回來的東西不是
    預期格式一律**丟例外**，讓呼叫端自己決定要擋還是要吞（追價那條會吞掉退回
    邊界價，按鈕那條會跳訊息）。
    """
    data_obj = f"{code}.O" if odd else code
    url = _URL + "?" + urllib.parse.urlencode({"aURL": _INNER, "dataObj": data_obj})
    req = urllib.request.Request(url, headers=_HEADERS)
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        rows = json.loads(resp.read().decode("utf-8", "ignore"))

    if not rows:
        return None
    row = rows[0]

    # 收盤前的零股（09:10 之前）欄位是在的，但價格全是 "0.00"——這不是報價，
    # 當作「還沒有資料」回 None，不要讓 0 元一路流進 chase_price 去算價格。
    try:
        bid = float(row["BidPrice1"])
        ask = float(row["AskPrice1"])
        last = float(row["SalePrice"])
    except (KeyError, TypeError, ValueError) as exc:
        raise ValueError(f"{data_obj} 回來的資料看不懂：{row}") from exc
    if bid <= 0 or ask <= 0:
        return None

    return {"bid": bid, "ask": ask, "last": last}
