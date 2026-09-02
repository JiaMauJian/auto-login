"""
即時報價 WebSocket：訂閱 FastQuote 那條推播（wss://push.tbbstock.com.tw/WEBSOCKET），
解出委買一／委賣一／成交價，給盤中模式的追價（見 orders.chase_price）跟畫面即時
顯示共用同一份資料來源——這裡只用委買一／委賣一（成交價一律來自 Excel I 欄，
見 ui_order.order_prices，2026/08/29 使用者確認拿掉了「未實現損益」那條現查
成交價的 AJAX），FastQuote 頁本來就是走這條推播更新報價，不是輪詢（見 docs/
自動下單與半自動下單規劃.pptx.txt「盤中」小節、偵察資料\\20260828_*_fastquote_*）。

## 一定要是 FastQuote 頁面自己開的那條連線，另開一條收不到任何東西

2026/08/28 實測過：在別的頁面（甚至同一個瀏覽器 context）用 `new WebSocket(url)`
接同一個網址、送一模一樣的訂閱指令，**連線會成功開啟，但完全收不到任何推播**
（等 5 秒、10 秒都一樣，readyState 全程是 OPEN，伺服器就是不推）。只有真的
`page.goto(FASTQUOTE_URL)`、讓那頁自己的 JS 開的那條連線才收得到東西——伺服器
顯然是靠連線之外的某種狀態（session？cookie？）認這條線合不合法，不是只看
URL 裡的 USER=dir&PASSWORD=password。原因沒有查到底，只確認了現象。

FastQuote 頁自己把 WebSocket 物件存在 closure 裡，`window` 上找不到任何全域
變數指到它（掃過 window 本身跟它底下一層物件的屬性都找不到），沒辦法從外面
直接拿到參考去呼叫 .send() 訂閱別的股票。解法是在頁面任何 script 執行「之前」
（`page.add_init_script`）把 `window.WebSocket` 換成一個會記錄每個實例的版本，
FastQuote 自己的 JS 建立連線時就會被這個 patch 攔下來、存進
`window.__capturedSockets`，之後就能透過這個陣列拿到那個真正的連線物件。
這個手法 2026/08/28 測過確實抓得到、也送得出訂閱指令（見下面 FastQuoteStream）。

## 另開分頁不是一律不行——要看是哪一種「另開」

2026/08/28 測到：**用 `context.new_page()` 另開一個完全獨立的分頁**，就算先
導去帳戶頁確認 cookie 有效，也一樣抓不到任何 WebSocket——那個新分頁被網站
導去了 `index/home.jsp`（不是 `/account/`），`sessionStorage` 全部是空的
（`branch_id`／`cust_id` 都是 None）。cookie 是整個 context 共用的沒錯，但
sessionStorage 不是，而且光靠「帶著有效 cookie 去逛帳戶頁」並不會讓新分頁
自己重新生出一份 sessionStorage。

但同一天稍晚用 `recon_fastquote_popup.py` 另外測了一條路：網站畫面上本來就有
一個連結（`<a href="javascript:fastQuoteUtil.openWinURL('../FastQuote/index.jsp')">
簡易看盤下單</a>`，login.py 的 `wait_until_finished()` 也處理過使用者手動點這條
連結的情境），內部是 `window.open()`。瀏覽器對「同一個分頁自己呼叫
`window.open()` 開出來的新視窗」有專屬規則：新視窗會複製一份 opener 那個分頁
當下的 sessionStorage（跟瀏覽器「複製分頁」sessionStorage 會延續是同一條
規則），跟 `context.new_page()` 那種完全獨立、沒有 opener 關係的新分頁不是
同一回事。實測結果：**用登入完的 page 呼叫 `fastQuoteUtil.openWinURL(...)`
開出來的彈出視窗，sessionStorage 是有效的**（`branch_id='112'`、
`cust_id='0108640'`，跟登入時同一組），WebSocket 也真的收到報價（見
偵察資料\\20260828_1938_fastquote_popup_摘要.txt）。也就是說能用來開 FastQuote
WebSocket 的不是只有 `login.do_login()` 回傳的那個 page 物件本身——用它
`window.open()` 出來的彈出視窗一樣有效，而且是完全獨立的分頁，不會佔用登入
分頁本身。

`FastQuoteStream` 因此不是「接管」呼叫端傳進來的 page，而是借它已經登入這件
事，用 `fastQuoteUtil.openWinURL()` 開一個彈出視窗，之後全程只碰這個彈出
視窗——呼叫端傳進來的 page 開完彈出視窗後可以立刻去做別的查詢/填單，兩者
不衝突，原本擔心的「跟哪一組帳號的 page 搶」這個資源分配問題也就不存在了。

## 訂閱「不在自選清單裡的股票」原本以為收不到，其實是 subscribe() 自己的 bug

早先那次側錄以為收盤後訂閱 2330（不在自選清單 0050、006208 裡）收不到任何
回應，懷疑是伺服器收盤後不理新訂閱、或自選清單以外的股票訂閱機制本來就不是
這樣。2026/08/28 用改成彈出視窗的正式 `FastQuoteStream` 重測才發現：問題其實
出在 `subscribe()` 自己——剛開完彈出視窗、WebSocket 才剛開始 handshake，
`readyState` 還不是 `OPEN`，這時候呼叫 `subscribe()` 幾乎必定送不出去
（早先那次測試沒有檢查 `subscribe()` 的回傳值，才誤以為指令送出去了但伺服器
不回）。`subscribe()` 已經改成會重試等到 `OPEN`（見下面），改完之後 2330
（收盤後、不在自選清單）也正常收到報價——訂閱機制本身沒有這個限制，是我們
自己的程式沒把指令真的送出去。

## 二進位協定格式（側錄反推，不是官方文件）

channel 42 的推播裡，一支股票拆成好幾段子紀錄，每段開頭都是固定的 9 bytes
錨點 `2a <欄位id 1B> 4e <長度? 1B> 00 00 00 20 00`，接著 2 bytes 股票內部代碼，
再接欄位各自的內容。長度那個 byte 的精確算法還沒對上，不用它切分紀錄——
改成整段 frame 直接掃錨點，找到幾筆算幾筆，不依賴知道上一筆在哪裡結束。

目前只解出兩種欄位，其餘（漲跌、五檔剩下四檔、成交量…）沒解，不需要就沒解：

    0x06  股票代號（ASCII，後面補 0）——用來建立「內部代碼 -> 股票代號」對照表，
          後面 0x02 那筆只帶內部代碼，得先靠這筆才知道是哪一檔股票
    0x02  委買一／委賣一／成交價，各 4 bytes little-endian、值 ÷100 才是價格。
          兩檔不同股票的真實封包交叉驗證過，數字跟畫面顯示一致（見記憶
          fastquote-ws-binary-decoded）

## expect_page() 有時候真的等不到「page」事件，但視窗其實開出來了

2026/09/02 真帳號盤中實測踩到：`_order_quotes_job` 丟出 `TimeoutError`，30 秒內
沒等到 `context.expect_page()` 的「page」事件；但事後去看 Chrome，那個彈出視窗
其實真的開出來了（標題列有「Chrome 目前受到自動測試軟體控制」，確認是同一個
自動化 context 開的，不是使用者自己手動點開），而且已經在正常收報價。也就是說
`openWinURL()` 那次呼叫**沒有失敗**，只是 Playwright 這次沒能在 timeout 內把
新視窗跟 `context.expect_page()` attach 起來——原因不明（懷疑跟 `window.open()`
帶 width/height 開成獨立視窗、不是同一個瀏覽器視窗裡的分頁有關，CDP 的
auto-attach 偶爾比較慢），只確認了現象，沒查到根因。

沒接住的那個視窗會變孤兒：例外發生在 `__init__` 裡面、`self._page` 還沒被
指派，呼叫端包的 `try/finally: stream.close()` 救不了它（那個 finally 包的是
`stream = FastQuoteStream(page)` 這一整行執行完之後）。孤兒視窗留著不會馬上壞
事，但 `openWinURL()` 開窗用固定視窗名稱，下一次呼叫如果只是抓到同一個名稱的
舊視窗（不是真的開一個新的），Playwright 又會等不到「page」事件、重演同一個
逾時。

`__init__` 因此改成兩段防呆：進來就先找一次 `context.pages` 裡有沒有 FastQuote
的孤兒分頁、有就關掉，保證每次都是從乾淨狀態開新視窗；`expect_page()` 逾時之後
不當場放棄，改成用同一招（`page.wait_for_timeout` 分段等，理由同 `wait_for()`）
在 `context.pages` 裡多等一段時間看視窗會不會自己冒出來——冒出來就照用，真的
沒有才是這次真的失敗，把原本的 `TimeoutError` 丟出去給呼叫端。
"""

import re
import threading
import time

from playwright.sync_api import Error as PlaywrightError
from playwright.sync_api import TimeoutError as PlaywrightTimeoutError

# 開彈出視窗要用網站自己的 fastQuoteUtil.openWinURL()，不是 page.goto 這個
# 絕對網址——見模組說明「另開分頁不是一律不行」，直接 goto 過去就沒有 opener
# 關係，sessionStorage 不會被複製，跟 context.new_page() 一樣抓不到身分。
_OPEN_POPUP_JS = "() => { fastQuoteUtil.openWinURL('../FastQuote/index.jsp'); }"

# 見模組說明「另開分頁不是一律不行」：這段要在彈出視窗的 script 執行之前塞進
# 去，才攔得到它自己建立的那個 WebSocket 實例。__capturedSockets 只留第一個
# （FastQuote 這頁只會開這一條）。
_PATCH_WEBSOCKET_JS = """
(function() {
    var OrigWS = window.WebSocket;
    window.__capturedSockets = [];
    function PatchedWS(url, protocols) {
        var ws = (protocols !== undefined) ? new OrigWS(url, protocols) : new OrigWS(url);
        window.__capturedSockets.push(ws);
        return ws;
    }
    PatchedWS.prototype = OrigWS.prototype;
    PatchedWS.CONNECTING = OrigWS.CONNECTING;
    PatchedWS.OPEN = OrigWS.OPEN;
    PatchedWS.CLOSING = OrigWS.CLOSING;
    PatchedWS.CLOSED = OrigWS.CLOSED;
    window.WebSocket = PatchedWS;
})();
"""

_SEND_JS = """(cmd) => {
    const ws = window.__capturedSockets && window.__capturedSockets[0];
    if (!ws) return 'no captured socket';
    if (ws.readyState !== 1) return 'not open, state=' + ws.readyState;
    ws.send(cmd);
    return 'sent';
}"""

# 見模組說明「二進位協定格式」。`.` 是萬用位元組（field id、長度各佔 1 byte），
# 中間 5 bytes `00 00 00 20 00` 是側錄裡每一筆子紀錄都有的固定值。
_RECORD_RE = re.compile(rb"\x2a(?P<field>.)\x4e.\x00\x00\x00\x20\x00(?P<iid>..)", re.DOTALL)


def _decode_records(data):
    """
    掃一個二進位封包，回傳看得懂的子紀錄：委買賣一用 ("quote", 內部代碼, 委買, 委賣, 成交)，
    股票代號用 ("code", 內部代碼, 股票代號)。看不懂的欄位（field id 不是 0x02/0x06）
    直接跳過——不是漏掉，是還沒解、也用不到。
    """
    out = []
    for m in _RECORD_RE.finditer(data):
        field = m.group("field")[0]
        internal_id = int.from_bytes(m.group("iid"), "little")
        start = m.end()

        if field == 0x06:
            raw = data[start:start + 16]
            code = raw.split(b"\x00", 1)[0].decode("ascii", "ignore").strip()
            if code:
                out.append(("code", internal_id, code))

        elif field == 0x02:
            if len(data) < start + 12:
                continue
            bid = int.from_bytes(data[start:start + 4], "little") / 100
            ask = int.from_bytes(data[start + 4:start + 8], "little") / 100
            last = int.from_bytes(data[start + 8:start + 12], "little") / 100
            # 錨點只有 9 bytes，理論上不是不可能剛好在隨機的價格資料裡假性命中；
            # 拿「像不像一個股價」擋一下，不是嚴謹的驗證，只是不要把明顯不合理
            # 的雜訊當成報價收下來。
            if 0 < bid < 100000 and 0 < ask < 100000:
                out.append(("quote", internal_id, bid, ask, last))

    return out


_FASTQUOTE_URL_PART = "FastQuote/index.jsp"


def _find_fastquote_page(context):
    """context 現有分頁裡找 FastQuote 彈出視窗（URL 含 FastQuote/index.jsp）。找不到回 None。"""
    for pg in context.pages:
        try:
            if not pg.is_closed() and _FASTQUOTE_URL_PART in pg.url:
                return pg
        except PlaywrightError:
            continue
    return None


def _wait_for_fastquote_page(page, context, timeout_ms, poll_ms=500):
    """
    見模組說明「expect_page() 有時候真的等不到 page 事件」：expect_page() 逾時
    之後用這個補救，在 context.pages 裡多等一段時間看視窗是不是其實已經開出來
    了。用 page.wait_for_timeout 分段等、不是 time.sleep——跟 wait_for() 同一個
    理由，plain sleep 不會讓 Playwright 同步 API 去處理「新分頁出現」這種協定
    訊息，事件會卡住收不到。
    """
    deadline = time.monotonic() + timeout_ms / 1000
    while True:
        found = _find_fastquote_page(context)
        if found is not None:
            return found
        remaining_ms = (deadline - time.monotonic()) * 1000
        if remaining_ms <= 0:
            return None
        page.wait_for_timeout(min(poll_ms, remaining_ms))


class FastQuoteStream:
    """
    訂閱幾檔股票的即時委買一／委賣一／成交價，背景持續更新，只能在建立這個物件的
    那個 Playwright 執行緒裡呼叫（跟其他 Playwright 物件一樣的限制，見
    ui_background.py 開頭的執行緒說明）。

    呼叫方給一個**真的跑過 `login.do_login()` 的那個 page 物件**，只是借它
    「已經登入」這件事——這個物件會用它呼叫網站自己的 `fastQuoteUtil.
    openWinURL()` 開一個獨立的彈出視窗（見模組說明「另開分頁不是一律不行」），
    之後全程只碰這個彈出視窗，傳進來的 page 完全不受影響，開完彈出視窗就可以
    立刻拿去做別的查詢／填單，不必等這個 stream 用完。

    用完要呼叫 close() 把彈出視窗關掉（連帶斷開 WebSocket）——這個彈出視窗
    除了給這個 stream 用之外沒有別的用途，呼叫方看不到它，不會有人手動關掉。
    """

    def __init__(self, page):
        self._codes_by_id = {}   # 股票內部代碼 -> 股票代號（如 "0050"）
        self._quotes = {}        # 股票代號 -> {"bid":.., "ask":.., "last":..}
        self._subscribed = set()
        # 目前只有背景那個 Playwright 執行緒會呼叫這個物件，上鎖是防呆用，
        # 不是真的有多執行緒在搶——寧可多這一道，也不要以後有人在別的執行緒
        # 呼叫 latest() 時默默拿到讀寫中途的髒資料。
        self._lock = threading.Lock()

        # context.add_init_script 是 context 層級的（彈出視窗還沒被建立出來，
        # 沒有 page 物件可以呼叫 page.add_init_script），之後這個 context 裡
        # 任何分頁一律都會被 patch 到——patch 本身只是把建構出來的 WebSocket
        # 實例記進 window.__capturedSockets，行為透明、不影響原本功能，重複
        # 註冊（例如同一個 context 裡建立第二個 FastQuoteStream）也只是多疊
        # 一層無害的 wrapper，不需要特別防止。
        context = page.context
        context.add_init_script(_PATCH_WEBSOCKET_JS)

        # 上一個 FastQuoteStream 萬一在這段 __init__ 裡例外過（見模組說明），
        # 會留下一個孤兒彈出視窗——openWinURL() 開窗用固定視窗名稱，留著不清
        # 掉的話，這次呼叫可能只是抓到那個舊視窗，不會觸發新的「page」事件，
        # expect_page() 就會白等到逾時。先清乾淨再開，保證這次一定是新視窗。
        stray = _find_fastquote_page(context)
        if stray is not None:
            try:
                stray.close()
            except PlaywrightError:
                pass

        try:
            with context.expect_page() as popup_info:
                page.evaluate(_OPEN_POPUP_JS)
            self._page = popup_info.value
        except PlaywrightTimeoutError:
            # 見模組說明「expect_page() 有時候真的等不到 page 事件」：視窗可能
            # 其實開出來了，只是沒等到通知。多等一段時間找一次，找不到才是這
            # 次真的失敗。
            self._page = _wait_for_fastquote_page(page, context, timeout_ms=15000)
            if self._page is None:
                raise

        self._page.wait_for_load_state("domcontentloaded")

        self._page.on("websocket", self._on_websocket)

    def close(self):
        """關掉彈出視窗，連帶斷開 WebSocket。可以重複呼叫。"""
        try:
            self._page.close()
        except PlaywrightError:
            pass

    def _on_websocket(self, ws):
        if "push.tbbstock.com.tw" not in ws.url:
            return
        ws.on("framereceived", self._on_frame)

    def _on_frame(self, payload):
        if isinstance(payload, str):
            return
        with self._lock:
            for record in _decode_records(bytes(payload)):
                if record[0] == "code":
                    _, internal_id, code = record
                    self._codes_by_id[internal_id] = code
                else:
                    _, internal_id, bid, ask, last = record
                    code = self._codes_by_id.get(internal_id)
                    if code:
                        self._quotes[code] = {"bid": bid, "ask": ask, "last": last}

    def subscribe(self, codes, timeout_ms=3000):
        """
        訂閱這幾檔股票的即時報價。只會累加，不會取消先前訂閱過的。

        剛建立完 stream 就馬上呼叫的話，WebSocket 連線通常還在 handshake、
        readyState 還不是 OPEN（2026/08/28 用真帳號測到：緊接著 __init__
        呼叫幾乎每次都失敗）——這裡會在 timeout_ms 之內重試等到 OPEN 或逾時，
        不是呼叫一次不行就算了。「還沒抓到 socket」（極早期，__on_websocket
        都還沒觸發）也算在同一種「還沒 ready」情況一起重試；只有真的斷線／
        頁面掛掉這種 evaluate 直接丟例外，才會立刻放棄不重試。

        回傳最後有沒有送出去。逾時或送失敗都回 False——呼叫方（例如
        chase_price 的 best_opposite）本來就把「訂閱失敗」跟「訂閱成功但沒
        推播」同樣當成「沒有即時資料」處理，不用特別分辨兩種失敗。
        """
        new = [code for code in codes if code not in self._subscribed]
        if not new:
            return True
        cmd = "addSUBSCRIBEX:42*" + "*".join(new) + "*"
        deadline = time.monotonic() + timeout_ms / 1000
        while True:
            try:
                result = self._page.evaluate(_SEND_JS, cmd)
            except PlaywrightError:
                return False
            if result == "sent":
                self._subscribed.update(new)
                return True
            remaining_ms = (deadline - time.monotonic()) * 1000
            if remaining_ms <= 0:
                return False
            self._page.wait_for_timeout(min(200, remaining_ms))

    def latest(self, code):
        """code 目前已知的最新報價（{"bid","ask","last"}）。還沒收到任何推播就是 None。"""
        with self._lock:
            quote = self._quotes.get(code)
            return dict(quote) if quote else None

    def wait_for(self, code, timeout_ms=3000, poll_ms=200):
        """
        等到 code 收到第一筆報價，或等到 timeout_ms 逾時就放棄，回傳
        latest(code)（逾時就是 None）——給「開一個 stream 就為了這一檔股票、
        等到就關掉」這種一次性用法，不必自己在呼叫端寫輪詢迴圈。

        用 self._page.wait_for_timeout 分段等，不是一次 sleep 到底：
        Playwright 的同步 API 只有在呼叫進去的時候才會處理 WebSocket 收到的
        frame，一次睡整段時間會讓事件全部卡住收不到（跟 login.py
        wait_until_finished 的說明同一個道理，也是 recon_fastquote_popup.py
        踩過的坑）。

        逾時或這檔股票收不到報價（例如不在自選清單，見模組說明「還沒驗證過
        的部分」）都回 None，不當例外拋出——呼叫方（chase_price 的
        best_opposite）本來就把 None 當成合法的「沒有即時資料，用邊界」。
        """
        deadline = time.monotonic() + timeout_ms / 1000
        while True:
            quote = self.latest(code)
            if quote is not None:
                return quote
            remaining_ms = (deadline - time.monotonic()) * 1000
            if remaining_ms <= 0:
                return None
            self._page.wait_for_timeout(min(poll_ms, remaining_ms))
