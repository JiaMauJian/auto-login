"""
下單分頁的「依序執行」引擎：把凍結好的一份委託清單一筆一筆送出去，跑完一輪
再決定要不要接下一輪。

跟 ui_order.py 分開的界線是**「這一輪要送什麼」跟「怎麼把它送出去」**：
ui_order.py 收使用者的輸入（選帳戶、選股票、填比重）、讀 Excel、查即時報價、
算出執行預覽；這裡從按下「開始下單」那一刻接手——把當下的設定整份凍結起來
（見 _order_exec_init_state 的說明），之後每一筆、每一輪都只認凍結的那份，
使用者中途改了畫面上的東西都只影響下一輪。

所以這裡完全不管左邊那格填的是什麼：不論是出清股票、買賣股票還是全持股交易
（見 docs/介面規劃.md 第九節），送到這裡的都是同一種形狀的 queue。三個作業
共用這一整套，不是各自複製一份。

半自動填單只做到「開出委託確認視窗」為止；要不要真的按下裡面的「確認」由
order_auto_confirm 決定（關＝停在那裡等人按，開＝程式自己按）。
"""

import threading
import tkinter as tk

from playwright.sync_api import Error as PlaywrightError

import excel_io
import fastquote
import fetch as fetch_mod
import order_fill
import orders
from ui_common import ask_confirm, show_error, show_info, show_warning

# 「多輪直到出清」的安全上限：不管有沒有真的出清，跑滿這個輪數就一定停下來
# 等人看過再決定要不要繼續，不無限跑下去——2026/08/28 使用者確認要做這個
# 開關的時候沒特別要求上限數字，這裡抓一個「明顯比一般情境需要的輪數多，
# 但出問題也不會跑太久」的值，不是照文件或使用者指定的數字。
ORDER_MULTI_ROUND_CAP = 10


class UiOrderExecMixin:
    # ---------- 依序執行（半自動下單） ----------
    #
    # 跟 order_fill.py 同一個界線：程式只填到「開出委託確認視窗」為止，不按
    # 裡面的「確認」，那一步留給人。這裡多做的是把執行預覽裡一整批委託接起來，
    # 一筆一筆換帳戶（換 cookie，不重登，見 fetch.new_store）自動填單——但兩筆
    # 之間一定要等使用者在瀏覽器裡把上一筆的確認視窗處理掉（按確認或取消）才
    # 能繼續，不能自動往下接：整個瀏覽器只有一組 cookie，要是視窗還開著就換了
    # 下一個帳戶的 cookie，使用者事後才回去按那個視窗的「確認」，送出的會是
    # 「現在瀏覽器帶著的身分」而不是視窗原本對應的那個帳戶——這是真的會送錯
    # 帳戶的委託，不是好看不好看的問題。「下一筆」按鈕只在背景執行緒親眼確認
    # 視窗真的關了（_order_dialog_closed）之後才解鎖，就是為了擋住這條路。
    #
    # 這一輪按下「開始下單」凍結的 self.order_exec_queue 用的是 self.busy
    # （見 ui_background._set_busy）——跟更新分頁共用同一個總開關，是因為
    # 兩邊操作的是同一個瀏覽器 context、同一組 cookie：這一輪還沒跑完時，
    # 「登入」「更新」「全部登出」都會換手上這組 cookie，一樣會製造上面
    # 那種送錯帳戶的風險，所以整輪期間直接借用同一顆busy鎖把那幾顆按鈕鎖住。

    def _order_exec_init_state(self):
        """_order_init_state 呼叫一次（見 ui_order.py）。"""
        # 半自動送單（見 start_order_execution）：這一輪凍結起來要依序送出的
        # 委託清單（orders.executable_orders／executable_intraday_orders 過濾過
        # 的執行預覽），跟 order_rows／order_holdings 這些會一直變動的畫面狀態
        # 分開——按下「開始下單」那一刻就定案，之後就算使用者又改了比重或
        # 勾選，也不會半路影響正在跑的這一輪。order_exec_mode／order_exec_ticks
        # ／order_exec_auto 也是同一輪按下去那一刻凍結的（不是每筆下單前重讀
        # 畫面上的即時值），理由跟凍結 queue 一樣。
        self.order_exec_queue = []
        self.order_exec_pos = 0        # 下一筆（或正在處理的這一筆）在 queue 裡的位置
        self.order_exec_busy = False   # 背景正在登入／填單，還沒回話
        self.order_exec_watching = False  # 委託確認視窗開著，等它關閉才能按「下一筆」
        self.order_exec_mode = "pre"
        self.order_exec_ticks = None
        self.order_exec_auto = False
        self.order_exec_side = orders.SIDE_SELL
        self.order_exec_last_note = ""    # 上一筆自動送出的結果，_update_order_exec_ui 顯示用

        # 多輪直到出清（見 _order_round_finished／_prepare_next_round）：跟
        # order_exec_mode 這些一樣，是按下「開始下單」那一刻凍結的值，不是
        # 每輪重讀畫面。order_exec_stock_settings／order_exec_accounts 也要
        # 凍結——第 2 輪以後要拿同一組股票設定、同一批帳戶去對「重讀之後的
        # 新持股」重新組隊列，不是每輪重新看畫面上現在勾了誰、填了什麼。
        self.order_multi_round = tk.BooleanVar(value=False)
        self.order_auto_price = tk.BooleanVar(value=False)
        self.order_exec_multi_round = False
        self.order_exec_auto_price = False
        self.order_exec_round = 1
        self.order_exec_stock_settings = []
        self.order_exec_accounts = []
        # 代號 -> Excel 讀回來的股價，盤中模式 _order_fill_job 算追價時一律
        # 拿這裡的值當 pricenow，不現查網頁成交價（2026/08/29 使用者確認：
        # 成交價已經在新增股票／讀取持股時讀進 Excel，不必等下單前
        # 才查）。開始下單那一刻先用 self.order_prices（見那裡的說明）當第 1
        # 輪的起始值（start_order_execution），之後每輪重讀就整份換掉，不是
        # 累加（同一檔股票這一輪的價格只有一個版本）——order_exec_auto_price
        # 這個開關現在只決定重讀前要不要先觸發「更新股價」巨集，不再決定
        # pricenow 走哪條路（見 _on_order_price_refresh）。
        self.order_exec_prices = {}
        # 開始下單那一刻凍結的 self.order_quotes 快照，只給第 1 輪用——第 2
        # 輪以後在 _on_order_price_refresh 會清成空字典，逼所有列都退回
        # 「下單前才查」那條舊路，不讓第 1 輪查到的即時報價被沿用到之後幾輪
        # （市場已經過了一段時間，繼續當最新報價用是在猜數字）。
        self.order_exec_quotes = {}
        self.order_exec_price_busy = False  # 輪與輪之間正在重讀 Excel／觸發巨集，還沒回話
        # 「這一整批多輪出清作業還在不在跑」，跟 order_exec_queue 分開——
        # 輪與輪之間 queue 會是空的（下一輪還沒組出來），但整批作業並沒有
        # 結束，「停止」要能在這個空檔也按得下去（見 stop_order_execution）。
        self.order_exec_active = False

    def start_order_execution(self):
        """
        「開始下單／下一筆」共用同一顆按鈕：queue 是空的就是「開始」（重新算一次
        執行預覽、跳確認視窗、凍結成 queue），queue 還有東西就是「下一筆」，
        直接送出目前這一筆（可能是全新的一筆，也可能是上一筆失敗後的重試，
        見 _on_order_filled）。

        模式（盤前／盤中）、盤中的追價檔數、買賣方向，都在這裡凍結成
        self.order_exec_mode／self.order_exec_ticks／self.order_exec_side，
        之後每一筆都用凍結的值，不是每次都重讀畫面——理由跟凍結 queue 一樣：
        按下「開始下單」之後使用者還是可以去改上面的設定，那些改動只影響
        「下一輪」，不能半路插進正在跑的這一輪。
        """
        if self.order_exec_queue:
            self._dispatch_next_order()
            return

        # 看 _excel_in_use()：勾了「自動更新股價」的話，按下去第一件事就是
        # _prepare_next_round → 用 COM 跑巨集、重讀 Excel（見那條路），不能在
        # 別人正在動同一份活頁簿的時候開始。
        if self._excel_in_use():
            return

        job = self.order_job.get()
        mode = self.order_mode.get()
        side = self.order_side.get()
        multi_round = self.order_multi_round.get()
        auto_price = self.order_auto_price.get()
        ordered = self._order_execution_accounts()
        # 凍結起來給多輪用的股票設定（比重／價格）。只有出清作業有這種東西，
        # 買賣股票的數字全在 Excel 那頁，畫面上一個都沒有——那個作業也不支援
        # 多輪（切作業時就強制關掉了，見 ui_order._on_order_job_changed），
        # 所以留空清單，不是漏給。
        stock_settings = []

        if job == orders.JOB_TRADE:
            # 買賣股票：張數、價格、方向全部來自各帳戶自己那一頁的下單試算，
            # 畫面上沒有任何數字要讀（見 orders.plan_trade_orders）。價格已經是
            # 數字，所以濾法跟盤前那條一樣用 executable_orders。
            ticks = None
            preview = orders.plan_trade_orders(
                [{"code": row["code"], "name": row["name"]} for row in self.order_rows],
                ordered, self.order_plans, self.order_holdings, self.order_unit.get(),
                loaded_sheets=self.order_loaded)
            queue_rows = orders.executable_orders(preview)
        elif self._order_intraday():
            stock_settings = self._order_stock_settings()
            ticks = self._order_ticks_setting()
            if ticks is None:
                show_error(self.root, "追價檔數不對", "追價檔數要填 0 以上的整數。")
                return
            preview = orders.plan_intraday_orders(
                stock_settings, ordered, self.order_holdings, ticks, side,
                prices=self.order_prices, quotes=self.order_quotes)
            queue_rows = orders.executable_intraday_orders(preview)
        else:
            ticks = None
            stock_settings = self._order_stock_settings()
            preview = orders.plan_stock_orders(stock_settings, ordered,
                                               self.order_holdings, side)
            queue_rows = orders.executable_orders(preview)

        if not queue_rows:
            if job == orders.JOB_TRADE:
                reason = ("勾選的帳戶都沒有這幾檔、下單試算是空的、只有零股，"
                          "或沒填價格")
            elif self._order_intraday():
                reason = "沒有持股，或比重不到 1 張"
            else:
                reason = "沒有持股、比重不到 1 張，或還沒填價格"
            show_info(self.root, "沒有可以執行的委託",
                f"執行預覽裡沒有一列可以送出委託（{reason}）。")
            return

        # 「共 N 張／N 股」那句 2026/09/01 拿掉了（使用者指定）：總量在執行預覽
        # 那張表上一列一列看得到，確認框再報一次總和是重複講同一件事。留下來的
        # 數字都是「那張表上數不出來」的東西——幾筆、買幾筆賣幾筆、用哪種委託別。
        auto = self.order_auto_confirm.get()
        side_word = "買進" if side == orders.SIDE_BUY else "賣出"

        if job == orders.JOB_TRADE:
            # 買賣股票的方向是逐筆的（試算正數買、負數賣），不能像出清那樣用一句
            # 「即將賣出 N 筆」帶過——那會讓人以為整批同一個方向。買賣確實可能
            # 混在同一輪裡（不同股票試算正負不同），所以買、賣兩個數字都要列。
            buys = sum(1 for row in queue_rows if row["side"] == orders.SIDE_BUY)
            unit_word = orders.UNIT_NAMES[self.order_unit.get()]
            head = (
                f"{unit_word}流程（ROD-當日有效）\n"
                f"即將依序處理 {len(queue_rows)} 筆委託\n"
                f"買 {buys} 筆、賣 {len(queue_rows) - buys} 筆\n\n"
            )
        elif self._order_intraday():
            # 按過「查詢委買賣」的那幾筆 row["price"] 已經是算好的數字（見
            # orders.plan_intraday_orders），開始下單時會直接照用、不再重查
            # （2026/08/29 使用者要求）；沒查過的那幾筆還是老路，下單前才
            # 臨時查一次——三種措辭分開講，不能讓人以為全部都是「現在看到
            # 的數字不會變」或全部都是「還會再變」，兩種情況可能同時存在。
            frozen = sum(1 for row in queue_rows if row["price"] is not None)
            if frozen == len(queue_rows):
                price_note = (f"每一筆的價格都已經用「查詢委買賣」查到的即時報價算好，"
                              f"下單就是直接用執行預覽上看到的數字，不會再重查。\n")
            elif frozen == 0:
                price_note = (f"每一筆的價格以 Excel 讀到的成交價為基準，下單前再追 {ticks} 檔、"
                              f"跟對手方第一檔比價算出來，不是現在看到的數字。\n")
            else:
                price_note = (f"其中 {frozen} 筆已經用「查詢委買賣」查到的即時報價算好，"
                              f"直接用執行預覽上看到的數字；其餘 {len(queue_rows) - frozen} 筆"
                              f"還沒查過，下單前才會即時查一次算出來。\n")
            head = (
                f"即將依序處理 {len(queue_rows)} 筆「{side_word}」委託，用 IOC。\n\n"
                f"{price_note}沒成交的部位 IOC 會自動取消，不會掛著。\n"
            )
        else:
            head = f"即將依序處理 {len(queue_rows)} 筆「{side_word}」委託，用 ROD-當日有效。\n\n"

        # 第一句先講是誰。2026/09/02 起一輪可以跑好幾位（勾選），所以這一句要
        # 列出**真的會被送出委託的那幾位**——不是勾了誰，是 queue 裡真的有東西
        # 的那幾位（勾了但沒這一檔、或試算是空的，整位都會被略過，寫進來只會
        # 讓人以為那一位也要下單）。勾錯一位整輪都會掛到別人帳上，而這一句是
        # 委託送出去之前最後一次讓人看見名字的機會。
        #
        # 超過 6 位就只報數字：這個框是要人看完再按的，一行 20 個名字沒有人會
        # 真的讀，逐位確認的地方本來就是執行預覽那張表。
        who = [account["sheet"] for account in ordered
               if any(row["sheet"] == account["sheet"] for row in queue_rows)]
        head = (f"帳戶：{'、'.join(who)}\n" if len(who) <= 6
                else f"帳戶：{len(who)} 位（依報酬率由低到高）\n") + head

        if multi_round:
            if auto_price:
                price_note = "每一輪開始前會先觸發 Excel 的「更新股價」巨集、重讀最新股價。\n"
            elif mode == "intraday":
                price_note = ("沒勾「自動更新股價」：後面幾輪的追價基準價還是會重讀一次 Excel，"
                              "但不會先觸發「更新股價」巨集，數字通常跟這一輪相同。\n")
            else:
                price_note = "盤前的價格不會跟著更新，後面幾輪會沿用你現在填的這個價格。\n"
            head += (
                f"已勾選「多輪直到出清」：這一輪的委託處理完之後，會自動重讀持股，"
                f"還有沒出清的部位就自動接下一輪，最多跑 {ORDER_MULTI_ROUND_CAP} 輪，"
                f"跑滿還沒出清會停下來等你決定。\n{price_note}"
                f"每一筆委託是否要停下來等你確認，規則不受影響。\n\n"
            )

        if auto:
            tail = (
                f"本輪為「自動送出」模式：每一筆委託皆會自動按下委託確認視窗中的"
                f"「確認」，不再等待人工確認，委託將直接送出。\n"
                f"執行節奏不變，仍為逐筆處理——視窗自動關閉後「下一筆」才會轉為"
                f"可按。\n\n"
                f"確定要開始嗎？"
            )
        else:
            tail = (
                f"每一筆都會停在委託確認視窗，不會自動送出，請自己確認或取消。\n"
                f"視窗關閉後「下一筆」才會亮起——換下一筆之前，瀏覽器裡不要留著沒處理的確認視窗。\n\n"
                f"確定要開始嗎？"
            )
        if not ask_confirm(self.root, "開始下單", head + tail, confirm_style="primary"):
            return

        self.order_exec_mode = mode
        self.order_exec_ticks = ticks
        self.order_exec_auto = auto
        self.order_exec_side = side
        self.order_exec_multi_round = multi_round
        self.order_exec_auto_price = auto_price
        self.order_exec_stock_settings = stock_settings
        self.order_exec_accounts = ordered
        # 第 1 輪（沒勾自動更新股價時就是唯一一輪）直接拿 self.order_prices
        # 當起點——那是新增股票／上次「讀取持股」讀進來的 Excel
        # 成交價，不用再另外查一次。勾了自動更新股價的話，這份值一送進
        # _prepare_next_round 馬上就會被剛重讀（含觸發巨集）的結果整份蓋掉
        # （見 _on_order_price_refresh），不是兩份資料混用。
        self.order_exec_prices = dict(self.order_prices)
        # 凍結這一刻查到的即時委買賣（見 order_exec_quotes 開頭的說明：只給
        # 第 1 輪用，第 2 輪以後 _on_order_price_refresh 會清空）。
        self.order_exec_quotes = dict(self.order_quotes)
        self.order_exec_active = True
        self._set_busy(True, "下單：準備第 1 輪…" if multi_round else "下單：準備第 1 筆…")

        if auto_price:
            # 「每一輪開始前都先觸發更新股價，包含第一輪」（規劃文件）——第
            # 一輪不能例外用這裡先算好的 queue_rows 直接送，要跟第 2 輪以後
            # 走同一條路（_prepare_next_round）：先觸發巨集、重讀 Excel，
            # 用讀回來的持股／股價重新組隊列再開始送。round 從 0 開始，
            # _on_order_price_refresh 判斷「還有東西可以送」時會 +1 變成
            # 第 1 輪——跟後面每一輪 +1 的邏輯是同一條，不必另外寫一次。
            self.order_exec_round = 0
            self._prepare_next_round()
        else:
            self.order_exec_queue = queue_rows
            self.order_exec_pos = 0
            self.order_exec_round = 1
            self._dispatch_next_order()

    def _dispatch_next_order(self):
        row = self.order_exec_queue[self.order_exec_pos]
        order_number = self._order_number_for_sheet(row["sheet"])
        if order_number is None:
            show_error(self.root, "找不到帳戶", f"{row['sheet']} 對不到任何一組帳號，這筆沒辦法執行，這一輪停止。")
            self.order_exec_active = False
            self.order_exec_queue = []
            self.order_exec_pos = 0
            self._set_busy(False)
            self._update_order_exec_ui()
            return

        account = self.accounts[order_number - 1]
        self.order_exec_busy = True
        self._update_order_exec_ui()
        self._say(f"下單：第 {self.order_exec_pos + 1}/{len(self.order_exec_queue)} 筆"
                  f"（{row['sheet']} {row['code']}）登入／填單中…")
        self._ensure_browser_thread()
        self.browser_waiting += 1
        self.browser_cmd_queue.put(
            ("order", (order_number, account, row, self.order_exec_mode,
                      self.order_exec_ticks, self.order_exec_auto, self.order_exec_side)))

    def stop_order_execution(self):
        """
        放棄整批作業（可能不只一輪）。瀏覽器裡如果還留著一個沒處理的委託
        確認視窗，程式不會再幫忙追蹤它——那一頁還在，使用者自己回去按
        「確認」或「取消」就好，只是「下一筆」的自動流程到這裡為止。

        order_exec_active 而不是 order_exec_queue 當守門——多輪出清輪與輪
        之間 queue 是空的（正在背景重讀 Excel），這個空檔也要按得下「停止」，
        不能因為 queue 剛好是空的就當作沒什麼在跑。這個空檔按下去沒辦法
        真的中斷背景那條讀 Excel 的執行緒，只能讓它讀完之後的結果被
        _on_order_price_refresh 忽略掉（見那裡的 order_exec_active 檢查）。
        """
        if not self.order_exec_active:
            return

        # 停止不問「確定嗎」（2026/09/01 使用者指定）：要停的人已經決定了，中間
        # 再擋一個對話框只是拖時間——而按停止的時候通常正是最急的時候。它也不是
        # 不可逆的動作，停完再按一次「開始」就是重新算一輪。
        #
        # 原本確認框裡那兩句提醒不能跟著消失，改成停完之後寫在狀態列上：那兩件事
        # 是「停了之後你還要自己處理什麼」，本來就比較適合當結果講，而不是當成
        # 攔在動作前面的問題。
        watching = self.order_exec_watching
        price_busy = self.order_exec_price_busy

        self.order_exec_active = False
        self.order_exec_queue = []
        self.order_exec_pos = 0
        self.order_exec_busy = False
        self.order_exec_watching = False
        self.order_exec_last_note = ""
        self._set_busy(False)
        self._update_order_exec_ui()

        notes = ["下單：已停止。"]
        if watching:
            notes.append("瀏覽器裡那個委託確認視窗程式不再追蹤，請自己按「確認」或「取消」。")
        if price_busy:
            notes.append("背景正在重讀 Excel／跑「更新股價」巨集，沒辦法中斷，"
                         "會等它跑完，但不會再接下一輪。")
        self._say(" ".join(notes))

    def _on_order_filled(self, payload):
        """
        背景回話：這一筆的登入／填單有沒有成功。

        半自動：成功不代表這一輪結束——委託確認視窗還開著，要等
        _on_order_dialog_closed 才會往下一筆走（見本節開頭「依序執行」的
        說明）。自動送出：這時候委託已經真的送出去了（見 order_fill.
        confirm_order），視窗也已經關掉，_on_order_dialog_closed 幾乎會
        立刻跟著發生，不用另外處理。

        maybe_submitted（見 order_fill.OrderMaybeSubmitted）是完全不同等級
        的錯誤：「確認」已經按下去、委託多半已經送出去了，只是沒等到結果，
        絕對不能沿用一般失敗那句「按下一筆會重試同一筆」——重試會把同一筆
        委託再送一次，這裡要用不一樣的標題跟文字整個攔下來。
        """
        self.browser_waiting = max(0, self.browser_waiting - 1)
        self.order_exec_busy = False
        row = self.order_exec_queue[self.order_exec_pos]

        if "error" in payload:
            self._update_order_exec_ui()
            detail = payload["error"][-1500:]
            hint = payload.get("hint")
            text = f"{hint}\n\n────────────────\n{detail}" if hint else detail
            if payload.get("maybe_submitted"):
                show_error(self.root,
                    "委託結果不確定——請先去網站查證",
                    f"{row['sheet']} {row['code']} 這一筆已經按下「確認」，但程式沒辦法"
                    f"確定送出去的結果。\n請先到瀏覽器裡「委託查詢」或「預約查詢」頁自己"
                    f"確認這筆的狀態，確認清楚之前不要按「下一筆」——重試可能會把同一筆"
                    f"委託再送一次。\n\n{text}")
                self._say(f"下單：第 {self.order_exec_pos + 1}/{len(self.order_exec_queue)} 筆"
                          f"結果不確定，先去網站查證，不要按「下一筆」。")
            else:
                show_error(self.root,
                    "這一筆下單失敗",
                    f"{row['sheet']} {row['code']} 這一筆失敗，這一輪先停在這裡。\n"
                    f"排除問題後按「下一筆」會重試同一筆，或按「停止」放棄這一輪。\n\n{text}")
                self._say(f"下單：第 {self.order_exec_pos + 1}/{len(self.order_exec_queue)} 筆失敗，等你處理。")
            return

        auto_result = payload.get("auto_result")
        self.order_exec_last_note = (f"已自動送出，結果：{auto_result['message']}"
                                     if auto_result else "")
        self.order_exec_watching = True
        self._update_order_exec_ui()
        if auto_result:
            self._say(f"下單：{row['sheet']} {row['code']} 已自動送出，結果：{auto_result['message']}")
        else:
            self._say(f"下單：{row['sheet']} {row['code']} 委託確認視窗已開啟，"
                      f"請到瀏覽器裡確認或取消。視窗關閉後「下一筆」才會亮起來。")

    def _on_order_dialog_closed(self, _payload):
        """背景執行緒親眼確認委託確認視窗真的關了（見 ui_background._browser_worker 的閒置輪詢）。"""
        self.order_exec_watching = False
        self.order_exec_pos += 1
        if self.order_exec_pos >= len(self.order_exec_queue):
            self.order_exec_queue = []
            self.order_exec_pos = 0
            self._order_round_finished()
            return
        self._update_order_exec_ui()

    def _order_round_finished(self):
        """
        這一輪的委託全部處理完了。沒勾「多輪直到出清」就跟原本行為一樣直接
        收尾；勾了的話要重讀 Excel（見 _prepare_next_round）才知道還有沒有
        沒出清的部位，不能在這裡就地判斷完成——判斷完成需要的是「這一輪
        委託真的成交之後」的持股，不是下單前那個舊的 self.order_holdings。
        """
        if not self.order_exec_multi_round:
            self.order_exec_active = False
            self._set_busy(False)
            self._say("下單：這一輪已經跑完。")
            self._update_order_exec_ui()
            return

        if self.order_exec_round >= ORDER_MULTI_ROUND_CAP:
            self.order_exec_active = False
            self._set_busy(False)
            self._update_order_exec_ui()
            self._say(f"下單：已經連續跑了 {ORDER_MULTI_ROUND_CAP} 輪，先停在這裡等你確認。")
            show_warning(self.root,
                "已到多輪上限",
                f"「多輪直到出清」已經連續跑了 {ORDER_MULTI_ROUND_CAP} 輪，為了安全先停下來，"
                f"不會無限跑下去。\n請自己檢查目前持股，如果還沒出清，可以再按一次"
                f"「開始下單」重新跑一輪。")
            return

        self._say(f"下單：第 {self.order_exec_round} 輪已跑完，重新讀取持股中…")
        self._prepare_next_round()

    def _prepare_next_round(self):
        """
        背景重讀 Excel（見 _order_price_refresh_worker），依 order_exec_auto_price
        決定要不要先觸發「更新股價」巨集。結果回來由 _on_order_price_refresh
        接手判斷是否已經出清、要不要繼續下一輪。

        第 1 輪也會走這裡（見 start_order_execution）——「每一輪開始前都先
        更新股價，包含第一輪」是規劃文件明講的，第一輪不能因為 queue 已經
        算好了就跳過這一步，兩者共用同一條路，不是各自維護一份類似的邏輯。
        """
        self.order_exec_price_busy = True
        self._apply_busy_state()
        self._update_order_exec_ui()
        sheets = sorted({account["sheet"] for account in self.order_exec_accounts})
        threading.Thread(
            target=self._order_price_refresh_worker,
            args=(self.path, sheets, self.order_exec_auto_price),
            daemon=True,
        ).start()

    def _order_price_refresh_worker(self, path, sheets, run_macro):
        """
        背景執行緒：多輪出清用。run_macro 為真就先觸發使用者既有的「更新
        股價」巨集、等它跑完，再重讀這幾個分頁的持股（E/F 欄）跟股價
        （I 欄）——巨集也好、重讀也好，都跟 refresh_order_data 一樣只在
        COM 層面動，不碰瀏覽器，所以不必透過 browser_cmd_queue，直接開一條
        執行緒做（同 _order_read_worker 的做法）。
        """
        import pythoncom

        pythoncom.CoInitialize()
        excel = workbook = sheet = None
        payload = {}
        try:
            # write=run_macro：巨集要往 I4:I13 寫回股價，開檔如果是唯讀的話
            # 巨集這一步可能直接被 Excel 擋下來或跑到一半出錯。實務上這裡
            # 幾乎一定會走「檔案已經開在使用者的 Excel 裡」那個接上既有
            # 視窗的分支（見 excel_io.open_workbook），write 參數對那個分支
            # 沒有影響，只有在真的還沒開過檔、要另外生一個隱形 Excel 實例
            # 那個少見的情況才有差——但既然要觸發會寫入的巨集，這裡就不要
            # 省這個 True，錯的方向（該可寫卻開成唯讀）比多寫一個參數危險。
            with excel_io.opened(path, run_macro) as (excel, workbook, _attached):
                data, errors = {}, {}
                with excel_io.keep_active_sheet(workbook):
                    for name in sheets:
                        sheet, error = excel_io.find_sheet(workbook, name)
                        if sheet is None:
                            errors[name] = error
                            continue
                        # 一頁一次，理由同 _order_read_worker。
                        if run_macro:
                            excel_io.run_update_price_macro(
                                excel, sheet, on_stuck=self._macro_stuck_notifier("更新股價", name))
                        data[name] = excel_io.read_sheet(sheet)
                # 巨集寫過 I4:I13 就要存檔，理由同 ui_order._order_read_worker。
                if run_macro:
                    workbook.Save()
                payload = {"sheets": data, "errors": errors}
        except Exception as exc:
            payload = {"error": str(exc)}
        finally:
            sheet = excel = workbook = None
            pythoncom.CoUninitialize()
        self.queue.put(("order_price_refresh", payload))

    def _on_order_price_refresh(self, payload):
        """
        _prepare_next_round 的背景讀取回話。重新組一次隊列，還有東西就是
        還沒出清，凍結成下一輪的 queue 繼續跑；空了就是這一批全部出清了。
        """
        self.order_exec_price_busy = False
        self._apply_busy_state()

        if not self.order_exec_active:
            # 讀到一半使用者按了「停止」——這個結果已經跟不上了，不要
            # 回頭把畫面/忙碌狀態又動一次（stop_order_execution 已經處理過）。
            return

        if "error" in payload:
            self.order_exec_active = False
            self._set_busy(False)
            self._update_order_exec_ui()
            show_error(self.root,
                "重讀持股失敗",
                f"第 {self.order_exec_round} 輪跑完後想重新讀取持股，但失敗了，"
                f"「多輪直到出清」先停在這裡：\n\n{payload['error']}")
            self._say("下單：重讀持股失敗，多輪出清已停止。")
            return

        errors = payload["errors"]
        if errors:
            # 部分帳戶讀不到就整批停下來，不要拿舊資料繼續算下一輪——這幾位
            # 沒讀到的話沒辦法判斷他們到底出清了沒，硬跑下去可能漏單也可能
            # 對著已經出清的部位重複下單。
            self.order_exec_active = False
            self._set_busy(False)
            self._update_order_exec_ui()
            show_error(self.root,
                "重讀持股失敗",
                f"這幾個帳戶讀不到，「多輪直到出清」先停在這裡：\n"
                f"{'、'.join(f'{name}：{msg}' for name, msg in errors.items())}")
            self._say("下單：部分帳戶讀不到持股，多輪出清已停止。")
            return

        for name, data in payload["sheets"].items():
            for row in data["rows"]:
                self.order_holdings[(name, row["code"])] = row["qty"]
            # 這個帳戶原本有、這次沒讀到的股票代表已經出清歸零（read_sheet
            # 空白列直接跳過，不會出現在 rows 裡）——不是「讀不到所以不管」，
            # 是這一列真的空了，_order_stock_settings 裡的每一檔都要能反映
            # 這件事，所以先把這個帳戶名下、這一輪處理過的股票代碼清成 0，
            # 再用這次讀到的蓋回去。
            codes_this_round = {s["code"] for s in self.order_exec_stock_settings}
            for code in codes_this_round:
                self.order_holdings.setdefault((name, code), 0)
            for code in codes_this_round - {r["code"] for r in data["rows"]}:
                self.order_holdings[(name, code)] = 0

        # I 欄跟 D 欄同一列對應同一檔股票（見 excel_io.py 開頭說明），這裡把
        # 讀到的股價彙整成 代號->價格，哪個帳戶先讀到就先用哪個——同一檔
        # 股票的市價理論上不會因為帳戶不同而不同，這裡不刻意比對多個帳戶
        # 讀到的數字是否一致，跟現有 fetch 那幾支查詢的態度不同，是因為這裡
        # 本來就只是「參考同一份 Excel 既有機制」，不是本程式自己認定的
        # 權威資料源。整份換掉不是累加——上一輪的舊價格這一輪不該還留著
        # （見 order_exec_prices 開頭的說明）。不管 order_exec_auto_price 這
        # 一輪有沒有先觸發「更新股價」巨集都覆蓋——這裡的 sheets 資料本來就
        # 已經讀進來了，沒理由只挑巨集有跑的那幾輪才拿來用（2026/08/29 使用
        # 者確認：盤中一律用 Excel 成交價，這個開關只決定重讀前要不要先觸發
        # 巨集，不再決定 pricenow 走 Excel 還是現查網頁）。
        # _order_fill_job 盤中算追價時會拿這份當 pricenow，跟 FastQuote 現查
        # 委買賣一一起餵給 chase_price 當邊界。
        price_by_code = {}
        for data in payload["sheets"].values():
            for row in data["rows"]:
                if row["code"] not in price_by_code and row["price"] is not None:
                    price_by_code[row["code"]] = row["price"]
        self.order_exec_prices = price_by_code
        # 第 1 輪凍結的即時委買賣只給第 1 輪用（見 order_exec_quotes 開頭的
        # 說明）——這裡是第 2 輪以後才會走到的路，清空逼這一輪的每一列都
        # 退回「下單前才查」，不把已經過時的報價繼續當最新的用。
        self.order_exec_quotes = {}

        side = self.order_exec_side
        if self.order_exec_mode == "intraday":
            preview = orders.plan_intraday_orders(
                self.order_exec_stock_settings, self.order_exec_accounts,
                self.order_holdings, self.order_exec_ticks, side,
                prices=self.order_exec_prices, quotes=self.order_exec_quotes)
            queue_rows = orders.executable_intraday_orders(preview)
        else:
            preview = orders.plan_stock_orders(
                self.order_exec_stock_settings, self.order_exec_accounts,
                self.order_holdings, side)
            queue_rows = orders.executable_orders(preview)

        self._render_order_preview(preview, [])

        if not queue_rows:
            self.order_exec_active = False
            self._set_busy(False)
            self._update_order_exec_ui()
            # queue 空了不一定是「真的出清歸零」——比重四捨五入到 1 張，
            # 剩不到半張的零股永遠湊不出下一張整張委託，會一直卡在這裡，
            # 不能因為 queue 空了就跟人講「已經出清」，那不是真的（見
            # docs 規劃裡「出清股票(零股)」是另一個還沒做的功能，這裡的
            # 多輪只處理整張）。
            leftover = sum(item["held_qty"] for item in preview if item["held_qty"] > 0)
            if self.order_exec_round == 0:
                # 第一輪重讀完就發現沒東西可送——可能是持股在按下「開始
                # 下單」之後、巨集跑完之前這個空檔剛好變了，一輪都還沒真的
                # 跑，跟「跑了幾輪之後出清」是不同的事，訊息不能講「跑了
                # 0 輪」，那不是人話。
                self._say("下單：重讀 Excel 之後，沒有需要處理的委託了（可能持股剛好在這個空檔變了）。")
            elif leftover > 0:
                self._say(f"下單：跑了 {self.order_exec_round} 輪，整張的部分已經出清，"
                          f"但還剩下不到 1 張的零股（比重換算不出下一張整張委託），"
                          f"多輪出清到這裡為止，零股要另外處理。")
            else:
                self._say(f"下單：跑了 {self.order_exec_round} 輪，已經全部出清。")
            return

        self.order_exec_round += 1
        self.order_exec_queue = queue_rows
        self.order_exec_pos = 0
        self._say(f"下單：開始第 {self.order_exec_round} 輪…")
        self._update_order_exec_ui()
        self._dispatch_next_order()

    def _update_order_exec_ui(self):
        if self.order_exec_price_busy:
            # 多輪出清輪與輪之間：queue 是空的，但整批作業還在跑（見
            # order_exec_active 的說明），「停止」要維持可以按。
            self.order_exec_button.configure(text="處理中…", state="disabled")
            self.order_exec_stop_button.configure(state="normal")
            self.order_exec_status.configure(
                text=f"第 {self.order_exec_round} 輪已跑完，重新讀取持股中…")
            return

        total = len(self.order_exec_queue)
        if total == 0:
            # 按鈕上的字跟著作業／單位／時機走（見 ui_order._order_exec_label）；
            # 還沒接上的作業一律灰掉，為什麼灰的由正上方那行預覽說明講（見
            # ui_order._recompute_order_preview），這裡不重複講第二次。
            ready = self._order_job_ready()
            self.order_exec_button.configure(
                text=self._order_exec_label(),
                state="disabled" if self.busy or not ready else "normal")
            self.order_exec_stop_button.configure(state="disabled")
            # 灰掉的第二個理由要自己講。作業還沒接上的那個理由由正上方預覽區
            # 那行說明講（見 ui_order._recompute_order_preview），但「背景有工作
            # 在跑」原本沒有任何地方講——按鈕就這樣灰在那裡，看起來像壞了。
            # self.busy 是登入／登出／讀取／寫入／查掛單／查委買賣共用的那一顆
            # （見 ui_background._set_busy），這幾條路跑完它就會自己亮回來。
            self.order_exec_status.configure(
                text="背景有工作在跑（登入／登出／讀取／寫入／查詢），跑完這顆才會亮。"
                     if self.busy and ready else "")
            return

        self.order_exec_stop_button.configure(state="normal")
        pos = self.order_exec_pos
        row = self.order_exec_queue[pos]
        label = f"{row['sheet']}　{row['code']} {row['name']}"
        # 多輪出清才標「第 N 輪」——單輪模式一直都只有一輪，標了反而多餘。
        prefix = f"第 {self.order_exec_round} 輪 " if self.order_exec_multi_round else ""
        if self.order_exec_watching:
            self.order_exec_button.configure(text="下一筆", state="disabled")
            note = self.order_exec_last_note or (
                "委託確認視窗已開啟，去瀏覽器確認或取消，視窗關閉後才能按「下一筆」。")
            self.order_exec_status.configure(text=f"{prefix}({pos + 1}/{total}) {label}：{note}")
        elif self.order_exec_busy:
            self.order_exec_button.configure(text="處理中…", state="disabled")
            self.order_exec_status.configure(text=f"{prefix}({pos + 1}/{total}) {label}：登入／填單中…")
        else:
            # 不在等視窗、也不在忙，卡在這裡代表上一筆失敗了（剛開始執行的當下
            # 一定馬上進 order_exec_busy，不會停留在這個分支，見 start_order_execution）。
            self.order_exec_button.configure(text=f"下一筆（{pos + 1}/{total}）", state="normal")
            self.order_exec_status.configure(text=f"{prefix}({pos + 1}/{total}) {label}：上一筆失敗，可以重試或停止。")

    def _order_fill_job(self, context, store, order_number, account, row, mode, ticks, auto, side):
        """
        背景執行緒用（只能在 ui_background._browser_worker 裡呼叫）：登入或換
        cookie 到 row 對應的帳戶、開下單頁、選股票、填單、開出委託確認視窗；
        auto 為真的話再多按下視窗裡的「確認」，真的把委託送出去（見
        order_fill.confirm_order）。

        回傳 (page, extra)：page 只能留在呼叫端（_browser_worker）手上、記著
        要盯哪一頁的確認視窗，絕對不能塞進 self.queue 送回主執行緒——
        Playwright 的物件被綁死在建立它的那個執行緒，主執行緒去碰會直接壞掉
        （見 ui_background.py 開頭的執行緒說明）。extra 是給主執行緒看的純
        資料，auto 送出成功時帶 "auto_result"，半自動或還沒送出前失敗時是
        空字典。

        半自動（auto=False）不按確認視窗裡的「確認」，理由跟 order_fill.
        fill_order 一樣：那一步留給人。

        mode=="intraday" 時 row["price"] 通常是 None（見 orders.plan_intraday_orders），
        要在這裡算出真正要送出的價格——除非按過「查詢委買賣」把這一檔的即時
        報價整批先查回來過（見 fetch_order_quotes），那樣 row["price"] 在
        建執行清單那一刻就已經是算好的數字，這裡直接照用，不再重查一次
        （2026/08/29 使用者要求：按下「開始下單」就是執行預覽上看到的價位，
        不要下單前又變成另一個數字）。

        pricenow（chase_price 的基準價）一律來自 self.order_exec_prices——
        Excel I 欄讀回來的成交價，不現查網頁（2026/08/29 使用者確認：沒必要
        等下單前才另外查一次）。第 1 輪這份值是開始下單那一刻拿 self.
        order_prices 當快照（見 start_order_execution），也就是上一次
        「讀取持股」讀到的數字——盤中模式的這顆按鈕會先觸發「更新股價」
        巨集才讀（見 refresh_order_data），所以只要開始下單前有按過一次
        「讀取持股」，這個基準價就是新的，不是放到過期的舊資料。
        order_exec_auto_price 這個開關現在只決定多輪出清時，輪與輪之間重讀
        Excel 前要不要先觸發巨集，不再決定 pricenow 的資料來源。
        """
        page, _, problems = fetch_mod.ensure_logged_in(
            context, [(order_number, account)], store)[order_number]
        if problems:
            raise RuntimeError("；".join(problems))

        # 方向與委託別一律看**那一列自己帶的值**，不是這一輪整批共用的那個：
        # 買賣股票的方向是逐檔逐帳戶由下單試算的正負算出來的（見
        # orders.plan_trade_orders），出清那兩支 plan_* 也一樣把 side 寫進每一列，
        # 所以兩邊共用同一條路，不必在這裡分作業。
        row_side = row.get("side") or side
        price = row["price"]
        if mode == "intraday" and price is None:
            pricenow = self.order_exec_prices.get(row["code"])
            if pricenow is None:
                raise RuntimeError(
                    f"沒有讀到 {row['code']} 的股價（Excel I 欄），這一筆沒辦法算追價。"
                    f"請先按「讀取持股」讓 Excel 更新股價。")

            # 對手方第一檔：借這個已登入的 page 開一個 FastQuote 彈出視窗，
            # 只為了這一檔股票訂閱、等一下、拿到就關掉（見 fastquote.py
            # 「另開分頁不是一律不行」）——刻意保持簡單，不常駐訂閱、不接
            # ui_background.py 的即時表格，每筆單各自開各自關。查哪一檔跟
            # side 是反的：買方向查委賣一（ask）、賣方向查委買一（bid），見
            # orders.chase_price 的說明。收不到（逾時、或 WebSocket 這輪
            # 剛好不穩）就讓 best_opposite 維持 None，chase_price 自己會退回
            # 邊界價，不擋單。
            best_opposite = None
            stream = fastquote.FastQuoteStream(page)
            try:
                if stream.subscribe([row["code"]]):
                    quote = stream.wait_for(row["code"])
                    if quote:
                        best_opposite = quote["ask"] if row_side == orders.SIDE_BUY else quote["bid"]
            finally:
                stream.close()

            price = orders.chase_price(pricenow, ticks, row_side, best_opposite)

        # 委託別跟著模式走，不是兩邊共用同一個值（見 orders.BS_FLAG_PRE 的
        # 說明）：盤前開盤前還沒有連續交易，只能用 ROD；盤中規劃文件明講
        # 用 IOC。2026/08/28 使用者更正過，之前這裡兩種模式都寫死 IOC 是錯的。
        # 買賣股票那幾列自己帶著 bs_flag（規劃文件明講用 ROD），沒帶的才照模式
        # 決定：盤前只能 ROD，盤中規劃文件明講 IOC。
        bs_flag = row.get("bs_flag") or (
            orders.BS_FLAG_INTRADAY if mode == "intraday" else orders.BS_FLAG_PRE)

        # 整張還是零股也看**那一列自己帶的值**（理由同上面的 side／bs_flag）：
        # 這台引擎吃的是凍結好的 queue，不回頭問畫面上現在選的是哪一個——多輪
        # 之間、或人在執行中動了那顆單選鈕，畫面上的值跟這批 queue 算出來的量
        # 就對不起來了。沒帶這一欄的（出清那兩支 plan_*）當整張。
        # row["lots"] 的單位跟著它走：整張是張、零股是股（見
        # orders.plan_trade_orders），兩支 order_fill 都要帶同一個 odd。
        odd = row.get("unit") == orders.UNIT_ODD

        page.goto(order_fill.ORDER_ENTRY_PAGE, wait_until="domcontentloaded")
        order_fill.open_order_form(page, odd=odd)
        order_fill.select_stock(page, row["code"])
        order_fill.fill_order(page, side=row_side, qty=row["lots"], price=price,
                              bs_flag=bs_flag, odd=odd)

        if not auto:
            return page, {}

        ok, message = order_fill.confirm_order(page)
        return page, {"auto_result": {"ok": ok, "message": message}}

    def _order_dialog_closed(self, page):
        """
        委託確認視窗還開著沒？只能在背景瀏覽器執行緒裡呼叫（同上，Playwright
        物件的執行緒限制）。

        頁面／瀏覽器已經不在了（使用者自己把分頁或整個瀏覽器關掉）就當作「已經
        關閉」——不能讓一個已經摸不到的頁面卡住整個「下一筆」流程，最壞的後果
        頂多是使用者接下來自己再重登一次，比程式在這裡卡死好處理。
        """
        try:
            return not page.locator(".layui-layer-title", has_text="委託確認").first.is_visible()
        except PlaywrightError:
            return True
