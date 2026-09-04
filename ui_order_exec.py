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

一輪的委託送完之後，最多還有三段（都在這個檔案裡，一段接一段）：

    撤零股單  只有出清零股會走。零股沒有 IOC 可用，掛出去不會自己取消，所以
              等 6 秒再把沒成交的撤掉（規劃文件流程第 2 步）。一個帳戶一則
              指令，見 _dispatch_next_odd_cancel。
    完整同步  只有勾了多輪才走。查網頁 → 把現金、股數、成本寫回持股管理檔 →
              落帳，走的是更新分頁那一整條路，不是另外寫一份簡化版
              （見 _start_round_sync）。
    重讀判斷  重讀一次 Excel，還組得出隊列就是還沒出清，接下一輪
              （_prepare_next_round／_on_order_price_refresh）。

這三段中間 queue 都是空的，但整批作業還在跑——「停止」在每一段都要按得下去，
靠的是 order_exec_active 而不是 queue 有沒有東西。
"""

import datetime
import threading
import tkinter as tk

from playwright.sync_api import Error as PlaywrightError

import excel_io
import fetch as fetch_mod
import order_fill
import order_query
import orders
import planner
import stockinfo
from ui_common import ask_confirm, show_error, show_info, show_warning
from util import show

# 「多輪直到出清」的安全上限：不管有沒有真的出清，跑滿這個輪數就一定停下來
# 等人看過再決定要不要繼續，不無限跑下去——2026/08/28 使用者確認要做這個
# 開關的時候沒特別要求上限數字，這裡抓一個「明顯比一般情境需要的輪數多，
# 但出問題也不會跑太久」的值，不是照文件或使用者指定的數字。
ORDER_MULTI_ROUND_CAP = 10

# 出清零股：全部掛完賣單之後等多久才回頭撤單。
#
# 規劃文件「出清股票－零股」流程第 2 步寫的是 20 秒，2026/09/04 使用者改成
# **6 秒**：台股盤中零股是**每 5 秒集合競價撮合一次**，6 秒保證涵蓋至少一次
# 撮合，多的 1 秒是緩衝。20 秒等於白等 3 次撮合週期——會成交的第一次撮合就
# 成交了，不會的等再久也不會。
#
# 代價是每一輪只有一次撮合機會，沒撮到就撤掉：單輪模式下這一輪就結束了，
# 勾了多輪則是下一輪重新查價再掛一次（而且價格是新的，比在市場上放著等更
# 貼近當下）。要改回長一點的話連同這段理由一起改，不要只動數字。
#
# 為什麼零股要「掛了再撤」而不是像整張那樣用 IOC：零股那一場根本沒有 IOC
# 可選（見 orders.BS_FLAG_ODD），只能用 ROD 掛出去，掛著不撤就會留在市場上
# 過夜。
ODD_CANCEL_WAIT_MS = 6000


class UiOrderExecMixin:
    # ---------- 依序執行（半自動下單／自動送出） ----------
    #
    # 跟 order_fill.py 同一個界線：程式只填到「開出委託確認視窗」為止，按不
    # 按裡面的「確認」由 order_exec_auto（畫面上「自動送出」開關）決定。這裡
    # 多做的是把執行預覽裡一整批委託接起來，一筆一筆換帳戶（換 cookie，不
    # 重登，見 fetch.new_store）自動填單——但兩筆之間一定要等上一筆的確認
    # 視窗真的處理掉（按確認或取消，不管是人按的還是程式自己按的）才能換下
    # 一個帳戶的 cookie，不能提早接：整個瀏覽器只有一組 cookie，要是視窗還
    # 開著就換了下一個帳戶的 cookie，之後才回去按那個視窗的「確認」，送出的
    # 會是「現在瀏覽器帶著的身分」而不是視窗原本對應的那個帳戶——這是真的會
    # 送錯帳戶的委託，不是好看不好看的問題。往下一筆走（不管是自動送出時
    # 程式自己接，還是半自動時解鎖「下一筆」按鈕給人按）只在背景執行緒親眼
    # 確認視窗真的關了（_order_dialog_closed）之後才會發生，就是為了擋住
    # 這條路。
    #
    # 自動送出開著時，視窗一關（程式自己按的確認，理論上馬上就會關）就直接
    # 接下一筆，不再多等人按「下一筆」（2026/09/02 使用者要求，推翻先前
    # 「不管開不開，節奏都不變」那個決定）；半自動（人在瀏覽器裡自己按確認
    # 或取消）維持原樣，還是要人按過「下一筆」才會送下一筆——那是半自動流程
    # 裡使用者唯一的節奏控制點。
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
        # 成交價已經在新增股票／讀取試算時讀進 Excel，不必等下單前
        # 才查）。開始下單那一刻先用 self.order_prices（見那裡的說明）當第 1
        # 輪的起始值（start_order_execution），之後每輪重讀就整份換掉，不是
        # 累加（同一檔股票這一輪的價格只有一個版本）——order_exec_auto_price
        # 這個開關現在只決定重讀前要不要先觸發「更新股價」巨集，不再決定
        # pricenow 走哪條路（見 _on_order_price_refresh）。
        self.order_exec_prices = {}
        # 開始下單那一刻凍結的 self.order_quotes 快照。_on_order_price_refresh
        # 會把它清成空字典，逼所有列都退回「下單前才查」那條舊路，不讓舊報價被
        # 沿用到之後幾輪（市場已經過了一段時間，繼續當最新報價用是在猜數字）。
        #
        # **所以「這份快照給第 1 輪用」只有沒勾「自動更新股價」時才成立**：勾了的話
        # 第 1 輪也會先經過 _on_order_price_refresh（start_order_execution 設
        # round=0 就呼叫 _prepare_next_round），這份快照在送出任何一筆之前就被清掉
        # 了，第 1 輪也是送單前重查。結果是對的（股價跟委買賣一都是新的），但執行
        # 預覽那句「委買一 X 價送出」在那個組合下語意不正確——見
        # _on_order_price_refresh 清空那一行旁邊的說明。
        self.order_exec_quotes = {}
        self.order_exec_price_busy = False  # 輪與輪之間正在重讀 Excel／觸發巨集，還沒回話
        # 「這一整批多輪出清作業還在不在跑」，跟 order_exec_queue 分開——
        # 輪與輪之間 queue 會是空的（下一輪還沒組出來），但整批作業並沒有
        # 結束，「停止」要能在這個空檔也按得下去（見 stop_order_execution）。
        self.order_exec_active = False

        # 這一輪凍結的「作業」（orders.JOB_*）。之前不必存是因為只有出清整張
        # 一種行為接上了；零股接上之後，輪與輪之間重組隊列要知道該叫哪一支
        # plan_*（見 _on_order_price_refresh），而畫面上的 self.order_job 隨時
        # 可能被使用者切走——凍結的理由跟 order_exec_mode 那幾個一模一樣。
        self.order_exec_job = orders.JOB_CLEAR
        self.order_exec_unit = orders.UNIT_LOT

        # 出清零股「掛完之後隔幾秒取消全部零股單」那一段（規劃文件流程第 2 步）。
        # 等待用 root.after 排一次，id 存起來是為了「停止」按得掉——不存的話
        # 那顆計時器還是會在時間到的時候醒來，對著一個已經停掉的作業派出取消指令。
        self.order_exec_cancel_timer = None
        self.order_exec_cancel_queue = []   # [(第幾組, 帳號設定, 分頁名, 這一輪的股票代號)]
        self.order_exec_cancel_pos = 0
        self.order_exec_cancel_results = []  # 每個帳戶回報的取消結果，跑完統計一次
        self.order_exec_cancel_problems = []
        # 撤單那一段的結論，留給 _after_round_actions 接著講（見那裡的說明）。
        self.order_exec_cancel_note = ""

        # 多輪之間那一次「更新持股管理檔的現金、股數、成本」（規劃文件流程第
        # 3 步）還沒回話。走的是更新分頁原本那一整條路（fetch → planner →
        # 寫 Excel → 落帳），不是另外寫一份，所以這裡只留一個「在等它」的旗標，
        # 真正的收尾在 _order_sync_finished（由 ui_background 的 _on_fetched／
        # _on_written 末端呼叫）。
        self.order_exec_sync_busy = False
        # 借用更新分頁 cash_method 那顆 Var 算完這一輪要換回來的原值（見
        # _start_round_sync／_order_sync_finished）。None＝目前沒有借用中。
        self._round_sync_prev_method = None
        # 上一輪委託佇列的指紋（見 _queue_signature）。多輪的「沒有進展就停」
        # 那道保險靠它，None＝這一批還沒有跑過任何一輪。
        self.order_exec_last_signature = None

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
        # 多輪要用到的東西（Excel、紀錄檔、今天的現金算法）在這裡一次問完，
        # 不留到第一輪跑完才問（見 _order_round_sync_ready）。
        if multi_round and not self._order_round_sync_ready():
            return
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
        elif self._order_clear_odd():
            # 出清零股：量是持股的零股那一段（沒有比重），委託別固定 ROD，
            # 價格跟盤中出清同一條追價路（見 orders.plan_clear_odd_orders）。
            stock_settings = self._order_stock_settings()
            ticks = self._order_ticks_setting()
            if ticks is None:
                show_error(self.root, "追價檔數不對", "追價檔數要填 0 以上的整數。")
                return
            preview = orders.plan_clear_odd_orders(
                stock_settings, ordered, self.order_holdings, ticks,
                prices=self.order_prices, quotes=self.order_quotes)
            queue_rows = orders.executable_intraday_orders(preview)
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
            elif self._order_clear_odd():
                reason = "勾選的帳戶都沒有這幾檔，或持股剛好是整張、沒有零股可以出清"
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

        # 標題行三個作業統一格式：「{作業} {單位}{ 盤前／盤中}（委託別）」，四種
        # 情境（買賣股票；出清・整張・盤前／盤中；出清・零股）共用同一套排法，
        # 2026/09/03 使用者要求別讓人自己比對四句長得不一樣的話（見同一天
        # ui_order.py 那則 KeyError 修復旁邊的討論）。
        unit_word = orders.UNIT_NAMES[self.order_unit.get()]
        if job == orders.JOB_TRADE:
            # 買賣股票的方向是逐筆的（試算正數買、負數賣），不能像出清那樣用一句
            # 「即將賣出 N 筆」帶過——那會讓人以為整批同一個方向。買賣確實可能
            # 混在同一輪裡（不同股票試算正負不同），所以買、賣兩個數字都要列。
            buys = sum(1 for row in queue_rows if row["side"] == orders.SIDE_BUY)
            head = (
                f"買賣股票 {unit_word}（ROD-當日有效）\n"
                f"即將依序處理 {len(queue_rows)} 筆委託\n"
                f"買 {buys} 筆、賣 {len(queue_rows) - buys} 筆\n\n"
            )
        elif self._order_clear_odd():
            # 零股跟整張最大的差別要寫在最前面：它不是 IOC，是「掛了再撤」，
            # 而那個撤是程式幾秒後自己做的（規劃文件流程第 1、2 步）。人按下
            # 去之前要知道「這批單會真的掛在市場上一段時間」。追價價格是凍結還是
            # 即時算那段說明 2026/09/03 拿掉了（使用者指定），跟取消掛單確認視窗
            # 同一個簡化方向。
            seconds = ODD_CANCEL_WAIT_MS // 1000
            head = (
                f"出清股票 零股 盤中（ROD-當日有效）\n"
                f"即將依序處理 {len(queue_rows)} 筆零股「賣出」委託。\n\n"
                f"全部掛完之後會等 {seconds} 秒，再取消掛單。\n\n"
            )
        elif self._order_intraday():
            head = (
                f"出清股票 整張 盤中（IOC-立即成交否則取消）\n"
                f"即將依序處理 {len(queue_rows)} 筆「{side_word}」委託。\n\n"
            )
        else:
            head = (
                f"出清股票 整張 盤前（ROD-當日有效）\n"
                f"即將依序處理 {len(queue_rows)} 筆「{side_word}」委託。\n\n"
            )

        # 勾了什麼就列一行，沒勾的完全不出現——不解釋那個選項會做什麼、也不解釋
        # 沒勾會怎樣（2026/09/04 使用者定稿）。這兩個開關的實際行為（多輪之間跑
        # 的是完整同步不是只重讀 Excel、跑滿 ORDER_MULTI_ROUND_CAP 輪一定停、
        # auto_price 只決定重讀前要不要先觸發巨集）都還在，只是不寫進確認框：
        # 那幾句是「按下去之後會發生什麼」的說明書，不是這一刻要做的決定。行為
        # 本身見 _start_round_sync／_on_order_price_refresh 跟 docs/介面規劃.md 9.9。
        if multi_round:
            head += "已勾選「多輪直到出清」\n"
            if auto_price:
                head += "已勾選「自動更新股價」\n"
            head += "\n"

        if auto:
            tail = "「自動送出委託單」模式，確定嗎？"
        else:
            tail = "「手動送出委託單」模式，確定嗎？"
        if not ask_confirm(self.root, "開始下單", head + tail, confirm_style="primary",
                           emphasize="自動送出委託單" if auto else None):
            return

        self.order_exec_mode = mode
        self.order_exec_ticks = ticks
        self.order_exec_auto = auto
        self.order_exec_side = side
        # 作業與單位也要凍結：輪與輪之間重組隊列時要知道叫哪一支 plan_*（見
        # _on_order_price_refresh），而畫面上那兩個變數隨時可能被切走。
        self.order_exec_job = job
        self.order_exec_unit = self.order_unit.get()
        self.order_exec_multi_round = multi_round
        self.order_exec_auto_price = auto_price
        self.order_exec_stock_settings = stock_settings
        self.order_exec_accounts = ordered
        # 第 1 輪（沒勾自動更新股價時就是唯一一輪）直接拿 self.order_prices
        # 當起點——那是新增股票／上次「讀取試算」讀進來的 Excel
        # 成交價，不用再另外查一次。勾了自動更新股價的話，這份值一送進
        # _prepare_next_round 馬上就會被剛重讀（含觸發巨集）的結果整份蓋掉
        # （見 _on_order_price_refresh），不是兩份資料混用。
        self.order_exec_prices = dict(self.order_prices)
        # 凍結這一刻查到的即時委買賣（見 order_exec_quotes 開頭的說明）。真的用得
        # 到它的只有「沒勾自動更新股價」那條路——勾了的話下面 _prepare_next_round
        # 進去就被 _on_order_price_refresh 清空，第 1 輪也是送單前重查。
        self.order_exec_quotes = dict(self.order_quotes)
        # 新的一批從頭開始比對「有沒有進展」，不能沿用上一批留下來的指紋。
        self.order_exec_last_signature = None
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
            # 沒勾自動更新股價時第 1 輪不走 _prepare_next_round，指紋要在這裡自己
            # 記一次，否則第 2 輪沒有東西可以比、那道保險等於從第 3 輪才開始生效。
            self.order_exec_last_signature = self._queue_signature(queue_rows)
            self._dispatch_next_order()

    def _order_round_sync_ready(self):
        """
        勾了「多輪直到出清」才要做的檢查，全部在按下「開始」之前做完。

        多輪的每一輪之間會做一次完整同步（規劃文件流程：更新持股管理檔的現金、
        股數、成本），走的是更新分頁那一整條路（見 _start_round_sync）——它要
        Excel 開著、要紀錄檔是好的。

        **現金算法固定用「初始餘額累加」，不問、不跳視窗**（2026/09/03 使用者
        訂正）：規劃文件原文「更新持股管理檔的現金（初始餘額累加）」寫的就是
        這個算法本身，不是「跳視窗讓人選」。跟更新分頁的全域開關（20 人共用、
        今天要不要用銀行餘額推算）是兩件事——出清這幾位當下用哪一種不受那顆
        開關影響，也不會因為先跑了出清就把「今天問過了」那個旗標點掉，更新
        分頁該問的還是照問（見 _start_round_sync 怎麼暫時切換 cash_method）。

        **一定要在整批開始之前檢查完，不能等第一輪跑完才發現 Excel 沒開好**：
        那時候委託已經送出去了。跟 ui_background.start_fetch 在派工之前先確認
        是同一個道理。
        """
        if not self._require_excel():
            return False
        if self.ledger_error:
            show_error(self.root, "紀錄檔有問題", self.ledger_error)
            return False
        self.today = datetime.date.today()
        return True

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
        sync_busy = self.order_exec_sync_busy
        # 撤單計時器一定要收掉：它不像背景執行緒那樣「醒來發現作業停了就
        # 什麼都不做」，它醒來就會真的派出一整輪撤單指令。
        pending_cancel = self.order_exec_cancel_timer is not None
        if pending_cancel:
            self.root.after_cancel(self.order_exec_cancel_timer)
            self.order_exec_cancel_timer = None

        self.order_exec_active = False
        self.order_exec_queue = []
        self.order_exec_pos = 0
        self.order_exec_busy = False
        self.order_exec_watching = False
        self.order_exec_last_note = ""
        self.order_exec_cancel_queue = []
        self.order_exec_cancel_pos = 0
        self.order_exec_cancel_note = ""
        self._set_busy(False)
        self._update_order_exec_ui()

        notes = ["下單：已停止。"]
        if watching:
            notes.append("瀏覽器裡那個委託確認視窗程式不再追蹤，請自己按「確認」或「取消」。")
        if pending_cancel:
            # 這是停止零股出清最要緊的一句：委託已經掛在市場上了，而那個「幾秒
            # 後自動撤掉」的承諾剛剛被取消掉。不講的話，人會以為停止＝什麼都沒
            # 發生，那批 ROD 賣單就這樣留到收盤。
            notes.append("零股賣單還掛在市場上，之後的自動撤單已經取消——"
                         "請到「掛單」分頁自己取消。")
        if price_busy:
            notes.append("背景正在重讀 Excel／跑「更新股價」巨集，沒辦法中斷，"
                         "會等它跑完，但不會再接下一輪。")
        if sync_busy:
            notes.append("背景正在更新持股管理檔，沒辦法中斷，會照常寫完並落帳，"
                         "但不會再接下一輪。")
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
        """
        背景執行緒親眼確認委託確認視窗真的關了（見 ui_background._browser_worker
        的閒置輪詢）。

        自動送出（order_exec_auto）：程式自己按過確認、視窗也已經自己關了，
        不必再讓人多按一次「下一筆」——直接接下一筆（2026/09/02 使用者
        推翻先前「不管開不開，節奏都不變」那個決定，見 ui_order.py
        order_auto_confirm 開頭的說明）。半自動（人在瀏覽器裡自己按確認／
        取消）維持原樣，「下一筆」還是要人按過才會送下一筆——那是使用者
        在畫面上唯一的節奏控制點，這條路沒有要動。
        """
        self.order_exec_watching = False
        self.order_exec_pos += 1
        if self.order_exec_pos >= len(self.order_exec_queue):
            self.order_exec_queue = []
            self.order_exec_pos = 0
            self._order_round_finished()
            return
        if self.order_exec_auto:
            self._dispatch_next_order()
            return
        self._update_order_exec_ui()

    def _order_round_finished(self):
        """
        這一輪的委託全部處理完了。接下來還有兩段，順序不能換：

        1. **出清零股要先撤單**（規劃文件「出清股票－零股」流程第 2 步）：零股
           只能用 ROD（見 orders.BS_FLAG_ODD），掛出去不會自己取消，所以等幾秒
           再把沒成交的撤掉。這一段跟有沒有勾多輪**無關**——單輪也要撤，不然那
           批單就一直留在市場上了。
        2. 勾了多輪才有的收尾：同步、判斷出清了沒、決定要不要再跑一輪（見
           _after_round_actions）。

        判斷「出清了沒」需要的是這一輪委託真的成交**之後**的持股，不是下單前那
        份舊的 self.order_holdings，所以不能在這裡就地判斷完成。
        """
        if not self.order_exec_active:
            # 使用者已經按過「停止」，只是最後那則回話晚一步才到（委託確認視窗
            # 的關閉偵測是背景輪詢的，見 ui_background._browser_worker）。這裡要
            # 是照常往下走，就會替一個已經停掉的作業排出撤單計時器——
            # 「停止」擋的正是這件事。
            return
        if self.order_exec_job == orders.JOB_CLEAR and self.order_exec_unit == orders.UNIT_ODD:
            self._start_odd_cancel_wait()
            return
        self._after_round_actions()

    # ---------- 出清零股：等幾秒之後撤掉沒成交的那幾筆 ----------

    def _start_odd_cancel_wait(self):
        """
        零股全部掛完了，等 ODD_CANCEL_WAIT_MS 再撤（規劃文件流程第 2 步）。

        用 root.after 而不是背景執行緒 sleep：這段等待裡主執行緒要照常回應（尤其
        是底部那顆「停止全部操作」），而醒來之後要做的事（派指令、改狀態列）本來
        就只能在主執行緒上做。計時器 id 存起來是為了停止時取消得掉——不存的話它
        還是會在時間到的時候醒來，對著一個已經停掉的作業派出一整輪撤單。
        """
        seconds = ODD_CANCEL_WAIT_MS // 1000
        self._say(f"下單：零股已經全部掛出去，{seconds} 秒後自動取消沒成交的部分…")
        self._update_order_exec_ui()
        self.order_exec_cancel_timer = self.root.after(
            ODD_CANCEL_WAIT_MS, self._start_odd_cancel)

    def _start_odd_cancel(self):
        """時間到了：把這一輪掛出去的零股賣單整批撤掉，一個帳戶一則指令。"""
        self.order_exec_cancel_timer = None
        if not self.order_exec_active:
            return

        codes = sorted({stock["code"] for stock in self.order_exec_stock_settings})
        queue = []
        for account in self.order_exec_accounts:
            sheet = account["sheet"]
            order_number = self._order_number_for_sheet(sheet)
            if order_number is None:
                # 送單那一步就會先擋下來（見 _dispatch_next_order），真的走到這裡
                # 代表這一位這一輪一筆都沒送出去，沒有東西要撤。
                continue
            queue.append((order_number, self.accounts[order_number - 1], sheet, codes))

        self.order_exec_cancel_queue = queue
        self.order_exec_cancel_pos = 0
        self.order_exec_cancel_results = []
        self.order_exec_cancel_problems = []
        self.order_exec_cancel_note = ""
        self._dispatch_next_odd_cancel()

    def _dispatch_next_odd_cancel(self):
        """
        派下一個帳戶。**一個帳戶一則指令**，跟掛單分頁那三顆取消按鈕同一個形狀
        （見 docs/介面規劃.md 10.3 第六點）：這樣「停止」在帳戶與帳戶之間就是免費
        的——不派下一則就是停了，不必為了中斷另外發明一個跨執行緒的旗標。
        """
        if (not self.order_exec_active
                or self.order_exec_cancel_pos >= len(self.order_exec_cancel_queue)):
            self._odd_cancel_finished()
            return

        order_number, account, sheet, codes = \
            self.order_exec_cancel_queue[self.order_exec_cancel_pos]
        self._say(f"下單：取消零股單，第 {self.order_exec_cancel_pos + 1}/"
                  f"{len(self.order_exec_cancel_queue)} 個帳戶（{sheet}）…")
        self._update_order_exec_ui()
        self._ensure_browser_thread()
        self.browser_waiting += 1
        self.browser_cmd_queue.put(
            ("order_odd_cancel", (order_number, account, sheet, codes)))

    def _on_order_odd_cancelled(self, payload):
        """一個帳戶撤完了（或失敗了）。"""
        self.browser_waiting = max(0, self.browser_waiting - 1)
        sheet = payload.get("sheet", "")

        if payload.get("maybe_submitted"):
            # 「確認」已經按下去了，那一批刪單多半已經送到券商——跟
            # ui_pending._on_pending_cancelled 同一條規矩（10.3 第九點）：整批停
            # 在這裡，不再往下派，讓人自己去掛單分頁看實際狀態。繼續撤別的帳戶
            # 只會讓「現在到底發生了什麼」更難回答。
            self.order_exec_cancel_problems.append(f"{sheet}：{payload.get('error', '')}")
            self.order_exec_active = False
            self._odd_cancel_finished(maybe_submitted=True)
            return

        if "error" in payload:
            self.order_exec_cancel_problems.append(f"{sheet}：{payload['error']}")
        else:
            self.order_exec_cancel_results.append(payload)

        self.order_exec_cancel_pos += 1
        self._dispatch_next_odd_cancel()

    def _odd_cancel_finished(self, maybe_submitted=False):
        """
        撤單這一段跑完了。講一次結果，然後接多輪那一段。

        撤單失敗**不會**讓整批停下來（maybe_submitted 那一種除外，那個在
        _on_order_odd_cancelled 就已經把 active 關掉了）：沒撤掉的單留在市場上是
        要人知道的事，但它跟「下一輪該不該跑」是兩件事——而且下一輪開始前那次
        完整同步會把真實持股讀回來，多掛著的那幾筆造成的差異會反映在下一輪的量
        上，不是拿舊資料硬跑。
        """
        done = sum(1 for item in self.order_exec_cancel_results
                   for result in item.get("results", []) if result["ok"])
        failed = sum(1 for item in self.order_exec_cancel_results
                     for result in item.get("results", []) if not result["ok"])
        # 「查到幾筆還掛著」跟「撤掉幾筆」要分開講：查到 0 筆代表這一輪掛出去的
        # 全部成交了（出清零股最想看到的結果），跟「撤了 0 筆但其實有 5 筆撤失敗」
        # 是完全相反的兩件事，寫成同一句話會讓人看不出差別。
        found = sum(item.get("found", 0) for item in self.order_exec_cancel_results)
        problems = self.order_exec_cancel_problems

        if found == 0 and not problems:
            summary = "這一輪掛出去的零股都成交了，沒有需要取消的單。"
        else:
            parts = [f"取消零股單完成，查到 {found} 筆還掛著，取消 {done} 筆"]
            if failed:
                parts.append(f"、{failed} 筆取消不了（去掛單分頁看原因）")
            parts.append("。")
            summary = "".join(parts)
        if problems:
            summary += f"　有 {len(problems)} 個帳戶沒取消成功。"
        # 存起來給 _after_round_actions 接著講：單輪的話它下一句就是「這一輪已經
        # 跑完」，直接 _say 會把這段結論整句蓋掉——而這是唯一講得出「掛出去的單
        # 後來怎麼了」的地方，蓋掉等於沒撤單報告可看。
        self.order_exec_cancel_note = summary
        self._say(f"下單：{summary}")

        if problems:
            show_warning(self.root,
                "取消零股單沒有全部成功",
                "這幾個帳戶的零股單沒有取消，請自己到「掛單」分頁查一次、"
                "必要時手動取消：\n\n" + "\n\n".join(problems[:5]))

        if maybe_submitted:
            # active 已經被關掉了，這裡只負責把畫面收乾淨。
            self._set_busy(False)
            self._update_order_exec_ui()
            return
        self._after_round_actions()

    # ---------- 一輪結束之後：同步、判斷出清了沒 ----------

    def _after_round_actions(self):
        """
        委託送完（零股的話還撤完了）之後，決定這一批要不要繼續。

        沒勾多輪就到此為止；勾了就先做規劃文件流程最後那一步——**更新持股管理檔
        的現金、股數、成本**（見 _start_round_sync），再判斷出清了沒。
        """
        if not self.order_exec_active:
            # 停止之後才收到最後一個帳戶的撤單回話（見 _on_order_odd_cancelled）
            # ——撤單的結果照樣要報（那是人現在最需要知道的事，_odd_cancel_finished
            # 已經講過了），但絕對不能接著派下一輪同步。stop_order_execution 已經
            # 把畫面收好了，這裡什麼都不要再動。
            return

        if not self.order_exec_multi_round:
            self.order_exec_active = False
            self._set_busy(False)
            # 撤單那一段的結論要接在同一句裡（見 _odd_cancel_finished）：它跟
            # 「跑完了」是同一個時間點的兩件事，分兩句講的話後面那句會把前面
            # 蓋掉，人只看得到「跑完了」，看不到那批單後來怎麼了。
            head = f"{self.order_exec_cancel_note} " if self.order_exec_cancel_note else ""
            self._say(f"下單：{head}這一輪已經跑完。")
            self.order_exec_cancel_note = ""
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

        self._start_round_sync()

    def _start_round_sync(self):
        """
        規劃文件流程的最後一步：**更新持股管理檔的現金（初始餘額累加）、股數、
        成本**，然後才判斷出清了沒。

        走的是更新分頁那一整條路（`("fetch", …)` → `_on_fetched` → 寫 Excel →
        `_on_written` → `_commit_round`），等同程式自己按一次「更新（這幾位）
        帳戶」，不是另外寫一份簡化版：現金基準一天只設一次、`round_scope` 只准碰
        這一輪讀到的那幾位、寫入成功才落帳——這些規則都在那條路上，複製一份出來
        遲早會少掉其中一條（見 CLAUDE.md「改動前必看」）。

        **現金算法固定用初始餘額累加**（2026/09/03 使用者訂正，見
        `_order_round_sync_ready`），不是跟著更新分頁那顆全域開關走。`cash_method`
        是更新分頁跟這裡共用的同一個 Var，所以這裡用「暫時切過去、同步結束再切
        回來」而不是把它永久改掉——不然出清跑完之後，更新分頁會被靜靜換成初始
        餘額累加，跟使用者今天在那邊選的算法對不上。切回來的動作在
        `_order_sync_finished` 開頭，兩邊要對稱。

        **2026/09/03 之前這一步根本不存在**，多輪只有 _prepare_next_round 那次
        重讀 Excel。但 E/F 只有更新分頁寫得到（`excel_io.write_cells` 全專案只有
        `ui_background._write_worker` 一個呼叫端），所以重讀讀回來的永遠是下單前
        那份數字：第 2 輪會照著「還沒賣掉」的持股再算一次量，把同一批部位重複賣
        出去，一路重複到輪數上限為止。整張與零股共用這一支，就是因為那個缺口兩邊
        一模一樣。

        **不要「改成直接看網頁持股、省掉寫 Excel 這一段」**（2026/09/04 試過一次，
        當天就改回來，見 commit 8b9aa54 與它的 revert）。那個念頭是這樣來的：出清
        這兩個作業的量是從持股直接算的，而持股就是網頁那份未實現損益，寫進 Excel
        再讀回來確實只是把同一份數字轉手一次。**但那只對出清成立。**

        全持股交易／買賣股票的量**不是**從持股算的：張數與價格來自各帳戶自己那一頁
        的下單試算 M19:N28（見 `orders.plan_trade_orders` 與 docs/介面規劃.md 9.5），
        而那幾格是 `自動計算` 巨集拿 E/F/B8 算出來的。E/F/B8 沒有寫回去，下一輪讀到
        的試算就還是上一輪的計畫——跟「重讀 Excel 讀回下單前的股數」是同一種錯，只是
        換一組格子。

        所以這一段的作用不是搬數字，是**讓 Excel 的公式有機會重算**，那是網頁給不了
        的東西。目前接上多輪的只有出清（買賣股票切過去就強制關掉多輪，見
        `ui_order._on_order_job_changed`），所以這條路現在還沒有人非它不可——但它是
        全持股交易接上多輪的前提，2026/09/04 使用者確認要留著。
        """
        self.order_exec_sync_busy = True
        # 範圍就是這一輪凍結的那幾位，順序照執行順序（報酬率由低到高）——跟
        # _dispatch_next_order 用的是同一支對照（_order_number_for_sheet），不另外
        # 反查一次 trader_of，免得兩邊對出不同的人。
        selected = []
        for account in self.order_exec_accounts:
            order_number = self._order_number_for_sheet(account["sheet"])
            if order_number is not None:
                selected.append((order_number, self.accounts[order_number - 1]))
        if not selected:
            # 對不到任何一組帳號就沒辦法查網頁，也就沒辦法判斷出清了沒。硬跑下一
            # 輪等於拿下單前的舊持股再送一次，正是這一整段要消滅的那個錯。
            self.order_exec_sync_busy = False
            self._order_sync_failed("這一輪的帳戶對不到任何一組帳號，沒辦法更新持股管理檔。")
            return

        self.today = datetime.date.today()
        self.round_target = None
        self._say(f"下單：第 {self.order_exec_round} 輪已跑完，正在更新持股管理檔"
                  f"（現金、股數、成本）…")
        self._update_order_exec_ui()
        self._ensure_browser_thread()
        self.browser_waiting += 1
        # 暫時蓋掉更新分頁的算法選擇，讓這一趟 fetch → replan → 寫入全程都用
        # 初始餘額累加算（見本函式說明）。need_bank 因此固定是 False——方法一
        # 用不到銀行餘額，一併省掉那支查詢。_order_sync_finished 開頭會切回來。
        self._round_sync_prev_method = self.cash_method.get()
        self.cash_method.set(planner.METHOD_OPENING)
        self.browser_cmd_queue.put((
            "fetch",
            (selected, self.path, False),
        ))

    @staticmethod
    def _queue_signature(rows):
        """
        一輪委託佇列的指紋，給「沒有進展就停」那道保險比對用（見
        _on_order_price_refresh）。

        比的是**帳戶、股票、買賣別、數量**——也就是「這一輪要送出去的是不是同一
        批東西」。價格刻意不算進去：追價每一輪都會查一次即時委買賣一，價格本來
        就會跳，把它算進指紋的話兩輪永遠不會相等，這道保險就等於沒有。

        排序過再比，不倚賴佇列順序（帳戶順序是報酬率排出來的，理論上穩定，但
        不該讓一道安全機制去依賴那個假設）。
        """
        return tuple(sorted(
            (row["sheet"], row["code"], row.get("side") or "", row["lots"])
            for row in rows))

    def _round_zero_missing(self):
        """
        這一趟 `planner.plan()` 要把哪幾檔「網頁上已經不見了」當成出清完成、歸零
        寫回 Excel（見 planner.plan 的 zero_missing）。

        **只有多輪出清那一輪的同步期間才是非空的**，判斷依據是
        `order_exec_sync_busy`——那個旗標在 `_start_round_sync` 打開、
        `_order_sync_finished` 關掉，本來就只涵蓋這一趟同步。不另外開一個新變數
        記狀態：新變數萬一忘了清掉，更新分頁自己按「更新」時就會跟著歸零，
        那正是使用者明確不要的行為（人手動賣掉、忘了刪 Excel 那一列，程式不該
        自作主張清成 0）。

        範圍也只到「這一輪凍結的那幾檔股票」為止。同一次同步順便讀到的其他股票
        行為完全不變——程式沒動過它們，就沒有立場替它們下「已經賣光」的結論。
        """
        if not self.order_exec_sync_busy:
            return ()
        return {stock["code"] for stock in self.order_exec_stock_settings}

    def _order_sync_finished(self, ok, reason=""):
        """
        `_start_round_sync` 派出去那一輪同步的收尾，由 `ui_background._on_fetched`
        （沒有格子要寫的時候）與 `_on_written`（寫完之後）末端呼叫——那兩處本來
        就是更新這條路的兩個終點，多輪只是接在它們後面，不是另外開一條。

        使用者在同步期間按了「停止」的話這裡什麼都不做：那一輪同步照樣跑完（資料
        已經查回來了，寫進 Excel、落帳都是對的，沒有理由丟掉），只是不再接下一輪。
        """
        if not self.order_exec_sync_busy:
            return
        self.order_exec_sync_busy = False
        # 跟 _start_round_sync 開頭對稱：這一輪借用 cash_method 算完了，換回
        # 使用者在更新分頁原本選的那一種，不管這一輪成功、失敗還是使用者中途
        # 按了停止都要換回來（下面幾個 return 之前一定要先執行到這裡）。
        if self._round_sync_prev_method is not None:
            self.cash_method.set(self._round_sync_prev_method)
            self._round_sync_prev_method = None
        if not self.order_exec_active:
            return
        if not ok:
            self._order_sync_failed(reason)
            return
        # **busy 要自己補回來。** 同步走的是更新分頁那條路，而它的兩個終點
        # （_on_fetched／_on_written）都會 _set_busy(False)——那顆是整支程式共用
        # 的 cookie 鎖，鬆掉的話「登入」「更新」「全部登出」在第 2 輪開始前就變成
        # 可以按的，而它們都會換掉手上這組 cookie，接下來那一輪的委託就會掛到別
        # 人帳上（見本節開頭「依序執行」那段對送錯帳戶風險的說明）。
        self._set_busy(True, f"下單：第 {self.order_exec_round} 輪已跑完，重新讀取持股中…")
        # 同步已經把 E/F/B8 寫成最新的了，接著照原本那條路重讀一次 Excel：要不要
        # 先跑「更新股價」巨集、怎麼判斷出清了沒，都在那邊（_on_order_price_refresh）。
        self._prepare_next_round()

    def _order_sync_failed(self, reason):
        """同步沒成功就整批停下來——不拿舊持股硬跑下一輪（理由見 _start_round_sync）。"""
        self.order_exec_active = False
        self._set_busy(False)
        self._update_order_exec_ui()
        show_error(self.root,
            "更新持股管理檔失敗",
            f"第 {self.order_exec_round} 輪跑完之後要更新持股管理檔（現金、股數、成本），"
            f"但沒有成功，「多輪直到出清」先停在這裡：\n\n{reason}\n\n"
            f"這一輪送出去的委託不受影響，請自己到「更新」分頁跑一次，再決定要不要"
            f"接著跑下一輪。")
        self._say("下單：更新持股管理檔失敗，多輪出清已停止。")

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
        （I 欄）——巨集也好、重讀也好，都跟 refresh_order_plans 一樣只在
        COM 層面動，不碰瀏覽器，所以不必透過 browser_cmd_queue，直接開一條
        執行緒做（同 _order_plans_worker 的做法）。
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
                        # 一頁一次，理由同 _order_plans_worker。
                        if run_macro:
                            excel_io.run_update_price_macro(
                                excel, sheet, on_stuck=self._macro_stuck_notifier("更新股價", name))
                        data[name] = excel_io.read_sheet(sheet)
                # 巨集寫過 I4:I13 就要存檔，理由同 ui_order._order_plans_worker。
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
        # 清空「開始下單那一刻凍結的即時委買賣」，逼這一輪的每一列都退回
        # 「下單前才查」，不把已經過時的報價繼續當最新的用。
        #
        # **這一行是無條件的，而且不只第 2 輪以後會走到**：勾了「自動更新股價」
        # 時第 1 輪也會經過這裡（start_order_execution 設 round=0 就呼叫
        # _prepare_next_round），所以那個組合下，人按「查詢委買賣」查到的價格
        # 在按下「開始下單」之後就被丟掉、送單前重查一次。結果是對的（股價跟
        # 委買賣一都是新的，不會一新一舊），但**執行預覽那句「委買一 X 價送出」
        # 在這個組合下語意不正確**——它承諾「下單會直接用這個價格」，實際上會
        # 重查。沒勾自動更新股價時才是真的。2026/09/04 發現，還沒修，修法跟
        # 「多輪收斂改看網頁持股」那個更大的改動糾纏在一起，見記憶
        # order-multiround-pending-decisions。
        self.order_exec_quotes = {}

        side = self.order_exec_side
        # 重組隊列要問**凍結的**作業與單位，不是畫面上現在選的那個（見
        # order_exec_job／order_exec_unit）：使用者在多輪跑到一半切走了作業，
        # 這裡照畫面走就會拿另一套規則去算這一批的量。
        if self.order_exec_unit == orders.UNIT_ODD:
            preview = orders.plan_clear_odd_orders(
                self.order_exec_stock_settings, self.order_exec_accounts,
                self.order_holdings, self.order_exec_ticks,
                prices=self.order_exec_prices, quotes=self.order_exec_quotes)
            queue_rows = orders.executable_intraday_orders(preview)
        elif self.order_exec_mode == "intraday":
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
            # 整張那一版 queue 空了不一定是「真的出清歸零」——比重四捨五入到
            # 1 張，剩不到半張的零股永遠湊不出下一張整張委託，會一直卡在這裡，
            # 不能因為 queue 空了就跟人講「已經出清」，那不是真的。零股那一版
            # 沒有這個問題（見下面那個分支）。
            leftover = sum(item["held_qty"] for item in preview if item["held_qty"] > 0)
            if self.order_exec_round == 0:
                # 第一輪重讀完就發現沒東西可送——可能是持股在按下「開始
                # 下單」之後、巨集跑完之前這個空檔剛好變了，一輪都還沒真的
                # 跑，跟「跑了幾輪之後出清」是不同的事，訊息不能講「跑了
                # 0 輪」，那不是人話。
                self._say("下單：重讀 Excel 之後，沒有需要處理的委託了（可能持股剛好在這個空檔變了）。")
            elif self.order_exec_unit == orders.UNIT_ODD:
                # 零股的隊列空了就是**真的**沒有零股了：每一列的量就是持股的零股
                # 那一段（見 orders.plan_clear_odd_orders），不像整張那樣會卡在
                # 「剩下不到 1 張、比重永遠湊不出下一張」。剩下的整張是這個作業
                # 刻意不動的，不是沒出清——這句話要講清楚，不然人會以為還有事沒做完。
                rest = (f"（各帳戶手上還有 {show(leftover)} 股，都是整張的部分，"
                        f"這個作業不動它）" if leftover else "")
                self._say(f"下單：跑了 {self.order_exec_round} 輪，零股已經全部出清{rest}。")
            elif leftover > 0:
                self._say(f"下單：跑了 {self.order_exec_round} 輪，整張的部分已經出清，"
                          f"但還剩下不到 1 張的零股（比重換算不出下一張整張委託）。"
                          f"要把它清掉的話，把「單位」切到零股再跑一次。")
            else:
                self._say(f"下單：跑了 {self.order_exec_round} 輪，已經全部出清。")
            return

        # 沒有進展就停。這道保險跟「為什麼沒有進展」無關——2026/09/04 踩到的那次
        # 是 planner 對「網頁已無此檔」刻意不歸零（見 planner.plan 的 zero_missing），
        # 但任何讓持股沒跟著更新的原因，症狀都是這一輪算出來的委託跟上一輪一模
        # 一樣，然後照著已經賣掉的部位再送一次，一路重複到輪數上限。根因修好了
        # 這道還是要留：下一個沒想到的原因也會被它接住。
        signature = self._queue_signature(queue_rows)
        if signature == self.order_exec_last_signature:
            self.order_exec_active = False
            self._set_busy(False)
            self._update_order_exec_ui()
            self._say("下單：這一輪算出來的委託跟上一輪完全一樣，多輪出清已停止。")
            show_warning(self.root,
                "多輪出清沒有進展",
                f"第 {self.order_exec_round + 1} 輪算出來的委託跟上一輪完全一樣"
                f"（同帳戶、同股票、同數量），代表持股沒有跟著更新，"
                f"再跑下去會把同一批部位重複送出去。已經停在這裡。\n\n"
                f"請自己確認持股管理檔的股數是不是最新的，再決定要不要繼續。")
            return
        self.order_exec_last_signature = signature

        self.order_exec_round += 1
        self.order_exec_queue = queue_rows
        self.order_exec_pos = 0
        self._say(f"下單：開始第 {self.order_exec_round} 輪…")
        self._update_order_exec_ui()
        self._dispatch_next_order()

    def _update_order_exec_ui(self):
        # 「queue 是空的，但整批作業還在跑」現在有四種（見 order_exec_active 的
        # 說明），每一種都要讓「停止」維持可以按，而且要講得出卡在哪一段——都寫
        # 「處理中…」的話，撤單前的等待跟一次幾分鐘的同步在畫面上長得一模一樣，
        # 人會以為當掉了。順序就是它們實際發生的順序。
        seconds = ODD_CANCEL_WAIT_MS // 1000
        cancelling = (self.order_exec_cancel_queue
                      and self.order_exec_cancel_pos < len(self.order_exec_cancel_queue))
        if self.order_exec_cancel_timer is not None:
            waiting = f"零股已全部掛出，{seconds} 秒後自動取消沒成交的部分…"
        elif cancelling:
            waiting = (f"取消零股單：第 {self.order_exec_cancel_pos + 1}/"
                       f"{len(self.order_exec_cancel_queue)} 個帳戶…")
        elif self.order_exec_sync_busy:
            waiting = f"第 {self.order_exec_round} 輪已跑完，正在更新持股管理檔（現金、股數、成本）…"
        elif self.order_exec_price_busy:
            waiting = f"第 {self.order_exec_round} 輪已跑完，重新讀取持股中…"
        else:
            waiting = None

        if waiting is not None:
            self.order_exec_button.configure(text="處理中…", state="disabled")
            self.order_exec_stop_button.configure(state="normal")
            self.order_exec_status.configure(text=waiting)
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
        「讀取試算」讀到的數字——盤中模式的這顆按鈕會先觸發「更新股價」
        巨集才讀（見 refresh_order_plans），所以只要開始下單前有按過一次
        「讀取試算」，這個基準價就是新的，不是放到過期的舊資料。
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
        # 「這一列的價格要不要現在算」問那一列自己帶的 "chase"，不是問這一輪的
        # 模式（見 orders.plan_intraday_orders／plan_clear_odd_orders）：出清零股
        # 也是追價來的，但它的「時機」是被固定成盤中的，靠 mode 判斷等於把兩件
        # 事綁在一起——哪天零股不再固定盤中，這裡就會靜靜地送出一個 None 價格。
        # 舊的 queue（沒有這一欄）退回原本的判斷，行為不變。
        chase = row.get("chase", mode == "intraday")

        # 整張還是零股看**那一列自己帶的值**（理由同上面的 side）：這台引擎吃的
        # 是凍結好的 queue，不回頭問畫面上現在選的是哪一個——多輪之間、或人在
        # 執行中動了那顆單選鈕，畫面上的值跟這批 queue 算出來的量就對不起來了。
        # 沒帶這一欄的（出清那兩支 plan_*）當整張。row["lots"] 的單位跟著它走：
        # 整張是張、零股是股（見 orders.plan_trade_orders），兩支 order_fill 都要
        # 帶同一個 odd。
        #
        # 這一行在追價之前算，不是在下面填單前才算：追價要查的是**這一列自己那
        # 本簿子**，整股零股是兩本，同一時刻可以差好幾檔（2026/09/04 實測 2454
        # 整股 4395/4400、零股 4385/4390）。以前這裡是先追價才算 odd，零股那幾
        # 筆等於拿整股的委買賣一去追價——不會報錯，只是靜靜算錯一個價。
        odd = row.get("unit") == orders.UNIT_ODD

        if chase and price is None:
            pricenow = self.order_exec_prices.get(row["code"])
            if pricenow is None:
                raise RuntimeError(
                    f"沒有讀到 {row['code']} 的股價（Excel I 欄），這一筆沒辦法算追價。"
                    f"請先按「讀取試算」讓 Excel 更新股價。")

            # 對手方第一檔：一個 HTTP GET 查回來（見 stockinfo.py，不用登入、
            # 不開瀏覽器，取代原本開 FastQuote 彈出視窗訂閱 WebSocket 那條路）。
            # 查哪一邊跟 side 是反的：買方向查委賣一（ask）、賣方向查委買一
            # （bid），見 orders.chase_price 的說明。
            #
            # 查不到就讓 best_opposite 維持 None，chase_price 自己會退回邊界價、
            # 不擋單——這是原本 WebSocket 收不到時就有的行為，換來源之後照舊。
            # 例外一律吞掉也是同一個理由：行情查不到不該讓整筆委託送不出去。
            best_opposite = None
            try:
                quote = stockinfo.quote(row["code"], odd=odd)
            except Exception:
                quote = None
            if quote:
                best_opposite = quote["ask"] if row_side == orders.SIDE_BUY else quote["bid"]

            price = orders.chase_price(pricenow, ticks, row_side, best_opposite)

        # 委託別跟著模式走，不是兩邊共用同一個值（見 orders.BS_FLAG_PRE 的
        # 說明）：盤前開盤前還沒有連續交易，只能用 ROD；盤中規劃文件明講
        # 用 IOC。2026/08/28 使用者更正過，之前這裡兩種模式都寫死 IOC 是錯的。
        # 買賣股票那幾列自己帶著 bs_flag（規劃文件明講用 ROD），沒帶的才照模式
        # 決定：盤前只能 ROD，盤中規劃文件明講 IOC。
        bs_flag = row.get("bs_flag") or (
            orders.BS_FLAG_INTRADAY if mode == "intraday" else orders.BS_FLAG_PRE)

        page.goto(order_fill.ORDER_ENTRY_PAGE, wait_until="domcontentloaded")
        order_fill.open_order_form(page, odd=odd)
        order_fill.select_stock(page, row["code"])
        order_fill.fill_order(page, side=row_side, qty=row["lots"], price=price,
                              bs_flag=bs_flag, odd=odd)

        if not auto:
            return page, {}

        ok, message = order_fill.confirm_order(page)
        return page, {"auto_result": {"ok": ok, "message": message}}

    def _order_odd_cancel_job(self, context, store, order_number, account, sheet, codes):
        """
        背景執行緒用（只能在 ui_background._browser_worker 裡呼叫）：把這一輪掛出
        去、還沒成交的零股賣單撤掉（規劃文件「出清股票－零股」流程第 2 步）。

        **要先查一次才知道要撤哪幾筆，不能靠「程式記得自己送出了什麼」**：半自動
        模式下按確認的是人，而 `order_fill.confirm_order` 只回一句訊息、拿不到委託
        書號，程式手上根本沒有那份清單。改成查回來再過濾，連「人自己在瀏覽器裡按
        下確認」那幾筆也一起撤得到。

        過濾條件是 2026/09/03 使用者選定的範圍——**這一輪的帳戶 × 這一輪的股票**：
        還掛在外面（`open`）、賣出、盤別是盤中零股（`apcode` 等於
        `order_fill.TAB1_ODD`），而且股票在這一輪指定的那幾檔裡。三個條件都要：

        - 不看盤別的話會撤到整股的單。同一個帳戶同一檔股票很可能同時掛著整張的
          委託（別的作業送的，或人自己掛的），那不是這個流程該碰的東西。
        - 不看股票的話會撤到這一輪沒有指定的其他零股單。
        - 只看賣出：買進的零股單跟出清這件事無關。

        撤單本身借掛單分頁那一套（見 ui_pending._cancel_orders_split）：委託單與
        預約單分兩支、確認視窗要收乾淨、例外的兩種意思，那邊全部處理好了（10.3）。
        盤中送出去的零股單是委託單（`ordstatus == '2'`），但收盤後送的會轉成預約
        單，所以兩種都要顧——這也是共用那一支而不是自己寫一遍的理由。
        """
        page, session, problems = fetch_mod.ensure_logged_in(
            context, [(order_number, account)], store)[order_number]
        if problems:
            raise RuntimeError(f"{sheet}：{'；'.join(problems)}")

        wanted = set(codes)
        rows = [row for row in order_query.query_orders(page, session, sheet)
                if row["open"]
                and row["side"] == orders.SIDE_SELL
                and row["apcode"] == order_fill.TAB1_ODD
                and row["code"] in wanted]
        committed = [row["ordno"] for row in rows
                     if row.get("ordstatus") != "1" and row["ordno"]]
        reservation = [row["ordno"] for row in rows
                       if row.get("ordstatus") == "1" and row["ordno"]]

        if not committed and not reservation:
            # 一筆都沒查到不是失敗，是**這一輪掛出去的零股全部成交了**——出清零股
            # 最想看到的結果。回一個 found=0 讓收尾那邊講得出這件事（見
            # _odd_cancel_finished），不要丟例外。
            return {"results": [], "missing": [], "locked": [], "found": 0}

        combined = self._cancel_orders_split(page, session, sheet, committed, reservation)
        combined["found"] = len(committed) + len(reservation)
        return combined

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
