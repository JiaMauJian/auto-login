"""
下單設定的純計算：比重轉張數、帳戶依報酬率排執行順序、組出執行預覽。

不碰 Excel、不碰瀏覽器——跟 planner.py 同樣的理由，介面跟之後要補的測試
可以共用同一份判斷邏輯，也可以拿假資料直接測。

服務兩個模式：
    盤前  一次設定（股票、比重、價格）套用到好幾個帳戶，價格是人手動填的
          固定值，預覽階段就算得出完整的執行清單（見 plan_stock_orders）。
    盤中  價格不是人填的，成交價來自 Excel I 欄（新增股票／讀取試算
          時就讀進來，見 ui_order.order_prices／order_exec_prices，不再現查
          網頁），下單前那一刻只再跟對手方第一檔比大小算出最終委託價
          （見 chase_price）。plan_intraday_orders 是純函式，不自己查 Excel
          或即時報價，但呼叫端（ui_order.py）如果已經把這兩份資料查回來
          （成交價、2026/08/29 新增「查詢委買賣」按鈕查到的即時委買賣一），
          可以當參數傳進來，這裡就直接算出實際會送出的價格；沒傳的話（還
          沒按過那顆按鈕）"price" 一律是 None，真正的價格留到填單那一步
          再算。
"""

from decimal import ROUND_HALF_UP, Decimal

from util import show, to_num

SHARES_PER_LOT = 1000

# 買賣別跟網站本身 buysell 的 B/S 代碼同一套（見 order-api-newOrder-encrypted
# 記憶）。2026/08/28 加了買/賣切換，是整批共用一個方向（跟 stock_settings
# 的「比重」一樣整批共用），不是每檔股票各自一個方向——切換買賣的時候畫面
# 會把股票清單清空重選（見 ui_order._on_order_job_changed），所以不會有
# 同一輪裡買賣混雜的情況，這裡的函式簽章才能只加一個 side 參數，不必把
# side 塞進每一筆 stock_settings。
SIDE_SELL = "S"
SIDE_BUY = "B"

# 「作業」：買賣股票／出清股票／全持股交易（見 docs/介面規劃.md 9.2、9.3）。
# 這三個不是三個分頁——它們的差別只有「左邊那格要填什麼」跟「張數與價格從哪
# 來」，右半邊（帳戶勾選、執行預覽、依序執行、多輪、自動送出）逐字相同，所以
# 往上長一層作業，不是往旁邊複製三份。
#
# 放在 orders.py 而不是某個 ui_ 檔：ui_layout（畫那三列）、ui_order（切換）、
# ui_order_exec（按鈕上的字）三個檔案都要用同一組值，而它們本來就都 import
# orders；之後 9.7 第 4 步每個作業各自的 plan_* 也會照這組值分流。
JOB_TRADE = "trade"
JOB_CLEAR = "clear"
JOB_FULL = "full"
JOB_NAMES = {JOB_TRADE: "買賣股票", JOB_CLEAR: "出清股票", JOB_FULL: "全持股交易"}

# 行為真的接上了的作業。還沒接的那個選得到、第二列也畫得出來（看得到版面長什麼
# 樣），但執行按鈕是灰的——見 ui_order._order_job_ready。
# 2026/08/30：買賣股票接上了（整張；零股還卡在下單表單的交易盤別要選哪個值）。
JOBS_READY = (JOB_CLEAR, JOB_TRADE)

# 單位：整張／零股。M19:M28 是股數不是張數，整張與零股是同一個數字拆兩段
# （見 9.4），「單位」決定這一輪送哪一段。
UNIT_LOT = "lot"
UNIT_ODD = "odd"
UNIT_NAMES = {UNIT_LOT: "整張", UNIT_ODD: "零股"}

# 執行預覽那一欄的欄名跟著單位換（2026/09/01 使用者指定）：單位寫在欄名上，
# 格子裡就只放數字，不必每一格都跟一個「張」或「股」。出清那兩個作業只有
# 整張，換不到「股數」那個字。
UNIT_COLUMN_TITLES = {UNIT_LOT: "張數", UNIT_ODD: "股數"}

# 哪個作業的哪個單位真的接上了——**不是三個作業共用一個開關**。「零股」在買賣
# 股票是同一個數字的另一半（照下單試算的股數送出去，見 9.4），在出清股票卻是
# 另一套流程（全部掛賣單、20 秒後全部取消，見規劃文件「出清股票－零股」），
# 兩邊不會同一天做完。共用一個開關的話，買賣股票這半邊接好的那天會順手把出清
# 那半邊還沒寫的流程也變成按得下去，而且不會報錯——它會照整張的流程送出去。
UNITS_READY = {
    JOB_TRADE: (UNIT_LOT, UNIT_ODD),   # 零股 2026/09/01 接上（order_fill.TAB1_ODD = "5"）
    JOB_CLEAR: (UNIT_LOT,),            # 零股：出清・零股 那套流程本身還沒寫
    JOB_FULL: (UNIT_LOT,),
}


def unit_ready(job, unit):
    """
    這個作業的這個單位接上了沒（見 UNITS_READY）。畫單選鈕（ui_layout.
    _build_order_unit）跟切作業時的保護（ui_order._on_order_job_changed）
    問的是同一支，兩邊不會各判各的。
    """
    return unit in UNITS_READY.get(job, (UNIT_LOT,))

# 委託別（bs_flag）兩個模式不一樣，不是同一個值兩邊共用：盤前是預約單，
# 掛在開盤前排隊等撮合，那個時間點根本還沒有連續交易，交易所不接受「當下
# 沒成交就取消」的 IOC/FOK，只能用 ROD（當日有效）；盤中規劃文件明講
# 「用IOC下單」（沒成交部位自動取消，不會像 ROD 掛著等）。2026/08/28 使用者
# 更正：早先以為盤前也能用 IOC 是錯的。
BS_FLAG_PRE = "R"
BS_FLAG_INTRADAY = "I"

# plan_stock_orders／plan_intraday_orders「note」欄位僅有的幾種固定文字，
# 不是使用者亂打的自由文字——執行預覽表格算「備註」欄要多寬時（見
# ui_layout._build_order_right）量的就是這幾句，拉成常數是為了讓兩邊量
# 的是同一份文字，這裡改了措辭那邊的欄寬會自動跟著變，不必兩處各改一次。
REASON_NO_HOLDING = "沒有這檔，略過"
REASON_UNDER_ONE_LOT = "比重算出來不到 1 張，略過"
REASON_NO_PRICE = "尚未設定價格"
REASON_CHASE_TEMPLATE = "以 Excel 成交價±{ticks}檔為邊界，下單前查{opposite}比價"
# 2026/08/29 使用者要求：先用「查詢委買賣」按鈕把即時委買賣一整批查回來
# （見 ui_order.fetch_order_quotes），這裡就能直接算出實際會送出的價格，
# 不必再等「開始下單」跑到那一筆才臨時查——note 講清楚「這個數字不會再變」，
# 跟 REASON_CHASE_TEMPLATE 那句「下單前查」是兩種不同的狀態，不能共用同一句。
REASON_CHASE_FROZEN_TEMPLATE = "已查{opposite} {value}，下單會直接用這個價格（不再重查）"

# 買賣股票（張數與價格來自 Excel 的下單試算 M19:N28）用得到的幾句。
REASON_NO_PLAN = "下單試算是空的，略過"
# 「這一位的持股清單裡沒有這一檔」跟「有這一檔但試算是 0」是兩件事
# （2026/09/02 使用者指定「其他帳戶沒找到這個股票就略過」）：前者是這一位
# 根本不玩這檔，後者是這一輪不必動它。兩種都略過，但寫同一句的話，人看不出
# 「是我選錯股票了嗎」還是「試算還沒跑」。
REASON_NO_STOCK = "這一位沒有這檔，略過"
# 勾了帳戶但還沒按「讀取試算」。這一位的持股與試算整份都還沒讀進來，看起來
# 會跟「試算是空的」一模一樣（都是讀不到數字）——不分開講的話，人會以為是
# Excel 那頭沒算，跑去按自動計算，而其實只要按「讀取試算」。
REASON_NOT_LOADED = "還沒讀取試算，略過"
REASON_ONLY_ODD_TEMPLATE = "試算只有零股 {odd} 股，整張是 0，略過"
REASON_WITH_ODD_TEMPLATE = "另有零股 {odd} 股，這一輪不送"
# 上面那句的鏡像：選零股的時候，沒送出去的是整張那一半。兩句都放備註欄，
# 「這一輪送多少」那一欄就只剩一個純數字（2026/09/01 使用者指定）。
REASON_WITH_LOT_TEMPLATE = "另有整張 {lots} 張，這一輪不送"


def split_lots(shares):
    """
    把股數拆成「整張幾張」與「零股幾股」兩段，各自帶回原本的正負號（見
    docs/介面規劃.md 9.4）。M19:M28 是**股數**不是張數，整張與零股不是兩份
    資料，是同一個數字拆兩段：

        2350 股  ->  (2, 350)     整張 2 張 ＋ 零股 350 股
        -800 股  ->  (0, -800)    整張 0 張 ＋ 零股 賣 800 股
        -2350 股 ->  (-2, -350)

    **往零的方向取整**——不是四捨五入，也不是 Python `//`：`-2350 // 1000`
    是 -3，那會變成「賣 3 張」，比 Excel 算出來的多賣一張。負數在這裡不是
    邊緣情況，是一半的情況（負數就是賣），所以這條不能靠「應該不會發生」。
    """
    sign = -1 if shares < 0 else 1
    qty = abs(int(shares))
    lots = qty // SHARES_PER_LOT
    return sign * lots, sign * (qty - lots * SHARES_PER_LOT)


def lots_from_weight(held_qty, weight_pct):
    """
    持股股數 × 比重（例如 25 代表 25%），四捨五入到最近的張。

    規劃文件的例子：8張 × 25% = 2張。

    四捨五入用 Decimal 的 ROUND_HALF_UP，不用內建 round()——後者是四捨六入
    五成雙（banker's rounding），2.5 會被算成 2 不是 3，跟「四捨五入」字面上
    的意思對不起來，這種差一張的錯不會報錯，只會默默少賣一張。
    """
    held_lots = Decimal(held_qty) / SHARES_PER_LOT
    lots = (held_lots * Decimal(weight_pct) / 100).quantize(Decimal("1"), rounding=ROUND_HALF_UP)
    return int(lots)


def order_accounts(accounts):
    """
    依今年報酬率（Excel B22）由低到高排序。

    accounts 是 [{"sheet": 分頁名, "return_rate": 數字或 None}, ...]。

    回傳 (ordered, skipped)：ordered 是排好序、多一個 "order"（從 1 開始）
    欄位的清單；skipped 是 B22 讀不到的那幾位，原樣列出來但不給順序——
    讀不到就當成最低硬排進去，等於用猜的決定誰先執行，比漏掉一位還危險，
    這裡刻意跟 planner.bank_balance「讀不懂就回 None，絕對不要回 0」同一種
    態度。

    這份順序同時是「執行帳戶」清單的排法**與**這一輪真正的執行順序（2026/09/02
    起一輪可以勾好幾位，見 docs/介面規劃.md 9.3 第 5 點）：號碼小的排前面、
    先送。09/01～09/02 之間一次只跑一位，那段期間它只是「建議的處理順序」。

    skipped 那幾位在清單上照樣勾得到、排在最後，號碼寫「－」，他們之間照活頁簿
    裡分頁本來的順序。這裡不給他們號碼的理由（用猜的決定誰先執行）還在——差別
    是「要不要跑他」現在由人一位一位勾，不是程式替他決定。
    """
    ready = [a for a in accounts if a.get("return_rate") is not None]
    skipped = [a for a in accounts if a.get("return_rate") is None]
    ready.sort(key=lambda a: a["return_rate"])
    ordered = [dict(a, order=i) for i, a in enumerate(ready, start=1)]
    return ordered, skipped


def plan_stock_orders(stock_settings, ordered_accounts, holdings, side):
    """
    組出「執行預覽」：排好序的帳戶 × 設定好的股票，一組一列。

    stock_settings 是 [{"code", "name", "weight_pct", "price"}, ...]，
    比重、價格照使用者確認過的決定，每檔股票各自設定，不是全部共用一個值。

    side 是 SIDE_SELL 或 SIDE_BUY，這一輪整批共用一個方向（見上面 SIDE_SELL
    的說明）。買跟賣共用同一條「比重 × 目前持股」公式（2026/08/28 使用者
    確認）：完全沒持有的股票，「比重」在買方向一樣算不出張數，會被判定
    「沒有這檔，略過」——這代表買方向目前只能加碼已經持有的部位，不能
    用這個功能建立全新部位，是刻意的限制，不是漏改。

    holdings 是 {(分頁名, 股票代號): 股數}，讀不到或該帳戶沒有這檔股票時
    對應的 key 不必存在，用 .get() 補 0。

    某個帳戶沒有這檔股票、或比重算出來不到 1 張，都在 note 裡講清楚原因、
    skip 設 True，不是默默省略那一列——半自動要讓人在送出前看得到「這一格
    為什麼被跳過」，不能只看到它消失。
    """
    preview = []
    for account in ordered_accounts:
        for stock in stock_settings:
            held_qty = holdings.get((account["sheet"], stock["code"]), 0) or 0
            lots = lots_from_weight(held_qty, stock["weight_pct"]) if held_qty else 0

            # skip 只跟「這一列會不會真的產生委託」有關（沒有持股／張數算出來
            # 是 0）；價格是另一回事——張數算得出來，缺的只是還沒填價格，
            # 這一列還在，只是多一句提醒，不算進 skip（見 CLAUDE.md：沒填的
            # 東西不能默默當成 0 或忽略，得讓人看得到還缺什麼）。
            reasons = []
            if held_qty == 0:
                reasons.append(REASON_NO_HOLDING)
            elif lots <= 0:
                reasons.append(REASON_UNDER_ONE_LOT)
            if not str(stock["price"]).strip():
                reasons.append(REASON_NO_PRICE)

            preview.append({
                "order": account["order"],
                "sheet": account["sheet"],
                "code": stock["code"],
                "name": stock["name"],
                "side": side,
                "held_qty": held_qty,
                "lots": lots,
                "price": stock["price"],
                "skip": held_qty == 0 or lots <= 0,
                "note": "；".join(reasons),
            })
    return preview


def plan_intraday_orders(stock_settings, ordered_accounts, holdings, ticks_down, side,
                         prices=None, quotes=None):
    """
    盤中模式的執行預覽，跟 plan_stock_orders 平行的一份。side 的規則（整批
    共用、買賣共用同一條比重公式）也跟 plan_stock_orders 一樣，見那邊的說明。

    stock_settings 是 [{"code", "name", "weight_pct"}, ...]——沒有 "price"，
    這是跟盤前最大的結構差異：盤中的成交價已經在新增股票／讀取試算時
    讀進 Excel（見 ui_order.order_prices／order_exec_prices）。

    prices／quotes 都是可選的（預設 None，當空字典用），呼叫端（ui_order.py）
    才碰得到 Excel 跟即時報價這兩份資料，這支函式本身仍然不碰 Excel、不碰
    瀏覽器——傳進來的只是現成的字典，不是要這裡自己去查：
        prices  code -> Excel I 欄讀回來的成交價（pricenow）
        quotes  code -> {"bid","ask","last"}（fastquote.FastQuoteStream.latest()
                的形狀），2026/08/29 新增「查詢委買賣」按鈕先整批查回來的
                即時委買賣一（見 ui_order.fetch_order_quotes）

    兩份都查得到同一檔股票的話，這裡就直接用 chase_price 算出實際會送出的
    價格填進 "price"，"note" 改講「已經查過、不會再變」；缺一份（通常是還
    沒按「查詢委買賣」，quotes 是空的）就跟以前一樣 "price" 留 None、"note"
    講清楚「下單前還會再查一次對手方第一檔比價」，不是留白也不是猜一個
    數字出來充數。

    skip 的判斷（沒有這檔／比重算出來不到 1 張）跟盤前完全一樣，因為這兩件事
    在預覽階段就看得出來，不需要等成交價。
    """
    prices = prices or {}
    quotes = quotes or {}
    preview = []
    for account in ordered_accounts:
        for stock in stock_settings:
            held_qty = holdings.get((account["sheet"], stock["code"]), 0) or 0
            lots = lots_from_weight(held_qty, stock["weight_pct"]) if held_qty else 0

            reasons = []
            price = None
            if held_qty == 0:
                reasons.append(REASON_NO_HOLDING)
            elif lots <= 0:
                reasons.append(REASON_UNDER_ONE_LOT)
            else:
                # 對手方第一檔查哪邊跟 side 是反的：買方向查委賣一、賣方向
                # 查委買一（見 chase_price 的說明），note 這句話跟著 side 換，
                # 不能整批寫死「委買一」。
                opposite = "委賣一" if side == SIDE_BUY else "委買一"
                pricenow = prices.get(stock["code"])
                quote = quotes.get(stock["code"])
                if pricenow is not None and quote is not None:
                    best_opposite = quote["ask"] if side == SIDE_BUY else quote["bid"]
                    price = chase_price(pricenow, ticks_down, side, best_opposite)
                    reasons.append(REASON_CHASE_FROZEN_TEMPLATE.format(
                        opposite=opposite, value=show(best_opposite)))
                else:
                    reasons.append(REASON_CHASE_TEMPLATE.format(ticks=ticks_down, opposite=opposite))

            preview.append({
                "order": account["order"],
                "sheet": account["sheet"],
                "code": stock["code"],
                "name": stock["name"],
                "side": side,
                "held_qty": held_qty,
                "lots": lots,
                "price": price,
                "skip": held_qty == 0 or lots <= 0,
                "note": "；".join(reasons),
            })
    return preview


def plan_trade_orders(stocks, ordered_accounts, plans, holdings, unit, loaded_sheets=None):
    """
    買賣股票的執行預覽（規劃文件「一、買賣股票」）：**張數與價格都來自各帳戶
    自己那一頁的下單試算 M19:N28**，人不必填任何數字；方向由試算股數的正負
    逐檔決定（正數買、負數賣），不是整批共用一個方向。

    所以這裡沒有 side 參數，也沒有 holdings／weight——跟 plan_stock_orders
    的差別不是「多幾個欄位」，是資料來源整個不一樣，硬要共用同一支只會變成
    兩條互不相干的路擠在一個函式裡。

    stocks 是使用者指定的股票（規劃文件：「指定股票（可多選）」），形狀跟
    plan_stock_orders 的 stock_settings 一樣是 [{"code", "name"}, ...]，只是不含
    比重與價格——名稱從這裡來而不是從 plans 來，是因為 plans 的名稱是 D 欄原文
    （含代號，像「台積電(2330)」），畫面上會變成「2330 台積電(2330)」。

    plans 是 {(分頁名, 股票代號): {"qty": 股數, "price": 價格}}，
    holdings 是 {(分頁名, 股票代號): 股數}（只為了在預覽裡顯示目前持股，這個
    作業不拿它算任何東西），
    unit 是 UNIT_LOT／UNIT_ODD 決定這一輪送哪一段（見 split_lots）。

    loaded_sheets 是「這一輪真的去 Excel 讀過的分頁」（不給就當成全部都讀過）。
    2026/09/02 起帳戶是勾選的、可以一次好幾位，勾了之後**還沒按「讀取試算」**
    的那幾位手上一格資料都沒有——不分開判的話，他們會跟「試算是空的」長得
    一模一樣（見 REASON_NOT_LOADED）。

    委託別固定 ROD（規劃文件明講「用ROD下單」）：這個作業沒有盤前／盤中之分，
    不吃 BS_FLAG_INTRADAY 那條路。
    """
    preview = []
    for account in ordered_accounts:
        for stock in stocks:
            code = stock["code"]
            sheet = account["sheet"]
            loaded = loaded_sheets is None or sheet in loaded_sheets
            # 讀過了才有資格說「這一位沒有這檔」——沒讀過的話 plans 裡本來
            # 就什麼都沒有，兩種情況在資料上分不出來，只有 loaded 分得出來。
            plan = plans.get((sheet, code))
            has_stock = plan is not None
            plan = plan or {}
            shares = int(plan.get("qty") or 0)
            price = plan.get("price")
            lots, odd = split_lots(shares)
            send = lots if unit == UNIT_LOT else odd

            reasons = []
            if not loaded:
                reasons.append(REASON_NOT_LOADED)
            elif not has_stock:
                # 使用者指定的股票，這一位的持股清單（D 欄）裡沒有——略過，
                # 不是錯誤：一次跑全部帳戶本來就會遇到有人沒這一檔。
                reasons.append(REASON_NO_STOCK)
            elif shares == 0:
                reasons.append(REASON_NO_PLAN)
            elif unit == UNIT_LOT and lots == 0:
                reasons.append(REASON_ONLY_ODD_TEMPLATE.format(odd=abs(odd)))
            elif unit == UNIT_LOT and odd:
                # 不是略過的理由，是一句提醒：選整張的時候送出去的量會比 Excel
                # 上那個數字少，人要看得出來為什麼（9.4）。
                reasons.append(REASON_WITH_ODD_TEMPLATE.format(odd=abs(odd)))
            elif unit == UNIT_ODD and lots:
                # 反過來那一半：選零股的時候沒送出去的是整張。
                reasons.append(REASON_WITH_LOT_TEMPLATE.format(lots=abs(lots)))
            # 價格只在「這一列本來真的要送」的時候才挑剔——已經因為沒讀到、
            # 沒這一檔、試算是空的而略過的列，再補一句「尚未設定價格」只是
            # 讓備註欄變長，人還要多讀一次才知道那不是真正的原因。
            if send and (to_num(price, None) is None or to_num(price, 0) <= 0):
                reasons.append(REASON_NO_PRICE)

            preview.append({
                "order": account["order"],
                "sheet": account["sheet"],
                "code": code,
                "name": stock.get("name", ""),
                "side": SIDE_BUY if shares > 0 else SIDE_SELL,
                "bs_flag": BS_FLAG_PRE,
                # 這一列送的是整張還是零股，寫進列裡跟著凍結好的 queue 一起走
                # （跟 side／bs_flag 同一個道理，見 ui_order_exec._order_fill_job）：
                # 下單表單的「交易盤別」要選哪一個、下面的 "lots" 是張還是股，
                # 都看這個值。出清那兩支 plan_* 不帶這一欄，執行端當整張。
                "unit": unit,
                # 「持股」欄就是持股，不塞試算股數——試算已經在「張數」欄有
                # 交代（見 _lots_text），這一欄拿來對照「要動的量跟手上有
                # 多少」反而有用。
                "held_qty": holdings.get((account["sheet"], code), 0) or 0,
                # 送出去的量是絕對值（方向靠 side，不是靠負號）。單位跟著
                # "unit" 走：整張是**張**、零股是**股**，同一個欄位名兩種單位，
                # 差 1000 倍——執行端填進 #qty 之前會照 unit 再量一次範圍
                # （見 order_fill._check_qty）。
                "lots": abs(send),
                "lots_text": _lots_text(lots, odd, unit),
                "price": price,
                "skip": send == 0,
                "note": "；".join(reasons),
            })
    return preview


def _lots_text(lots, odd, unit):
    """
    執行預覽那一欄要顯示的字（9.4）：**只有一個純數字**，沒有單位也沒有拆法。
    單位在欄名上（UNIT_COLUMN_TITLES：張數／股數），沒送出去的另一半在備註欄
    （REASON_WITH_ODD_TEMPLATE／REASON_WITH_LOT_TEMPLATE）——2026/09/01 使用者
    指定，原本這一欄是「847 股（另有 2 張）」那種寫法。

    出清那兩支 plan_* 不產生這一欄的文字，畫面上直接顯示 lots 那個數字，
    形狀本來就是純數字，這次的改法是讓買賣股票跟它們一致。
    """
    return str(abs(lots if unit == UNIT_LOT else odd))


def executable_orders(preview):
    """
    執行預覽裡真的可以送出委託的那幾列：沒被 skip（有持股、比重算得出張數），
    價格也填了、而且是個看得懂的正數。

    plan_stock_orders 的 skip 不管價格對不對（見那邊的說明，價格缺了只加提醒
    不算 skip）；半自動下單要操作真實表單，不能拿一個打不進網站欄位、或送出去
    意思不對的字串去填，這裡補上這一道，跟 fetch.settle_problem「讀不懂就整格
    擋住，不猜」同一種態度。
    """
    executable = []
    for row in preview:
        if row["skip"]:
            continue
        price = to_num(row["price"], None)
        if price is None or price <= 0:
            continue
        executable.append(row)
    return executable


def executable_intraday_orders(preview):
    """
    盤中模式的執行清單：只濾 skip，不驗價格——plan_intraday_orders 的每一列
    價格本來就還沒算（要等下單前那一刻查即時成交價才算得出來），拿「還沒發生
    的事」去擋會把整份清單濾空，跟 executable_orders 的判斷基準不一樣，不能
    共用同一個函式。
    """
    return [row for row in preview if not row["skip"]]


# 台股整股跳動點位（2020 年生效的最新一版），跟 dev_tools/simulate.py 的
# tick_size 同一份規則，只是那邊是假資料產生器只需要「看起來像」，這裡是真的
# 要送進下單表單的價格，補上原本沒有的「1000 元以上」那一階（simulate.py
# 500 元以上一律當 1 元一檔，1000 元以上其實是 5 元一檔，兩者只差在很少人會
# 買到那個價位的股票，不代表可以不管）。
_TICK_TIERS = ((10, 0.01), (50, 0.05), (100, 0.1), (500, 0.5), (1000, 1.0))
_TICK_ABOVE = 5.0


def tick_size(price):
    """price 這個價位的最小跳動單位（元）。"""
    for ceiling, unit in _TICK_TIERS:
        if price < ceiling:
            return unit
    return _TICK_ABOVE


def chase_price(pricenow, ticks_down, side, best_opposite=None):
    """
    依 2026/08/28 使用者確認的吃檔演算法算委託價（見 docs/自動下單與半自動
    下單規劃.pptx.txt「盤中」小節），不是單純「成交價−N檔」硬算：

    1. 先從 pricenow 算出邊界——買方向往上推 ticks_down 檔（最高願付價），
       賣方向往下推（最低願收價）。
    2. 邊界跟對手方第一檔的即時價格 best_opposite（買查委賣一、賣查委買一，
       方向不同查哪邊——呼叫端自己依 side 查好、對好方向再傳進來，這裡不
       負責查即時報價）比大小：買方向在邊界內（≤ 邊界）就用 best_opposite
       實際價格，超過邊界就用邊界價本身；賣方向相反（≥ 邊界才用實際價格）。

    best_opposite 傳 None 代表呼叫端還沒有即時委買/委賣一可查，直接回邊界
    價——邊界本身用檔位表就算得出來，缺的只是「能不能再貼近市價一點」這個
    優化，不是「有沒有價格可以掛」這件事，用邊界頂著先能動，不是在猜數字
    （跟 fetch.settle_problem「讀不懂就整格擋住，不猜」是不同情境）。

    每一步都用「當下那個價位」重新查 tick_size，不是拿起點價位的檔位算完
    一次乘以 ticks_down——價格每跨過一個級距門檻，最小跳動單位就會變小，
    一次算錯就是委託價格跟「規劃文件說的第 N 檔」對不起來，這種差一檔的錯
    不會報錯，只會默默用錯了價格（跟 lots_from_weight 的 ROUND_HALF_UP
    是同一種態度）。

    每一步都 round 到 2 位小數，收掉浮點累加誤差（例如 104.65 連減三次
    0.1 可能變成 104.34999999999999），不然送進網站價格欄位的字串會帶一長串
    小數，網站不一定接受。
    """
    boundary = float(pricenow)
    for _ in range(ticks_down):
        if side == SIDE_BUY:
            boundary = round(boundary + tick_size(boundary), 2)
        else:
            boundary = round(boundary - tick_size(boundary), 2)

    if best_opposite is None:
        return boundary
    if side == SIDE_BUY:
        return min(best_opposite, boundary)
    return max(best_opposite, boundary)
