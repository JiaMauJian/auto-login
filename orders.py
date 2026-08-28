"""
下單設定的純計算：比重轉張數、帳戶依報酬率排執行順序、組出執行預覽。

不碰 Excel、不碰瀏覽器——跟 planner.py 同樣的理由，介面跟之後要補的測試
可以共用同一份判斷邏輯，也可以拿假資料直接測。

服務兩個模式：
    盤前  一次設定（股票、比重、價格）套用到好幾個帳戶，價格是人手動填的
          固定值，預覽階段就算得出完整的執行清單（見 plan_stock_orders）。
    盤中  價格不是人填的，是下單前那一刻用成交價往下追 N 檔算出來的
          （見 chase_price），預覽階段查不到即時成交價，所以
          plan_intraday_orders 的每一列價格是 None，真正的價格由呼叫端
          （ui_order.py 的填單那一步）另外查、另外算。
"""

from decimal import ROUND_HALF_UP, Decimal

from util import to_num

SHARES_PER_LOT = 1000

# 盤前目前只支援賣出（規劃文件只有「設定賣出比重」，沒有買入這個選項），
# 買賣別先固定寫死在這裡，跟網站本身 buysell 的 B/S 代碼同一套（見
# order-api-newOrder-encrypted 記憶）。之後如果盤前也要支援買入，
# 這裡要改成 stock_settings 每檔各自帶一個方向，不是整批共用。
SIDE_SELL = "S"

# 委託別（bs_flag）兩個模式不一樣，不是同一個值兩邊共用：盤前是預約單，
# 掛在開盤前排隊等撮合，那個時間點根本還沒有連續交易，交易所不接受「當下
# 沒成交就取消」的 IOC/FOK，只能用 ROD（當日有效）；盤中規劃文件明講
# 「用IOC下單」（沒成交部位自動取消，不會像 ROD 掛著等）。2026/08/28 使用者
# 更正：早先以為盤前也能用 IOC 是錯的。
BS_FLAG_PRE = "R"
BS_FLAG_INTRADAY = "I"


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
    依今年報酬率（Excel B17）由低到高排出執行順序。

    accounts 是 [{"sheet": 分頁名, "return_rate": 數字或 None}, ...]。

    回傳 (ordered, skipped)：ordered 是排好序、多一個 "order"（從 1 開始）
    欄位的清單；skipped 是 B17 讀不到的那幾位，原樣列出來但不給順序——
    讀不到就當成最低硬排進去，等於用猜的決定誰先執行，比漏掉一位還危險，
    這裡刻意跟 planner.bank_balance「讀不懂就回 None，絕對不要回 0」同一種
    態度。
    """
    ready = [a for a in accounts if a.get("return_rate") is not None]
    skipped = [a for a in accounts if a.get("return_rate") is None]
    ready.sort(key=lambda a: a["return_rate"])
    ordered = [dict(a, order=i) for i, a in enumerate(ready, start=1)]
    return ordered, skipped


def plan_stock_orders(stock_settings, ordered_accounts, holdings):
    """
    組出「執行預覽」：排好序的帳戶 × 設定好的股票，一組一列。

    stock_settings 是 [{"code", "name", "weight_pct", "price"}, ...]，
    比重、價格照使用者確認過的決定，每檔股票各自設定，不是全部共用一個值。

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
                reasons.append("沒有這檔，略過")
            elif lots <= 0:
                reasons.append("比重算出來不到 1 張，略過")
            if not str(stock["price"]).strip():
                reasons.append("尚未設定價格")

            preview.append({
                "order": account["order"],
                "sheet": account["sheet"],
                "code": stock["code"],
                "name": stock["name"],
                "side": SIDE_SELL,
                "held_qty": held_qty,
                "lots": lots,
                "price": stock["price"],
                "skip": held_qty == 0 or lots <= 0,
                "note": "；".join(reasons),
            })
    return preview


def plan_intraday_orders(stock_settings, ordered_accounts, holdings, ticks_down):
    """
    盤中模式的執行預覽，跟 plan_stock_orders 平行的一份。

    stock_settings 是 [{"code", "name", "weight_pct"}, ...]——沒有 "price"，
    這是跟盤前最大的結構差異：盤中的價格是下單前那一刻用成交價往下追
    ticks_down 檔算出來的（見 chase_price），預覽階段沒有即時成交價可以算，
    這裡的 "price" 一律是 None，"note" 改講清楚「價格會在下單前才查」，不是
    留白、也不是猜一個數字出來充數。

    skip 的判斷（沒有這檔／比重算出來不到 1 張）跟盤前完全一樣，因為這兩件事
    在預覽階段就看得出來，不需要等成交價。
    """
    preview = []
    for account in ordered_accounts:
        for stock in stock_settings:
            held_qty = holdings.get((account["sheet"], stock["code"]), 0) or 0
            lots = lots_from_weight(held_qty, stock["weight_pct"]) if held_qty else 0

            reasons = []
            if held_qty == 0:
                reasons.append("沒有這檔，略過")
            elif lots <= 0:
                reasons.append("比重算出來不到 1 張，略過")
            else:
                reasons.append(f"成交價往下 {ticks_down} 檔（下單前才查）")

            preview.append({
                "order": account["order"],
                "sheet": account["sheet"],
                "code": stock["code"],
                "name": stock["name"],
                "side": SIDE_SELL,
                "held_qty": held_qty,
                "lots": lots,
                "price": None,
                "skip": held_qty == 0 or lots <= 0,
                "note": "；".join(reasons),
            })
    return preview


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


def chase_price(pricenow, ticks_down):
    """
    從成交價 pricenow 開始往下追 ticks_down 檔，回傳算出來的限價。

    每一步都用「當下那個價位」重新查 tick_size，不是拿起點價位的檔位算完
    一次乘以 ticks_down——價格每跨過一個級距門檻，最小跳動單位就會變小，
    一次算錯就是委託價格跟「規劃文件說的第 N 檔」對不起來，這種差一檔的錯
    不會報錯，只會默默用錯了價格（跟 lots_from_weight 的 ROUND_HALF_UP
    是同一種態度）。

    每一步都 round 到 2 位小數，收掉浮點累加誤差（例如 104.65 連減三次
    0.1 可能變成 104.34999999999999），不然送進網站價格欄位的字串會帶一長串
    小數，網站不一定接受。
    """
    price = float(pricenow)
    for _ in range(ticks_down):
        price = round(price - tick_size(price), 2)
    return price
