"""
下單設定的純計算：比重轉張數、帳戶依報酬率排執行順序、組出執行預覽。

不碰 Excel、不碰瀏覽器——跟 planner.py 同樣的理由，介面跟之後要補的測試
可以共用同一份判斷邏輯，也可以拿假資料直接測。

目前只服務「盤前」這個模式：一次設定（股票、比重、價格）套用到好幾個帳戶，
每個帳戶用自己的持股股數算出各自要賣的張數。
"""

from decimal import ROUND_HALF_UP, Decimal

SHARES_PER_LOT = 1000

# 盤前目前只支援賣出（規劃文件只有「設定賣出比重」，沒有買入這個選項），
# 買賣別先固定寫死在這裡，跟網站本身 buysell 的 B/S 代碼同一套（見
# order-api-newOrder-encrypted 記憶）。之後如果盤前也要支援買入，
# 這裡要改成 stock_settings 每檔各自帶一個方向，不是整批共用。
SIDE_SELL = "S"


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
