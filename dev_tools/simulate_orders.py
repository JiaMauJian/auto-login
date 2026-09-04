"""
假帳號的下單／查詢掛單／撤單：不操作真實網站 DOM，直接呼叫假網頁裡的
`window.__SIM_ORDER__`（見 dev_tools/simulate.py render_html 尾端那段 JS）。

為什麼假帳號完全不走 order_fill.py／order_query.py／order_cancel.py
--------------------------------------------------------------------
下單分頁「依序執行」引擎（ui_order_exec.py）跟掛單分頁（ui_pending.py）真正
複雜、真正容易出錯的地方是「多輪收斂、等 6 秒撤零股、每輪重新同步」這一層
自己的協調邏輯（見 docs/介面規劃.md 9.8、9.9，那幾個坑都是這一層踩出來的）。
操作真實網站表單那幾行 click／fill／等 layer.js 視窗，是「網站長什麼樣子」
決定的，不是我們自己寫的協調邏輯，而且已經用真帳號測過一次（見
order_fill.py／order_cancel.py 模組說明）。

所以這裡選擇完全跳過那三支，把心力放在協調邏輯測不測得到——代價是這三支
操作真實 DOM 的程式碼本身，在假帳號這條路上一行都不會被跑到，那一層的
正確性要另外找機會用真帳號小單位驗證，不是這支程式的責任範圍。

跟 order_query.normalize() 共用同一份判讀
------------------------------------------
假網頁的 `queryOrders()` 回傳的原始形狀刻意跟真實 `queryOrder` 回應的欄位
對齊（orgqty/matqty/celqty/errcode/celable/buysell/ordstatus/apcode/...），
再丟給 `order_query.normalize()` 轉成畫面看得懂的列——不是另外寫一份轉換，
「同一套判讀只能有一份」跟 order_query.py 開頭講的是同一個理由。

成不成交、撤單怎麼決定，寫在 dev_tools/simulate.py 那段 JS 註解裡，不重複寫
兩遍。
"""

from order_fill import TAB1_LOT, TAB1_ODD, _check_qty
from order_query import normalize


def fill_order(page, *, code, side, qty, price, bs_flag, odd=False):
    """
    下一筆委託，回傳 (ok, message, matched_shares)。

    ok／message 形狀等同 order_fill.fill_order + confirm_order 合起來的結果：
    假帳號的假網頁沒有真正的委託確認視窗可以等人按，呼叫端
    （ui_order_exec._order_fill_job）一律當場決定成不成功，不看「自動送出」
    那顆開關——半自動模式存在的理由是留一步給人在真錢送出前看一眼，假帳號
    沒有那個風險，讓它卡在等一個畫不出來的視窗只會讓測試跑不完，見呼叫端的
    說明。

    matched_shares 是這一筆「當場」確定成交的股數（IOC 送出去那一刻就有
    答案；ROD 一律是 0，成不成交留到之後被查詢時才決定，見
    dev_tools/simulate.py 的 resolvePending）——不是猜的、也不是靠事後比對
    持股推論的，是跟下單同一次呼叫拿到的真實結果。呼叫端拿它來判斷「上一輪
    是不是真的什麼都沒發生」，比事後看持股有沒有變准，見
    ui_order_exec._on_order_filled 裡「沒有進展就停」那段的說明。

    qty／bs_flag 的檢查沿用 order_fill._check_qty 與零股只能 ROD 那條規則
    （直接呼叫同一支函式，不重寫一份——單位算錯、委託別選錯這兩種粗心，
    假帳號跟真帳號沒理由分開防）。
    """
    qty = _check_qty(qty, odd=odd)
    if odd and bs_flag != "R":
        raise RuntimeError(
            f"零股的委託別只有 R（ROD-當日有效），收到 {bs_flag!r}——這一筆不送。")

    apcode = TAB1_ODD if odd else TAB1_LOT
    result = page.evaluate(
        "(opts) => window.__SIM_ORDER__.place(opts)",
        {"code": code, "side": side, "qty": qty, "price": price,
         "bsFlag": bs_flag, "apcode": apcode},
    )
    return result["ok"], result["message"], result["matched"]


def query_orders(page, session, sheet):
    """
    形狀跟 order_query.query_orders(page, session, sheet) 一模一樣，呼叫端
    可以直接互換（見 ui_order_exec._order_odd_cancel_job／ui_pending._pending_job
    怎麼依 account.get("fake") 選要呼叫哪一支）。session 這裡用不到（假網頁
    沒有身分被頂掉這種風險），留著只是讓兩邊呼叫端不必為了多一個參數改寫法。
    """
    rows = page.evaluate("() => window.__SIM_ORDER__.query()")
    return [normalize(row, sheet) for row in rows]


def cancel_orders(page, sheet, ordnos):
    """
    把 ordnos（委託書號）這幾筆撤掉，回傳形狀跟 order_cancel.cancel_orders
    一樣：{"results": [...], "missing": [...], "locked": [...]}。

    假帳號沒有預約單這回事（沒有模擬「盤前掛單、收進預約序」那一層），
    呼叫端（ui_pending._cancel_orders_split）已經把 committed／reservation
    兩份合併成一份 ordnos 再傳進來，這裡不必再分兩套機制各打一次。
    """
    ordnos = list(ordnos)
    outcome = page.evaluate("(wanted) => window.__SIM_ORDER__.cancel(wanted)", ordnos)
    known = {row["ordno"] for row in outcome["results"]} | set(outcome["locked"])
    missing = [o for o in ordnos if o not in known]
    return {"results": outcome["results"], "missing": missing, "locked": outcome["locked"]}
