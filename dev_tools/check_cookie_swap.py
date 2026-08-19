# -*- coding: utf-8 -*-
"""
驗證「換人不重登」到底行不行 —— 唯一一件介面測不到、只能連上真網站才知道的事。

    python dev_tools/check_cookie_swap.py

背景：整個瀏覽器只有一組 cookie，同時只帶得動一個帳號的身分；但伺服器那邊每個
帳號的 session 是各自獨立的（B 登入不會殺掉 A 的）。所以「換交易人」在程式裡不是
重新登入，是把 A 登入時收下來的整份 cookie 換回瀏覽器（見 fetch._swap_cookies）。
這支腳本就是去確認網站吃不吃這一招。

會真的登入，但只讀不寫：不碰 Excel、不下單。跑之前把所有 Chrome 關掉
（同一個 profile 資料夾不能同時被兩個 Chrome 開著）。

兩種跑法，看 .env 裡有幾組真帳號：

    兩組以上  真的換人：登入 A -> 登入 B（A 被頂掉）-> 換回 A 的 cookie -> 抓 A 的資料
    只有一組  模擬被頂掉：登入 A -> 把 cookie 清光 -> 確認真的死了 -> 換回去 -> 抓資料

要看的三件事：
    [2] 期望 False —— 沒有它就表示「被頂掉」這件事根本不存在，那整套機制都不必要
    [3] 秒數很短、而且上面沒有印出「已自動取得並填入驗證碼」 = 真的沒有重登
    [4] retcode 000000，而且帳號代號是 A 的 = 換回去的身分是對的
"""

import sys
import time
from pathlib import Path

# 直接用 `python dev_tools/check_cookie_swap.py` 執行時，sys.path 起點是
# dev_tools 這一層，補回專案根目錄才找得到 fetch / login / recon。
_ROOT = Path(__file__).resolve().parent.parent
if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))

from playwright.sync_api import sync_playwright

import fetch
from login import configure_browsers_path, load_accounts, open_context
from recon import query


def main():
    configure_browsers_path()
    real = [(i, a) for i, a in enumerate(load_accounts(), start=1) if not a.get("fake")]
    if not real:
        print("找不到真帳號，請先在 .env 填 TBB_ID_1 / TBB_PASSWORD_1")
        return 1

    first, second = real[0], (real[1] if len(real) > 1 else None)
    order = first[0]
    print(f"真帳號 {len(real)} 組，"
          + ("用第二組去頂掉第一組（真的換人）" if second else "只有一組，改用清 cookie 模擬被頂掉"))

    with sync_playwright() as playwright:
        context, browser = open_context(playwright)
        store = fetch.new_store()
        try:
            start = time.time()
            page, session, problems = fetch.ensure_logged_in(context, [first], store)[order]
            print(f"[1] 登入第 {order} 組：{problems or '成功'}（{time.time() - start:.1f} 秒）")
            if not (session or {}).get("cust_id"):
                return 1
            print(f"    收下 {len(store['jars'][order])} 個 cookie，帳號代號 {store['codes'][order]}")

            if second:
                other = second[0]
                start = time.time()
                _p, other_session, other_problems = fetch.ensure_logged_in(context, [second], store)[other]
                print(f"[2] 登入第 {other} 組把它頂掉：{other_problems or '成功'}"
                      f"（{time.time() - start:.1f} 秒）")
                if not (other_session or {}).get("cust_id"):
                    return 1
                store["owner"] = other
            else:
                context.clear_cookies()
                store["owner"] = None
            dead = fetch._revisit(page, store["codes"][order])
            print(f"[2] 第 {order} 組的分頁還活著嗎：{dead is not None}（期望 False）")
            print(f"    它現在的網址：{page.url[:90]}")

            start = time.time()
            page, session, problems = fetch.ensure_logged_in(context, [first], store)[order]
            print(f"[3] 換回第 {order} 組的 cookie：{problems or '成功'}（{time.time() - start:.1f} 秒）")
            if not (session or {}).get("cust_id"):
                return 1

            bid, cid = session["branch_id"], session["cust_id"]
            data, raw = query(page, "queryInstantAccount_new",
                              {"branchId": "1" + bid, "custId": cid,
                               "range": "stksum,stkdat", "stock_no": ""})
            arrays = (data or {}).get("arrays") or []
            print(f"[4] 抓資料：retcode={(data or {}).get('retcode')}、{len(arrays)} 筆、"
                  f"身分 {fetch.account_code(session)}"
                  + ("" if data else f"、原始回應 {str(raw)[:120]}"))
        finally:
            context.close()
            if browser is not None:
                browser.close()
    return 0


if __name__ == "__main__":
    sys.exit(main())
