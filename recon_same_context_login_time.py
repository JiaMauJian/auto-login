# -*- coding: utf-8 -*-
"""
recon_concurrent_login_time.py 測出「cookie 搬到全新瀏覽器會被擋」之後的對照組：
改成沿用 recon_same_identity.py 已經驗證過會成功的做法——**同一個 context 內**
用 fetch._swap_cookies 換 cookie（不是搬去別的瀏覽器），依序把每一組的 cookie
換進來，各自導去歡迎頁抓「前次登入資訊」，看讀到的登入時間會不會不一樣。

跟 recon_concurrent_login_time.py 的差別只有「cookie 搬去哪裡」：那支搬去全新
瀏覽器 process（失敗，見對話與 memory）；這支留在同一個 context 換（跟
recon_same_identity.py 的 _revisit 一樣，已知會成功）。這裡不是「真的同時」，
是依序換、依序查，但因為換 cookie 本身很快，仍然能回答「前次登入資訊」是不是
各自 session 在登入當下就凍結住的值。

只讀、不寫，不碰 Excel。不印密碼、不印身分證字號、不印 cookie 原始值。

用法：python recon_same_context_login_time.py
"""

import sys
import traceback

from playwright.sync_api import Error as PlaywrightError, TimeoutError as PlaywrightTimeoutError
from playwright.sync_api import sync_playwright

from fetch import _open_account_page, _swap_cookies, account_code
from login import (
    configure_browsers_path,
    do_login,
    load_accounts,
    open_context,
    pause,
    wait_until_finished,
)
from recon import SESSION_JS

WELCOME_PAGE = "https://www.tbbstock.com.tw/tbb/welcome/layout.jsp?type=1"

# 歡迎頁「前次登入資訊」那張卡片：一個 label 是「前次登入資訊」的 .wel-cont，
# 裡面的 li 是登入時間跟 IP。
LAST_LOGIN_JS = """
() => {
    const cards = Array.from(document.querySelectorAll('.wel-cont'));
    for (const card of cards) {
        if (card.textContent.includes('前次登入資訊')) {
            return Array.from(card.querySelectorAll('li')).map((li) => li.textContent.trim());
        }
    }
    return null;
}
"""


def main():
    accounts = load_accounts()
    if len(accounts) < 2:
        print(f"至少要在 .env 設兩組帳號（可以都是同一個人）才測得出東西，"
              f"目前只有 {len(accounts)} 組。")
        sys.exit(1)

    configure_browsers_path()
    print(f"這次要依序登入 {len(accounts)} 次，帳號設定照 .env 裡的順序（可能是同一人）。")

    with sync_playwright() as p:
        context, browser = open_context(p)
        spare = context.pages[0] if context.pages else None

        jars = []  # [(order, code, jar, page)]

        for order, account in enumerate(accounts, start=1):
            print(f"[{order}] 登入中...")
            try:
                page = do_login(context, account["id"], account["password"], spare)
                spare = None
            except (PlaywrightTimeoutError, PlaywrightError) as exc:
                print(f"[{order}] 登入失敗：{exc}")
                continue

            _open_account_page(page)
            session = page.evaluate(SESSION_JS)
            code = account_code(session)
            if not code:
                print(f"[{order}] 登入後沒有讀到身分，跳過。目前網址: {page.url}")
                continue

            jars.append((order, code, context.cookies(), page))
            print(f"[{order}] 登入成功，帳號代碼 {code}")

        print()
        print("全部登入完成，現在依序把每一組的 cookie 換回來，各自去看歡迎頁的「前次登入資訊」...")
        print()

        results = []
        for order, code, jar, page in jars:
            _swap_cookies(context, jar)
            page.goto(WELCOME_PAGE, wait_until="domcontentloaded")
            info = page.evaluate(LAST_LOGIN_JS)
            results.append((order, code, info, page.url))
            shown = "、".join(info) if info else "（讀不到「前次登入資訊」）"
            print(f"[{order}]（{code}）前次登入資訊 -> {shown}　目前網址: {page.url}")

        times = [info[0] for _, _, info, _url in results if info]

        print()
        print("=" * 60)
        if len(times) < 2:
            print("讀到的資料不夠（至少要兩組成功），測不出結論。")
        elif len(set(times)) == len(times):
            print("結論：換回各自的 cookie 依序查，時間全部不一樣 —— 「前次登入資訊」是"
                  "各自 session 在登入當下就凍結住的值，不是帳號共用、會被最新登入蓋過去的值。")
        elif len(set(times)) == 1:
            print("結論：換回各自的 cookie 依序查，時間全部一樣 —— 這欄位看起來是帳號層級"
                  "共用的，不分是哪一個 session 在看，都讀到同一個值。")
        else:
            print(f"結論：{len(set(times))} 種不同的時間、{len(times)} 筆資料，"
                  "不是全同也不是全不同，細節看上面每一行。")
        print("=" * 60)

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
