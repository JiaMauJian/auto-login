# -*- coding: utf-8 -*-
"""
延伸 recon_same_identity.py 的問題：同一身分的好幾個 session 都還活著之後，如果
「真的同時」（不是換 cookie 依序來）各自去看歡迎頁的「前次登入資訊」，拿到的登入
時間是各自登入當下凍結住的那一刻，還是同一個帳號共用、會被最新那次登入蓋過去的值？

分兩階段：

    第一階段（同步 API）：沿用 login.do_login，依序登入 .env 裡的每一組帳號
                          （可能都是同一人），收下每一組登入完當下的 cookie。
                          這段跟 recon_same_identity.py 的前半一樣，是已知穩定
                          的流程。

    第二階段（非同步 API）：開好幾個各自獨立的瀏覽器 context，各自套上其中一組
                          cookie，用 asyncio.gather 讓它們真的同時導去歡迎頁，
                          比較各自看到的「前次登入資訊」。一個 context 同時只能
                          帶一組 cookie，要讓好幾個 session 真的同時活著、同時
                          發請求，就必須是好幾個各自獨立的 context——這是跟
                          recon_session.py 的並行登入測試（_test_login_concurrency）
                          一樣的理由：Playwright 同步 API 不能跨執行緒安全共用，
                          只有 async API 的 asyncio.gather 能讓好幾個瀏覽器動作
                          真的重疊時間推進。

    第二階段刻意不沿用第一階段的 persistent profile／USER_DATA_DIR：基本登入
    （帳密 + 驗證碼）不需要用到數位憑證，這裡改用一次性、非 persistent 的
    Chromium／Chrome，直接把第一階段收下來的 cookie 套進去，不會跟憑證那份
    profile 打架（也不必等它關掉）。這等於是把「登入完當下的 cookie」原封不動
    搬到另一個瀏覽器視窗去用——跟 fetch.py 換 cookie 是同一招，只是搬去的地方
    從「同一個 context 換著用」變成「好幾個 context 同時用」。

只讀、不寫，不碰 Excel。不印密碼、不印身分證字號、不印 cookie 原始值，只印帳號
代碼跟頁面上看得到的登入時間。

用法：python recon_concurrent_login_time.py
"""

import asyncio
import sys
import traceback

from playwright.async_api import async_playwright
from playwright.sync_api import Error as PlaywrightError, TimeoutError as PlaywrightTimeoutError
from playwright.sync_api import sync_playwright

from fetch import _open_account_page, account_code
from login import (
    configure_browsers_path,
    do_login,
    launch_options,
    load_accounts,
    open_context,
    pause,
)
from recon import SESSION_JS

WELCOME_PAGE = "https://www.tbbstock.com.tw/tbb/welcome/layout.jsp?type=1"

# 歡迎頁「前次登入資訊」那張卡片：一個 label 是「前次登入資訊」的 .wel-cont，
# 裡面的 li 是登入時間跟 IP（見對話裡貼的 HTML 片段）。
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


def login_phase(accounts):
    """同步依序登入每一組帳號，回傳 [(order, code, jar)]；jar 是那組登入完當下的 context.cookies()。"""
    jars = []
    configure_browsers_path()

    with sync_playwright() as p:
        context, browser = open_context(p)
        spare = context.pages[0] if context.pages else None

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

            jars.append((order, code, context.cookies()))
            print(f"[{order}] 登入成功，帳號代碼 {code}")

        try:
            context.close()
            if browser is not None:
                browser.close()
        except PlaywrightError:
            pass

    return jars


async def concurrent_phase(jars):
    """開跟 jars 數量一樣多的獨立 context，各套一組 cookie，asyncio.gather 真的同時導去歡迎頁。"""
    async with async_playwright() as p:
        browser = await p.chromium.launch(**launch_options())

        pages = []
        for order, code, jar in jars:
            context = await browser.new_context()
            await context.add_cookies(jar)
            page = await context.new_page()
            pages.append(page)

        async def visit(order, code, page):
            await page.goto(WELCOME_PAGE, wait_until="domcontentloaded")
            info = await page.evaluate(LAST_LOGIN_JS)
            return order, code, info, page.url

        print()
        print(f"{len(jars)} 個獨立 session 同時導去歡迎頁...")
        results = await asyncio.gather(*(
            visit(order, code, page) for (order, code, _jar), page in zip(jars, pages)
        ))
        results.sort(key=lambda r: r[0])

        print()
        for order, code, info, url in results:
            shown = "、".join(info) if info else "（讀不到「前次登入資訊」）"
            print(f"[{order}]（{code}）前次登入資訊 -> {shown}　目前網址: {url}")

        times = [info[0] for _, _, info, _url in results if info]

        print()
        print("=" * 60)
        if len(times) < 2:
            print("讀到的資料不夠（至少要兩組成功），測不出結論。")
        elif len(set(times)) == len(times):
            print("結論：同時查，時間全部不一樣 —— 「前次登入資訊」是各自 session 在"
                  "登入當下就凍結住的值，不是帳號共用、會被最新登入蓋過去的值。")
        elif len(set(times)) == 1:
            print("結論：同時查，時間全部一樣 —— 這欄位看起來是帳號層級共用的，"
                  "不分是哪一個 session 在看，都讀到同一個值。")
        else:
            print(f"結論：{len(set(times))} 種不同的時間、{len(times)} 筆資料，"
                  "不是全同也不是全不同，細節看上面每一行。")
        print("=" * 60)

        for page in pages:
            await page.context.close()
        await browser.close()


def main():
    accounts = load_accounts()
    if len(accounts) < 2:
        print(f"至少要在 .env 設兩組帳號（可以都是同一個人）才測得出東西，"
              f"目前只有 {len(accounts)} 組。")
        sys.exit(1)

    jars = login_phase(accounts)
    if len(jars) < 2:
        print("成功登入的組數不夠（至少要 2 組），沒辦法比較，結束。")
        sys.exit(1)

    asyncio.run(concurrent_phase(jars))


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
