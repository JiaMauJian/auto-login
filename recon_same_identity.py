# -*- coding: utf-8 -*-
"""
小型偵察腳本：同一個真實身分依序登入好幾次，看伺服器允不允許同一個人的好幾個
session 同時活著。

背景
----
docs/分頁認人與並行登入偵察.md 測過的是「不同身分」能不能各自查回自己的資料、
能不能同時登入；這裡問的是不同的問題：**同一個身分**依序登入好幾次（每次都是
全新登入，不是換 cookie），前面那次的 session 會不會被後面的登入頂掉？

如果全部都還活著，代表這個網站對同一帳號沒有「單一 session」的限制，fetch.py
現有的換 cookie 機制（見 fetch._swap_cookies）對同一人一樣有效；如果只剩最後
一組活著，代表同一人開多分頁本來就註定只有最後一個能用，跟用不同帳號測出的
行為不一樣，值得記下來。

用法：先在 .env 設好幾組帳號（TBB_ID_1/TBB_ID_2...，可以都填同一個人），
然後 `python recon_same_identity.py`。

只讀、不寫，不碰 Excel。跟 recon.py / recon_session.py 一樣的安全保證：
不印密碼、不印身分證字號、不印 cookie 原始值，只印帳號代碼（1112-0108640 這種）。
"""

import sys
import traceback

from playwright.sync_api import Error as PlaywrightError, TimeoutError as PlaywrightTimeoutError
from playwright.sync_api import sync_playwright

from fetch import _open_account_page, _revisit, _swap_cookies, account_code
from login import (
    app_dir,
    configure_browsers_path,
    do_login,
    load_accounts,
    open_context,
    pause,
    wait_until_finished,
)
from recon import SESSION_JS


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
        print("全部登入完成，現在依序把每一組的 cookie 換回來，看還活不活著...")
        print()

        alive = dead = 0
        for order, code, jar, page in jars:
            _swap_cookies(context, jar)
            probe = _revisit(page, code)
            if probe:
                alive += 1
                print(f"[{order}]（{code}）換回它的 cookie 後 -> 還活著，資料是自己的")
            else:
                dead += 1
                print(f"[{order}]（{code}）換回它的 cookie 後 -> 已經失效（可能被後面的登入頂掉了）")

        print()
        print("=" * 60)
        if not jars:
            print("沒有任何一組成功登入，測不出結論。")
        elif alive == len(jars):
            print(f"結論：{len(jars)} 組全部還活著。同一身分可以同時撐好幾個 session，"
                  "fetch.py 現有的換 cookie 機制對同一人一樣有效。")
        elif alive == 1:
            print(f"結論：只有最後登入那組還活著，其餘 {dead} 組都被頂掉了。"
                  "同一身分同一時間只能有一個活的 session（單一登入限制）。")
        else:
            print(f"結論：{alive} 組還活著、{dead} 組被頂掉，不是全有全無，細節看上面每一行。")
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
