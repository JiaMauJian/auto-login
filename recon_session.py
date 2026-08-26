# -*- coding: utf-8 -*-
"""
偵察「這個網站認人到底是靠瀏覽器整組共用的 cookie，還是分頁自己有辦法認出身分」。

背景
----
fetch.py 的抓資料邏輯假設「整個瀏覽器只有一組 cookie，換帳號登入會把前一個頂掉」，
這件事在真實環境真的踩過一次 bug（2026/08/21，見 fetch.collect 的 docstring）：
20 組帳號全部登入完才回頭抓，前面那幾組拿到的其實是最後一組的資料。

但使用者反應：手動在同一個瀏覽器開 20 個分頁、20 個不同帳號，事後回頭任意點
哪個分頁，資料都還是對的——不管先後順序、不是「登入完立刻看」那種每次都點
最後一組的用法。這跟「共用一組 cookie」的假設矛盾，值得查清楚：會不會網站
其實在某個環節（例如網址裡夾帶的識別碼）讓每個分頁能認出自己是誰，只是我們
現在抓資料的方式（直接背景呼叫 API，只靠瀏覽器目前的 cookie）沒有用到那個機制。

這支腳本做的事
--------------
依序登入好幾組帳號（每組一個新分頁，過程中刻意不做任何 cookie 互換——這才是
在模擬使用者手動開分頁的情境）。全部登入完、瀏覽器目前帶著的 cookie 只剩
「最後登入那組」的之後，回頭測每一個分頁，兩種測法並列比較：

    測法一：分頁自己重新整理（模擬「回頭點分頁裡的連結」），看網站自己內部
            查詢回來的資料屬於誰。
    測法二：套用 fetch.py 現在的做法，直接對這個分頁背景呼叫查詢 API，
            看回來的資料屬於誰。

如果測法一每次都對、測法二會錯，就證實了使用者的觀察：網站有辦法讓分頁
「認出自己」，只是現在抓資料的方式沒吃到那個機制，值得再深入查怎麼讓
fetch.py 也用上（有機會整套拿掉「換 cookie」這個步驟，抓資料也許能一次
對好幾組平行處理，而不是被迫一組一組來）。如果兩種測法都錯，代表使用者
之前的操作方式可能其實是「登入完立刻看」（每次點的都剛好是最後一組），
而不是真的任意回頭點都對。

安全性
------
只讀、不寫，不會碰 Excel、不會下單、不會改網站上任何東西——跟 recon.py
同一個保證。**這支腳本刻意不把任何 cookie／session 的原始值印出來或存檔**：
那個值等同於這個帳號當下的登入憑證，外流出去等於讓別人能夠冒用這組帳號的
session。網址裡如果偵測到類似 session 識別碼的片段，只記錄「有沒有、
長度多少」，不記錄實際內容。輸出寫進 偵察資料/ 資料夾（已在 .gitignore
裡），而且全程只用帳號代碼（1112-0108640 這種，見 recon.py 的既有作法），
不碰密碼、不碰身分證字號。

用法：

    python recon_session.py

至少要兩組真帳號（.env 的 TBB_ID_1/2...）才測得出東西；帳號數愈多，
結果愈可信——這支只是多開幾個分頁，不會動到任何一組帳號本身的資料。
"""

import json
import re
import sys
import traceback
from datetime import datetime

from playwright.sync_api import Error as PlaywrightError, TimeoutError as PlaywrightTimeoutError
from playwright.sync_api import sync_playwright

from fetch import PAGE_READY_JS
from login import (
    app_dir,
    configure_browsers_path,
    do_login,
    load_accounts,
    open_context,
    pause,
    wait_until_finished,
)
from recon import ACCOUNT_PAGE, SESSION_JS, account_codes, query

OUTPUT_DIR_NAME = "偵察資料"

# 網址裡有沒有夾帶類似 session 識別碼的片段（老式 Java 網站常見的
# ;jsessionid=xxx 這種寫法）。只用來回答「有沒有」，抓到的內容本身不記錄。
SESSION_TOKEN_RE = re.compile(r"jsessionid=([^&;?/#]+)", re.IGNORECASE)

# 重新整理之後，給網站自己內部的查詢一點時間跑完，再去看攔到了什麼回應。
# 不用固定等更久：等不到就是等不到，報告裡會老實寫「沒攔到」。
SETTLE_AFTER_RELOAD_MS = 2000


def _token_len(url):
    """網址裡的 session 識別碼長度；沒有就回 None。只回長度，不回內容。"""
    match = SESSION_TOKEN_RE.search(url or "")
    return len(match.group(1)) if match else None


def _collect_codes_from_bodies(bodies):
    """把攔到的好幾份查詢回應（原始文字）解析出來，合併成一組帳號代碼。"""
    codes = set()
    for body in bodies:
        try:
            data = json.loads(body)
        except (json.JSONDecodeError, TypeError):
            continue
        codes |= account_codes(data)
    return codes


def _verdict(expect_code, got_codes, no_data_hint):
    """比對『這個分頁應該是誰』跟『這次查到的資料屬於誰』，回一句人看得懂的結論。"""
    if not got_codes:
        return f"無法判斷（{no_data_hint}）"
    if got_codes == {expect_code}:
        return "對，是自己的資料"
    return f"錯！資料屬於 {'、'.join(sorted(got_codes))}，不是自己的（{expect_code}）"


def _measure_ok(expect_code, got_codes):
    """跟 _verdict 同一個比對，只回 True/False/None（沒攔到資料，測不出來）。給結論彙整用。"""
    if not got_codes:
        return None
    return got_codes == {expect_code}


def _conclude(results, last_order):
    """
    把每個分頁的兩種測法結果，收成一句人看得懂的結論。

    只看「被晾著、cookie 已經不是自己的」那幾個分頁（排除最後登入那組——它的
    cookie 本來就是自己的，兩種測法一定都對，不能拿來當證據）。
    """
    displaced = [r for r in results if r["expect_code"] and r["order"] != last_order]
    if not displaced:
        return "帳號數不足（至少要兩組真帳號、且都成功登入），測不出結論。"

    v1 = [r["measure1_ok"] for r in displaced if r["measure1_ok"] is not None]
    v2 = [r["measure2_ok"] for r in displaced if r["measure2_ok"] is not None]

    if not v1 and not v2:
        return ("兩種測法這次都沒能攔到查詢結果（可能剛好都沒有庫存資料可以核對身分），"
                "測不出結論，建議再跑一次，或看報告細節。")

    if not v1:
        return ("測法一（分頁自己重新整理）這次沒攔到任何查詢結果，測不出它靠不靠 cookie，"
                "看報告細節自己判斷。")

    if all(v1) and v2 and not all(v2):
        return ("證實了你的觀察：分頁自己重新整理查得到自己的資料（不靠 cookie），"
                "但現在 fetch.py 的做法（背景直接呼叫 API，只看瀏覽器目前的 cookie）會查到別人的資料。"
                "值得深入研究怎麼讓 fetch.py 也用上這個機制，有機會不必換 cookie、甚至平行處理。")

    if not all(v1):
        return ("沒有證實你的觀察：分頁自己重新整理，也查得到別人的資料，不是每次回頭點都對。"
                "你之前手動開分頁的經驗，可能剛好都是登入完立刻看（或剛好點到最後登入那組）。"
                "目前 fetch.py「換 cookie」的架構還不能拿掉，細節看報告。")

    if v2 and all(v2):
        return ("這一次兩種測法都對，沒有重現出「認錯人」的現象——跟 fetch.py 之前踩過的 bug"
                "（20 組帳號抓錯資料）矛盾，建議多跑幾次、或確認這次真的沒有中途被重新登入，"
                "不能當作定論。")

    return "測法一、測法二的結果不完全一致，看報告細節自己判斷。"


def _login_all(context, accounts):
    """
    依序登入每一組帳號，過程中刻意不做任何 cookie 互換——這才是在模擬使用者
    手動一個一個分頁登入、彼此不管對方死活的情境。回傳每一組的分頁與身分。

    全部登入完之後，瀏覽器帶著的 cookie 只會是最後一組的，前面每一組的分頁
    都是「被晾在那裡、cookie 已經不是自己的」狀態——這正是測法一、測法二
    要拿來測試的處境。
    """
    tabs = []
    spare = context.pages[0] if context.pages else None

    for order, account in enumerate(accounts, start=1):
        # 不印身分證字號（即 account["id"]，見 login.do_login）——這支腳本承諾過
        # 不碰它，跟 recon.py／check_cookie_swap.py 同一個規矩：只講「第幾組」，
        # 真正的身分要登入完、從 sessionStorage 拿到帳號代碼才報。
        print(f"[{order}/{len(accounts)}] 登入第 {order} 組…")
        page = do_login(context, account["id"], account["password"], spare)
        spare = None

        page.goto(ACCOUNT_PAGE, wait_until="domcontentloaded")
        try:
            page.wait_for_function(PAGE_READY_JS, timeout=15000)
        except (PlaywrightError, PlaywrightTimeoutError):
            pass

        session = page.evaluate(SESSION_JS)
        bid, cid = session.get("branch_id"), session.get("cust_id")
        expect_code = f"1{bid}-{cid}" if bid and cid else None
        url_after_login = page.url

        tabs.append({
            "order": order,
            "page": page,
            "expect_code": expect_code,
            "url_after_login": url_after_login,
        })

        if not expect_code:
            print(f"    !! 沒有讀到登入身分（sessionStorage 空的），目前網址 {page.url}")

    return tabs


def _test_tab(tab):
    """
    對一個分頁跑兩種測法，回傳 (報告文字 list of str, 結構化結果 dict)。

    結構化結果給 _conclude 彙整結論用：
        {"order", "expect_code", "measure1_ok", "measure2_ok"}
    measure*_ok 是 True/False/None（None＝沒攔到資料，測不出來）。

    這時候瀏覽器目前帶著的 cookie 一定不是這個分頁自己的了（除非它剛好是
    最後登入那組）——兩種測法都是在問同一件事：不靠換 cookie，這個分頁
    能不能查回自己的資料。

    報告跟 result 都不放身分證字號（tab 裡本來就沒存，見 _login_all）——
    只用帳號代碼（1112-0108640 這種），跟 recon.py／check_cookie_swap.py 同一個規矩。
    """
    page, expect_code = tab["page"], tab["expect_code"]
    lines = [f"--- 第 {tab['order']} 組（帳號代碼 {expect_code or '（沒有）'}）---"]
    result = {"order": tab["order"], "expect_code": expect_code,
              "measure1_ok": None, "measure2_ok": None}

    token_len = _token_len(tab["url_after_login"])
    lines.append(
        f"登入完的網址{'有' if token_len else '沒有'}夾帶類似 session 的識別碼"
        + (f"（長度 {token_len}，內容不記錄）" if token_len else "")
    )

    if not expect_code:
        lines.append("這一組當初就沒有登入成功，兩種測法都跳過。")
        return lines, result

    if page.is_closed():
        lines.append("分頁已經被關掉了，兩種測法都跳過。")
        return lines, result

    # 測法一：分頁自己重新整理，看網站自己內部查詢回來的是誰的。
    bodies = []

    def on_response(resp, store=bodies):
        if "MainController" in resp.url:
            try:
                store.append(resp.text())
            except Exception:
                pass

    page.on("response", on_response)
    try:
        page.reload(wait_until="domcontentloaded")
        page.wait_for_timeout(SETTLE_AFTER_RELOAD_MS)
    except (PlaywrightError, PlaywrightTimeoutError) as exc:
        lines.append(f"測法一（分頁自己重新整理）：重新整理失敗，跳過（{exc}）")
        page.remove_listener("response", on_response)
        bodies = []
    else:
        page.remove_listener("response", on_response)
        codes1 = _collect_codes_from_bodies(bodies)
        result["measure1_ok"] = _measure_ok(expect_code, codes1)
        lines.append(
            "測法一（分頁自己重新整理，模擬回頭點連結）："
            + _verdict(expect_code, codes1, "這次沒攔到網站自己打的查詢，或剛好沒有庫存資料可以核對身分")
        )

    # 測法二：照 fetch.py 現在的做法，直接對這個分頁背景呼叫查詢 API。
    bid, cid = expect_code[1:].split("-", 1) if expect_code else (None, None)
    data, raw = query(page, "queryInstantAccount_new", {
        "branchId": "1" + bid, "custId": cid,
        "range": "stksum,stkdat", "stock_no": "",
    })
    if data is None:
        lines.append(f"測法二（現在的抓資料方式，直接呼叫 API）：查詢失敗，跳過（{str(raw)[:150]}）")
    else:
        codes2 = account_codes(data)
        result["measure2_ok"] = _measure_ok(expect_code, codes2)
        lines.append(
            "測法二（現在的抓資料方式，直接呼叫 API）："
            + _verdict(expect_code, codes2, "這次剛好沒有庫存資料可以核對身分")
        )

    return lines, result


def _collect(context, accounts):
    """
    核心流程：登入全部帳號（不換 cookie）、反過來逐一測試。回傳 (報告文字 list, 結構化結果 list)。

    CLI（main）跟圖形介面（run_headless）共用這一段，差別只在誰負責開/關瀏覽器、
    測完要不要留著給人繼續點。
    """
    report = []
    results = []
    tabs = _login_all(context, accounts)

    report.append("全部登入完成。瀏覽器目前帶著的 cookie 只會是最後一組的；")
    report.append("接下來反過來測，從最先登入、被晾最久的那一組開始看。")
    report.append("")

    for tab in reversed(tabs):
        lines, result = _test_tab(tab)
        report.extend(lines)
        report.append("")
        results.append(result)

    return report, results


def run_headless(accounts):
    """
    給圖形介面用：跑完整套偵察，回傳 (report_path, conclusion)。

    呼叫者要自己丟到背景執行緒跑（會真的開瀏覽器、依序登入，慢）。跟 main() 不同，
    這裡測完就把瀏覽器關掉，不留著等人手動操作——GUI 沒有終端機讓人看著它，
    留著也沒人知道要去點。
    """
    out_dir = app_dir() / OUTPUT_DIR_NAME
    out_dir.mkdir(exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M")

    configure_browsers_path()

    report = [
        f"偵察時間: {datetime.now():%Y/%m/%d %H:%M:%S}",
        f"本次測試組數: {len(accounts)}",
        "目的：確認『不靠換 cookie，分頁能不能各自查回自己的資料』，見檔案最上面的說明。",
        "",
    ]
    results = []

    with sync_playwright() as p:
        context, browser = open_context(p)
        try:
            body, results = _collect(context, accounts)
            report.extend(body)
        except PlaywrightTimeoutError:
            report.append("登入逾時，找不到欄位，網站版面可能已變更。")
        except PlaywrightError as exc:
            report.append(f"瀏覽器操作失敗：{exc}")
        finally:
            try:
                context.close()
                if browser is not None:
                    browser.close()
            except PlaywrightError:
                pass

    conclusion = _conclude(results, len(accounts))
    report.append("=" * 70)
    report.append("結論：")
    report.append(conclusion)

    report_path = out_dir / f"{stamp}_session偵察.txt"
    report_path.write_text("\n".join(report), encoding="utf-8")
    return report_path, conclusion


def main():
    accounts = [a for a in load_accounts() if not a.get("fake")]
    if len(accounts) < 1:
        print(f"至少要一組真帳號才跑得起來。請確認 {app_dir()} 的 .env 裡有 "
              f"TBB_ID_1/TBB_PASSWORD_1...")
        sys.exit(1)
    if len(accounts) < 2:
        print("!! 目前只有一組真帳號，這次只是跑一遍流程確認腳本本身沒問題，"
              "測不出『分頁會不會認錯人』——那個結論至少要兩組帳號才有意義。")

    out_dir = app_dir() / OUTPUT_DIR_NAME
    out_dir.mkdir(exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M")

    configure_browsers_path()

    report = [
        f"偵察時間: {datetime.now():%Y/%m/%d %H:%M:%S}",
        f"本次測試組數: {len(accounts)}",
        "目的：確認『不靠換 cookie，分頁能不能各自查回自己的資料』，見檔案最上面的說明。",
        "",
    ]

    results = []
    with sync_playwright() as p:
        context, browser = open_context(p)

        try:
            body, results = _collect(context, accounts)
            report.extend(body)
        except PlaywrightTimeoutError:
            report.append("登入逾時，找不到欄位，網站版面可能已變更。")
        except PlaywrightError as exc:
            report.append(f"瀏覽器操作失敗：{exc}")

        report.append("=" * 70)
        report.append("結論：")
        report.append(_conclude(results, len(accounts)))

        report_path = out_dir / f"{stamp}_session偵察.txt"
        report_path.write_text("\n".join(report), encoding="utf-8")

        print()
        print("\n".join(report))
        print()
        print("=" * 70)
        print(f"報告存在: {report_path}")
        print("這份報告沒有存密碼、也沒有存任何 cookie/session 的原始值，")
        print("可以放心把整份檔案內容貼給人看。")
        print("=" * 70)
        print("瀏覽器留著不關，你可以自己再點點看對照。")

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
