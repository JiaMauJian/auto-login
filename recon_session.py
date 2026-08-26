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
「最後登入那組」的之後，回頭測每一個分頁，三種測法並列比較：

    測法一：分頁自己重新整理（模擬「回頭點分頁裡的連結」），看網站自己內部
            查詢回來的資料屬於誰。
    測法二：套用 fetch.py 現在的做法，直接對這個分頁背景呼叫查詢 API，
            看回來的資料屬於誰。一組一組依序做，跟現在 fetch.py 的順序一樣。
    測法三：（2026/08/26 新增）不是一組一組依序做，是用 Promise.all 真的同時
            對每一組帳號各發一次背景查詢 API，看會不會彼此撞在一起。測法二
            就算每次都對，也只證明「事後依序查不會錯」，沒證明「好幾組同時
            查也不會錯」——而「同時查」才是真正評估「抓資料改平行處理」
            划不划算時，實際會發生的情境。

如果測法一每次都對、測法二會錯，就證實了使用者的觀察：網站有辦法讓分頁
「認出自己」，只是現在抓資料的方式沒吃到那個機制，值得再深入查怎麼讓
fetch.py 也用上（有機會整套拿掉「換 cookie」這個步驟，抓資料也許能一次
對好幾組平行處理，而不是被迫一組一組來）。如果兩種測法都錯，代表使用者
之前的操作方式可能其實是「登入完立刻看」（每次點的都剛好是最後一組），
而不是真的任意回頭點都對。測法三則是額外拿來單獨回答「平行處理到底安不
安全」這個問題：就算測法一、二都對，只要測法三出現撞資料，代表現在的
換 cookie 依序做法還不能拿掉。

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

import contextlib
import io
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

# 測法三用：跟 recon.FETCH_JS 幾乎一樣，差別是這裡一次接收好幾組帳號的查詢參數，
# 用 Promise.all 一次送出去——這才是「真的同時」，不是 Python 這邊一個一個呼叫
# page.evaluate（那樣還是被 Python 排隊，測不出瀏覽器真的同時發送會不會撞在一起）。
CONCURRENT_FETCH_JS = """
async ({ requests }) => {
    if (typeof B64_XOR_Encode !== 'function' || typeof XOR_KEY === 'undefined') {
        return requests.map(() => ({
            error: '這個頁面找不到 B64_XOR_Encode / XOR_KEY，common.js 可能沒載入'
        }));
    }
    const one = async ({ cmd, paramInfo }) => {
        const body = new URLSearchParams();
        body.set('CMD', cmd);
        body.set('Param', B64_XOR_Encode(JSON.stringify(paramInfo), XOR_KEY));
        const resp = await fetch('/tbb/MainController?timestamp=' + Date.now(), {
            method: 'POST',
            headers: { 'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8' },
            body: body.toString(),
            credentials: 'same-origin',
        });
        return { status: resp.status, text: await resp.text() };
    };
    return Promise.all(requests.map(one));
}
"""


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


def _test_concurrent(tabs):
    """
    測法三：不像測法一/二一組一組來，這裡用 Promise.all 真的同時對每一組帳號各發
    一次背景查詢 API，看會不會彼此撞在一起。

    不管由哪個分頁發起都一樣——cookie 是整個 context 共用的，這個網站分辨「查誰的
    資料」本來就不是靠「哪個分頁按下去」，是靠送出去的 branchId/custId 參數（跟
    fetch.py 現在的做法相同）。挑最後登入那組的分頁來發起，只是因為它一定還活著、
    還在帳戶頁上，不是因為發起的分頁有特殊地位。

    刻意不用 Python 這邊開執行緒/多個 page.evaluate 來模擬「同時」：Playwright 的
    同步 API 不能跨執行緒共用，硬做反而會出錯或變相排隊，測不出真正的同時。用
    瀏覽器自己的 Promise.all 在單一個 evaluate 呼叫裡一次送出去，才是「同時」
    這件事該有的樣子。

    回傳 (報告文字 list, {order: True/False/None})。
    """
    live = [t for t in tabs if t["expect_code"] and not t["page"].is_closed()]
    lines = ["--- 測法三：所有分頁同時（真的並行）呼叫背景 API ---"]

    if len(live) < 2:
        lines.append("帳號數不足（至少要兩組登入成功的帳號），跳過。")
        return lines, {}

    requests_payload = [
        {
            "cmd": "queryInstantAccount_new",
            "paramInfo": {
                "branchId": "1" + tab["expect_code"][1:].split("-", 1)[0],
                "custId": tab["expect_code"][1:].split("-", 1)[1],
                "range": "stksum,stkdat", "stock_no": "",
            },
        }
        for tab in live
    ]

    runner = live[-1]["page"]
    try:
        responses = runner.evaluate(CONCURRENT_FETCH_JS, {"requests": requests_payload})
    except (PlaywrightError, PlaywrightTimeoutError) as exc:
        lines.append(f"同時呼叫失敗，整批跳過（{exc}）")
        return lines, {}

    verdicts = {}
    for tab, resp in zip(live, responses):
        expect_code = tab["expect_code"]
        if isinstance(resp, dict) and resp.get("error"):
            lines.append(f"第 {tab['order']} 組（{expect_code}）：{resp['error']}")
            verdicts[tab["order"]] = None
            continue

        raw = resp.get("text", "") if isinstance(resp, dict) else ""
        try:
            data = json.loads(raw)
        except (json.JSONDecodeError, TypeError):
            lines.append(f"第 {tab['order']} 組（{expect_code}）：查詢失敗或回應不是 JSON")
            verdicts[tab["order"]] = None
            continue

        codes = account_codes(data)
        verdicts[tab["order"]] = _measure_ok(expect_code, codes)
        lines.append(
            f"第 {tab['order']} 組（{expect_code}）："
            + _verdict(expect_code, codes, "這次剛好沒有庫存資料可以核對身分")
        )

    return lines, verdicts


def _conclude_concurrent(results):
    """
    測法三專用的結論，跟 _conclude 分開講——問的是不同的問題。_conclude 關心的是
    「登入能不能同時做」跟「事後依序查會不會錯」；這裡關心的是「好幾組帳號同時
    呼叫背景 API，會不會彼此撞在一起」，這才是評估「抓資料改平行處理」划不划算
    真正該看的證據。

    這裡不排除最後登入那組：三個請求是同時發出去的，彼此互為對照，沒有哪一組是
    「自己 cookie 本來就對」的安全牌可以躺著過。
    """
    checked = [r["measure3_ok"] for r in results if r.get("measure3_ok") is not None]
    if not checked:
        return "測法三這次沒攔到任何一組的查詢結果，測不出結論。"

    if all(checked):
        return ("測法三：同時平行呼叫，每一組都拿到自己的資料，沒有彼此撞在一起——"
                "這是「拿掉換 cookie、抓資料改平行處理」這條路最直接的正面證據，值得認真評估。")

    wrong = sum(1 for ok in checked if not ok)
    return (f"測法三：{wrong}/{len(checked)} 組同時查詢時串到別人的資料——"
            "真的同時呼叫背景 API 會撞在一起，現在「換 cookie、依序做」的架構還不能拿掉，"
            "除非先查清楚是什麼原因撞車、能不能避開。")


def _login_all(context, accounts):
    """
    依序登入每一組帳號，過程中刻意不做任何 cookie 互換——這才是在模擬使用者
    手動一個一個分頁登入、彼此不管對方死活的情境。回傳每一組的分頁與身分。

    全部登入完之後，瀏覽器帶著的 cookie 只會是最後一組的，前面每一組的分頁
    都是「被晾在那裡、cookie 已經不是自己的」狀態——這正是測法一、測法二
    要拿來測試的處境。

    順便記錄每一組登入時，login.do_login 有沒有觸發它自己那條「登入頁沒有出現
    登入表單，清掉 cookie 再試一次」的備援路（見 login.py:316-323）。那段程式碼
    背後的假設是「換人登入、瀏覽器帶著上一組的 cookie，表單就會被卡住」——這個
    假設本身其實沒被真的驗證過，值得跟「回頭點分頁還是認得出自己」那個謎一樣，
    用實際跑出來的行為對照一次，而不是照單全收程式註解說的話。
    """
    tabs = []
    spare = context.pages[0] if context.pages else None

    for order, account in enumerate(accounts, start=1):
        # 不印身分證字號（即 account["id"]，見 login.do_login）——這支腳本承諾過
        # 不碰它，跟 recon.py／check_cookie_swap.py 同一個規矩：只講「第幾組」，
        # 真正的身分要登入完、從 sessionStorage 拿到帳號代碼才報。
        print(f"[{order}/{len(accounts)}] 登入第 {order} 組…")

        # do_login 內部會印身分證字號跟明文驗證碼（給終端機用的除錯訊息，不是這支
        # 腳本印的）。這裡整段攔下來，只從裡面找「有沒有觸發清 cookie 備援路」這
        # 一個是非值，攔到的原始文字用完即丟，絕對不會流進報告或存檔。
        buffer = io.StringIO()
        with contextlib.redirect_stdout(buffer):
            page = do_login(context, account["id"], account["password"], spare)
        needed_cookie_clear = "登入頁沒有出現登入表單" in buffer.getvalue()
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
            "needed_cookie_clear": needed_cookie_clear,
        })

        if not expect_code:
            print(f"    !! 沒有讀到登入身分（sessionStorage 空的），目前網址 {page.url}")

    return tabs


def _test_tab(tab):
    """
    對一個分頁跑兩種測法，回傳 (報告文字 list of str, 結構化結果 dict)。

    結構化結果給 _conclude 彙整結論用：
        {"order", "expect_code", "measure1_ok", "measure2_ok", "measure3_ok"}
    measure*_ok 是 True/False/None（None＝沒攔到資料，測不出來）。
    measure3_ok 這裡先填 None，等 _collect 跑完 _test_concurrent 之後才補上——
    測法三是所有分頁一次做，不是這支逐一測分頁的函式管得到的範圍。

    這時候瀏覽器目前帶著的 cookie 一定不是這個分頁自己的了（除非它剛好是
    最後登入那組）——兩種測法都是在問同一件事：不靠換 cookie，這個分頁
    能不能查回自己的資料。

    報告跟 result 都不放身分證字號（tab 裡本來就沒存，見 _login_all）——
    只用帳號代碼（1112-0108640 這種），跟 recon.py／check_cookie_swap.py 同一個規矩。
    """
    page, expect_code = tab["page"], tab["expect_code"]
    lines = [f"--- 第 {tab['order']} 組（帳號代碼 {expect_code or '（沒有）'}）---"]
    result = {"order": tab["order"], "expect_code": expect_code,
              "measure1_ok": None, "measure2_ok": None, "measure3_ok": None}

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

    # login.py 那條「換人登入表單被卡住、要清 cookie 才給」的備援路，
    # 這次跑下來到底有沒有真的被觸發——見 _login_all 的說明。
    triggered = [t["order"] for t in tabs if t.get("needed_cookie_clear")]
    if triggered:
        report.append(
            f"!! 第 {'、'.join(str(o) for o in triggered)} 組登入時，瀏覽器帶著別組的 cookie 導致"
            "登入頁一開始沒有畫出表單，程式自動清掉 cookie 才成功——這證實了 login.py 那個假設，"
            "換分頁登入下一組確實會被前一組的 cookie 卡住。"
        )
    else:
        report.append(
            "這次每一組登入時，登入表單都是一開始就直接出現，沒有一組觸發「清 cookie 才給表單」"
            "那條備援路——這跟 login.py 原本那個「換人登入一定會被上一組 cookie 卡住」的假設不符，"
            "值得懷疑那個假設是不是不成立（或只在特定情況才會發生），如果真的不會卡，"
            "「在同一個瀏覽器 profile 裡做到真平行登入」可能比之前想的更有機會，"
            "但這裡只驗證了「表單看不看得到」，登入送出後同一份 cookie 還是會被下一組頂掉，"
            "細節仍要另外查。"
        )
    report.append("")

    for tab in reversed(tabs):
        lines, result = _test_tab(tab)
        report.extend(lines)
        report.append("")
        results.append(result)

    concurrent_lines, concurrent_verdicts = _test_concurrent(tabs)
    report.extend(concurrent_lines)
    report.append("")
    for result in results:
        if result["order"] in concurrent_verdicts:
            result["measure3_ok"] = concurrent_verdicts[result["order"]]

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
    concurrent_conclusion = _conclude_concurrent(results)
    report.append("=" * 70)
    report.append("結論：")
    report.append(conclusion)
    report.append("")
    report.append(concurrent_conclusion)

    report_path = out_dir / f"{stamp}_session偵察.txt"
    report_path.write_text("\n".join(report), encoding="utf-8")
    return report_path, f"{conclusion}\n\n{concurrent_conclusion}"


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
        report.append("")
        report.append(_conclude_concurrent(results))

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
