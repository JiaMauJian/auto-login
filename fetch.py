"""
登入並抓回「未實現損益」與「當日淨收付」的原始資料。只讀網頁，不碰 Excel。

抓資料的方式是在已登入的分頁裡呼叫網站自己的 AJAX（POST /tbb/MainController），
沿用頁面現成的 B64_XOR_Encode / XOR_KEY 與 cookie，不做畫面解析 —— 表格是 JS
動態產生的，解析畫面既慢又容易被改版弄壞。
"""

from playwright.sync_api import Error as PlaywrightError, TimeoutError as PlaywrightTimeoutError

import simulate
from login import do_login
from recon import ACCOUNT_PAGE, SESSION_JS, account_codes, query


def ensure_logged_in(context, selected, pages=None):
    """
    確保每組帳號都已登入，回傳 {第幾組: (page, session, problems)}。

    session 是 SESSION_JS 探出來的 sessionStorage 內容，登入失敗時是 None；
    problems 非空代表這組帳號直接放棄（逾時或瀏覽器操作失敗）。

    pages：選填，一個「第幾組 -> Page」的 dict，由呼叫方在多次呼叫之間保留同一份。
    帶了它，每組帳號會重複用自己上次登入的那個分頁：session 還活著就直接沿用，
    過期了才重新跑一次登入流程；不會每次都多開一個新分頁。分頁被關掉的話就當作沒有，
    重新開一個。不帶就是原本的行為：只重用瀏覽器自帶的第一個空白分頁，每次都重新登入。
    """
    results = {}
    spare_page = context.pages[0] if context.pages else None

    for order, account in selected:
        reuse_page = pages.get(order) if pages is not None else None
        if reuse_page is not None and reuse_page.is_closed():
            reuse_page = None

        problems = []
        session = None
        page = reuse_page

        try:
            if account.get("fake"):
                # 模擬帳號：沒有登入這件事，只要確保它的假頁面開著就好（見 simulate.py）。
                # session 自己組一份，後面的程式碼就跟真帳號走同一條路。
                page = simulate.open_page(context, account, reuse_page or spare_page)
                if reuse_page is None:
                    spare_page = None
                session = {
                    "branch_id": account["branch_id"],
                    "cust_id": account["cust_id"],
                    "account": account["name"],
                }
                if pages is not None:
                    pages[order] = page
                results[order] = (page, session, problems)
                continue

            if page is not None:
                # 上次登入用的分頁還在：cookie（JSESSIONID）也還在同一個 context 裡，
                # 這時候 LOGIN_URL 不會照常顯示登入表單（網站會把已登入的人導去別處），
                # do_login 的選擇器只會白等到逾時。所以先探一下 session 還活不活著，
                # 活著就直接跳過整套登入流程。
                page.goto(ACCOUNT_PAGE)
                try:
                    page.wait_for_load_state("networkidle", timeout=15000)
                except PlaywrightTimeoutError:
                    pass
                probe = page.evaluate(SESSION_JS)
                if probe.get("branch_id") and probe.get("cust_id"):
                    session = probe

            if session is None:
                page = do_login(context, account["id"], account["password"], reuse_page or spare_page)
                spare_page = None
                page.goto(ACCOUNT_PAGE)
                try:
                    page.wait_for_load_state("networkidle", timeout=15000)
                except PlaywrightTimeoutError:
                    pass
                session = page.evaluate(SESSION_JS)

            if pages is not None:
                pages[order] = page
        except PlaywrightTimeoutError:
            problems.append("登入逾時，找不到欄位，網站版面可能已變更")
        except PlaywrightError as exc:
            problems.append(f"瀏覽器操作失敗：{exc}")
        except RuntimeError as exc:
            # 模擬帳號讀不到 Excel（沒指定檔案、少了分頁）走這裡。只讓這一組失敗，
            # 其他帳號照跑 —— 這是測試用的路，不值得把整批讀取拖下水。
            problems.append(str(exc))

        results[order] = (page, session, problems)

    return results


def login_only(context, selected, pages=None):
    """
    只確保登入，不查資料。給「登入」按鈕用：先把瀏覽器開好、帳號登入好，
    之後按「讀取網頁資料」（見 collect）就能直接重查，不必再等一次自動登入。
    """
    records = []
    logins = ensure_logged_in(context, selected, pages)

    for order, account in selected:
        record = {"order": order, "problems": []}
        records.append(record)

        page, session, problems = logins[order]
        if problems:
            record["problems"].extend(problems)
            continue

        bid, cid = (session or {}).get("branch_id"), (session or {}).get("cust_id")
        if not bid or not cid:
            record["problems"].append(f"沒有登入成功（sessionStorage 沒有帳號資料），目前網址 {page.url}")
            continue

        record["account_code"] = f"1{bid}-{cid}"
        record["sheet_name"] = (session.get("account") or "").strip()

    return records


def collect(context, selected, pages=None):
    """
    逐一確保登入並抓資料。selected 是 [(第幾組, 帳號設定)]，回傳每組一筆記錄。

    刻意「登入完立刻抓」而不是全部登入完再回頭抓：所有分頁共用同一個 JSESSIONID，
    後登入的會把前一個的 server session 頂掉，晚一步抓可能拿到別人的資料。
    pages 的用法見 ensure_logged_in。
    """
    records = []
    logins = ensure_logged_in(context, selected, pages)

    for order, account in selected:
        record = {"order": order, "problems": []}
        records.append(record)

        page, session, problems = logins[order]
        if problems:
            record["problems"].extend(problems)
            continue

        bid, cid = (session or {}).get("branch_id"), (session or {}).get("cust_id")
        if not bid or not cid:
            record["problems"].append(f"沒有登入成功（sessionStorage 沒有帳號資料），目前網址 {page.url}")
            continue

        record["account_code"] = f"1{bid}-{cid}"
        record["sheet_name"] = (session.get("account") or "").strip()

        queries = {
            "未實現損益": {"branchId": "1" + bid, "custId": cid, "range": "stksum,stkdat", "stock_no": ""},
            "當日淨收付": {
                "branchId": "1" + bid, "custId": cid,
                "his": "y", "queryType": "1",
                "range": "stksum,stkdat,matsum,matdat", "stock_no": "",
            },
        }

        # 模擬帳號的資料是從它自己的假頁面即時算出來的，形狀跟真 API 的回應一樣，
        # 所以下面的 retcode 檢查、帳號回顯檢查全部照跑，不特別放水。
        simulated = None
        if account.get("fake"):
            try:
                simulated = simulate.read_page(page)
            except (PlaywrightError, RuntimeError) as exc:
                record["problems"].append(f"讀取模擬頁面失敗：{exc}")
                continue

        for title, param_info in queries.items():
            if simulated is not None:
                data, raw = simulated.get(title), "（模擬資料）"
            else:
                data, raw = query(page, "queryInstantAccount_new", param_info)
            if data is None:
                record["problems"].append(f"{title} 查詢失敗：{str(raw)[:200]}")
                continue
            if data.get("retcode") != "000000":
                record["problems"].append(f"{title} 回應異常：{data.get('retcode')} {data.get('retmsg')}")
                continue

            # 每一列都帶 bhno/cseq。跟這個分頁登入的身分不符，就是 session 被別人頂掉了。
            codes = account_codes(data)
            if codes and codes != {record["account_code"]}:
                record["problems"].append(
                    f"{title} 的資料屬於 {'、'.join(sorted(codes))}，"
                    f"與登入的 {record['account_code']} 不符（session 可能被其他帳號頂掉）"
                )
                continue

            record[title] = data.get("arrays") or []

    return records
