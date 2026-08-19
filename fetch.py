"""
登入並抓回「未實現損益」與「當日淨收付」的原始資料。只讀網頁，不碰 Excel。

抓資料的方式是在已登入的分頁裡呼叫網站自己的 AJAX（POST /tbb/MainController），
沿用頁面現成的 B64_XOR_Encode / XOR_KEY 與 cookie，不做畫面解析 —— 表格是 JS
動態產生的，解析畫面既慢又容易被改版弄壞。
"""

from playwright.sync_api import Error as PlaywrightError, TimeoutError as PlaywrightTimeoutError

from dev_tools import simulate
from login import do_login
from profile_tools import CERT_MSG_JS, CERT_PAGE, parse_cert_expiry
from recon import ACCOUNT_PAGE, SESSION_JS, account_codes, query


def new_store():
    """
    瀏覽器活著的期間要記住的東西。由呼叫方保管，跨多次 ensure_logged_in 沿用；
    瀏覽器重開就換一份新的。

        pages   第幾組 -> 那一組在用的分頁
        jars    第幾組 -> 那一組登入完當下的整份 cookie
        codes   第幾組 -> 那一組的帳號代號（1分行-客戶號），換 cookie 後要拿來對身分
        owner   現在瀏覽器帶著的是哪一組的 cookie

    不帶 store 就是最原始的行為：每次都重新登入（命令列的 update_excel 走這條）。
    """
    return {"pages": {}, "jars": {}, "codes": {}, "owner": None}


def account_code(session):
    """sessionStorage 探到的身分 -> 帳號代號。網站的每一列資料都帶著它。"""
    bid, cid = (session or {}).get("branch_id"), (session or {}).get("cust_id")
    return f"1{bid}-{cid}" if bid and cid else ""


def _settle(page):
    """等頁面安定下來。等不到就算了 —— 這個網站常有背景請求一直在跑。"""
    try:
        page.wait_for_load_state("networkidle", timeout=15000)
    except PlaywrightTimeoutError:
        pass


def _fetch_cert_status(page):
    """
    在剛登入的分頁裡開一次「系統訊息與憑證狀態」頁，抓「憑證有效終止時間」。

    只在剛做完整套登入流程時呼叫一次（見 ensure_logged_in）—— 換 cookie 復用
    session 的那幾條路不會重跑，憑證到期日不會在同一次瀏覽器存活期間內變動，
    沒必要每次讀取都多開一次頁面。抓不到就靜靜回傳空值，不能讓這一步拖垮登入。
    """
    try:
        page.goto(CERT_PAGE)
        _settle(page)
        text = (page.evaluate(CERT_MSG_JS) or "").strip()
    except (PlaywrightError, PlaywrightTimeoutError):
        return "", None
    return text, parse_cert_expiry(text)


def _revisit(page, expect_code):
    """
    回到帳戶頁，確認「現在瀏覽器帶著的這組 cookie」真的還是這個人的。
    還活著就回傳探到的 session，否則回 None（呼叫方就去重登）。

    三道都要過：

        還在帳戶頁          session 過期的話網站會把人踢回登入／首頁
        sessionStorage 有值 頁面真的認得出登入身分
        身分跟預期的一樣    這是換 cookie 之後最重要的一道

    最後那道特別重要：分頁自己的 sessionStorage 不會因為 cookie 被換掉就失效，
    只看它的話，一個早就被頂掉的分頁探起來一樣「還活著」，照它走就會拿到別人的
    資料寫進別人的 Excel。
    """
    page.goto(ACCOUNT_PAGE)
    _settle(page)
    if "/account/" not in page.url:
        return None
    probe = page.evaluate(SESSION_JS)
    if not (probe.get("branch_id") and probe.get("cust_id")):
        return None
    if expect_code and account_code(probe) != expect_code:
        return None
    return probe


def _swap_cookies(context, jar):
    """
    把某一組登入時收下來的整份 cookie 換回瀏覽器。

    這是「換人不必重登」的全部秘密：整個瀏覽器只有一組 cookie，所以同時只帶得動
    一個人的身分；但伺服器那邊 20 個帳號是 20 個各自獨立的 session，B 登入不會
    殺掉 A 的 session，A 只是暫時沒有人帶著他的 cookie 去敲門而已。把 cookie 換
    回去就等於回到 A 登入完的那一刻，省掉整套登入流程（也省掉驗證碼）。

    只有一種情況換不回去：那一組在伺服器上已經逾時。那時候 _revisit 會看出來，
    照樣走重登 —— 最壞的結果等於沒有這個機制，不會更糟。
    """
    context.clear_cookies()
    context.add_cookies(jar)


def ensure_logged_in(context, selected, store=None):
    """
    確保每組帳號都已登入，回傳 {第幾組: (page, session, problems)}。

    session 是 SESSION_JS 探出來的 sessionStorage 內容，登入失敗時是 None；
    problems 非空代表這組帳號直接放棄（逾時或瀏覽器操作失敗）。

    store 見 new_store()。帶了它，同一組帳號會重複用自己的分頁，而且換人時先試
    「換回他的 cookie」，換不回去（逾時了）才重登。不帶就是每次都重新登入。

    一組帳號走到這裡有三條路，成本差很多：

        本來就是他的回合   直接回帳戶頁確認一下          最快
        換人、cookie 還在  換回他的 cookie 再確認一次    快（不必驗證碼）
        沒有 cookie 或逾時 整套登入流程跑一遍            慢
    """
    results = {}
    spare_page = context.pages[0] if context.pages else None

    for order, account in selected:
        reuse_page = (store["pages"].get(order) if store is not None else None)
        if reuse_page is not None and reuse_page.is_closed():
            reuse_page = None

        problems = []
        session = None
        page = reuse_page

        try:
            if account.get("fake"):
                # 模擬帳號：沒有登入這件事，只要確保它的假頁面開著就好（見 dev_tools/simulate.py）。
                # session 自己組一份，後面的程式碼就跟真帳號走同一條路。
                # 它不連網站，所以不碰 cookie，也不會把 owner 從真帳號手上搶走。
                page = simulate.open_page(context, account, reuse_page or spare_page)
                if reuse_page is None:
                    spare_page = None
                session = {
                    "branch_id": account["branch_id"],
                    "cust_id": account["cust_id"],
                    "account": account["name"],
                }
                if store is not None:
                    store["pages"][order] = page
                results[order] = (page, session, problems)
                continue

            if page is not None:
                if store is None:
                    # 沒有 store 就沒有 cookie 可以換，也不知道現在是誰的回合，
                    # 只能照最原始的做法探一下（探錯了還有 collect 那道身分檢查）。
                    session = _revisit(page, None)
                elif store["owner"] == order:
                    session = _revisit(page, store["codes"].get(order))
                elif store["jars"].get(order):
                    _swap_cookies(context, store["jars"][order])
                    session = _revisit(page, store["codes"].get(order))

            if session is None:
                page = do_login(context, account["id"], account["password"], reuse_page or spare_page)
                spare_page = None
                cert_text, cert_expiry = _fetch_cert_status(page)
                page.goto(ACCOUNT_PAGE)
                _settle(page)
                session = page.evaluate(SESSION_JS)
                if session is not None:
                    session["cert_text"] = cert_text
                    session["cert_expiry"] = cert_expiry.isoformat() if cert_expiry else None

            if store is not None:
                store["pages"][order] = page
                # 登入（或換 cookie）成功的那一刻，瀏覽器帶著的就是這一組的身分。
                # cookie 每次都重收一份：網站可能在過程中換過 JSESSIONID，
                # 留著舊的那份，下次換回去的就是一個已經沒用的值。
                if (session or {}).get("cust_id"):
                    store["owner"] = order
                    store["codes"][order] = account_code(session)
                    store["jars"][order] = context.cookies()
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


def login_only(context, selected, store=None):
    """
    只確保登入，不查資料。給「登入」按鈕用：先把瀏覽器開好、帳號登入好，
    之後按「讀取網頁資料」（見 collect）就能直接重查，不必再等一次自動登入。

    20 組一次登完之後，每一組的 cookie 都收在 store 裡（見 new_store），
    所以之後「只更新某一位」不必重登 —— 換回他的 cookie 就好。
    """
    records = []
    logins = ensure_logged_in(context, selected, store)

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

        record["account_code"] = account_code(session)
        record["sheet_name"] = (session.get("account") or "").strip()
        record["cert_text"] = session.get("cert_text")
        record["cert_expiry"] = session.get("cert_expiry")

    return records


def collect(context, selected, store=None):
    """
    逐一確保登入並抓資料。selected 是 [(第幾組, 帳號設定)]，回傳每組一筆記錄。

    刻意「登入完立刻抓」而不是全部登入完再回頭抓：所有分頁共用同一個 JSESSIONID，
    後登入的會把前一個的 server session 頂掉，晚一步抓可能拿到別人的資料。
    store 的用法見 new_store 與 ensure_logged_in。
    """
    records = []
    logins = ensure_logged_in(context, selected, store)

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

        record["account_code"] = account_code(session)
        record["sheet_name"] = (session.get("account") or "").strip()
        record["cert_text"] = session.get("cert_text")
        record["cert_expiry"] = session.get("cert_expiry")

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
