"""
登入並抓回帳務資料的原始回應。只讀網頁，不碰 Excel。

固定抓兩支：未實現損益、當日淨收付。現金餘額用「銀行餘額推算」那種算法時
再多抓交割金額與銀行餘額（見 docs/現金餘額兩種算法.md）—— 用另一種算法的日子
一次都用不到，20 個帳號省下 40 次往返。

抓資料的方式是在已登入的分頁裡呼叫網站自己的 AJAX（POST /tbb/MainController），
沿用頁面現成的 B64_XOR_Encode / XOR_KEY 與 cookie，不做畫面解析 —— 表格是 JS
動態產生的，解析畫面既慢又容易被改版弄壞。
"""

from playwright.sync_api import Error as PlaywrightError, TimeoutError as PlaywrightTimeoutError

from dev_tools import simulate
from login import do_login
from recon import ACCOUNT_PAGE, SESSION_JS, account_codes, query
from util import to_num


def settle_problem(rows):
    """
    交割金額查詢（query610）的回應能不能用。可以就回 None，不能就回一句給人看的話。

    這支的回應一天一列，金額已經是那一天的淨額（負數＝要付出去）：

        {"trade": "20260821", "cdate": "20260825", "pay_amt": "-238"}

    一列都沒有就是不對。真的網站不管有沒有成交，都會把最近幾個交易日各列一列
    （沒成交的那幾天是 0），所以「空的」代表這支查詢沒查成，不是「沒有錢要交割」——
    當成 0 會讓還沒扣的錢憑空消失，而畫面上不會有任何徵兆。

    每一列的交割日與金額也要讀得懂：讀不懂交割日就無法判斷該不該補，
    讀不懂金額就不知道要補多少，兩種都只能整格擋住。
    """
    if not rows:
        return "交割金額查詢沒有回傳任何一天的資料"

    for row in rows:
        cdate = str(row.get("cdate") or "").strip()
        if len(cdate) != 8 or not cdate.isdigit():
            return f"交割金額查詢有一列讀不到交割日：{row}"
        if to_num(row.get("pay_amt"), None) is None:
            return f"交割金額查詢裡交割日 {cdate} 那一列讀不到金額：{row.get('pay_amt')!r}"
    return None


def bank_problem(rows, cid):
    """
    銀行餘額回應能不能用。可以就回 None，不能就回一句給人看的話。

    這支查詢的回應裡**沒有 bhno/cseq**，所以現有那道「這份資料是不是這個人的」核對
    用不上，只剩銀行帳號裡含著客戶號這個線索（1112-0108640 的帳號是 71017108640）。
    銀行餘額是單一數字，抄錯人不會有任何徵兆，所以這道再弱也要做。

    多於一筆的情況沒有人看過（可能是綁了兩個銀行帳戶），不知道該用哪一個就不要猜。
    """
    if not rows:
        return "銀行餘額查詢沒有回傳任何帳戶"
    if len(rows) > 1:
        accounts = "、".join(str(row.get("bnkacc") or "?") for row in rows)
        return f"銀行餘額查詢回傳了 {len(rows)} 個帳戶（{accounts}），不知道該用哪一個"

    account = str(rows[0].get("bnkacc") or "")
    trimmed = (cid or "").lstrip("0")
    if trimmed and trimmed not in account:
        return f"銀行帳號 {account} 裡看不到客戶號 {cid}，這筆餘額可能不是這個人的"
    return None


def new_store():
    """
    瀏覽器活著的期間要記住的東西。由呼叫方保管，跨多次 ensure_logged_in 沿用；
    瀏覽器重開就換一份新的。

        pages   第幾組 -> 那一組在用的分頁
        jars    第幾組 -> 那一組登入完當下的整份 cookie
        codes   第幾組 -> 那一組的帳號代號（1分行-客戶號），換 cookie 後要拿來對身分
        owner   現在瀏覽器帶著的是哪一組的 cookie

    不帶 store 就是最原始的行為：每次都重新登入。
    """
    return {"pages": {}, "jars": {}, "codes": {}, "owner": None}


def account_code(session):
    """sessionStorage 探到的身分 -> 帳號代號。網站的每一列資料都帶著它。"""
    bid, cid = (session or {}).get("branch_id"), (session or {}).get("cust_id")
    return f"1{bid}-{cid}" if bid and cid else ""


# 帳戶頁「可以判斷了」的條件，也就是 _open_account_page 真正在等的東西。
#
# 後半是這個分頁接下來要用的兩樣：呼叫 AJAX 用的 common.js 全域函式（見
# recon.FETCH_JS），以及認得出登入身分的 sessionStorage（見 recon.SESSION_JS）。
#
# 前半的「已經不在 /account/ 了」是給失敗那條路的：session 過期時網站會把人踢回
# 登入／首頁，那一頁永遠等不到後半那兩樣東西。沒有這一條，最該快點放棄的情況
# （踢回去了，要重登）反而要白等滿逾時；有了它就立刻回來，讓 _revisit 從網址看出來。
PAGE_READY_JS = """
() => !location.pathname.includes('/account/')
   || (typeof B64_XOR_Encode === 'function' && typeof XOR_KEY !== 'undefined'
       && !!(sessionStorage.getItem('branch_id') && sessionStorage.getItem('cust_id')))
"""


def _open_account_page(page):
    """
    導到帳戶頁，等到能判斷它可不可以用為止。等不到就算了 —— 缺什麼由呼叫方自己去看
    （網址、sessionStorage、身分核對三道都在 _revisit 裡）。

    刻意不等整頁 load、也不等 networkidle：這個網站的帳戶頁有背景請求一直在跑，
    等它安靜是在等一件我們不在乎的事。2026/08/21 實測，同一次登入裡等 networkidle
    要 0.90 秒，等 PAGE_READY_JS 只要 0.33 秒，而且等到的正好是下一步真正要用的東西
    —— 等完立刻打 AJAX，retcode 照樣是 000000。整趟帳戶頁從 2.96 秒降到 2.18 秒。
    """
    page.goto(ACCOUNT_PAGE, wait_until="domcontentloaded")
    try:
        page.wait_for_function(PAGE_READY_JS, timeout=15000)
    except (PlaywrightError, PlaywrightTimeoutError):
        pass


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
    _open_account_page(page)
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


def spare_page(context, store=None):
    """
    可以借來登入用的空白分頁：persistent context 一啟動就自帶一個，不借白不借。

    只認第一個分頁，而且只在「還沒有人認領」的時候才借。兩道限制都是必要的：

        只認第一個  後面那些分頁可能是使用者自己開的（網站的看盤視窗就是
                    window.open 開出來的），把它導去登入頁等於把人家的視窗弄掉
        沒人認領    第一次登入之後，那個分頁就是第一組帳號在用的分頁了，
                    再借給第二組等於把第一組踢掉

    一輪之內還要靠 _ensure_one 把它一路傳下去（借掉一次就不再是空的了）——
    這裡看得到的只有 store，看不到同一輪剛剛才借出去的那一次。
    """
    first = context.pages[0] if context.pages else None
    if first is None:
        return None
    claimed = {id(page) for page in (store or {}).get("pages", {}).values()}
    return None if id(first) in claimed else first


def _ensure_one(context, order, account, store, spare):
    """
    確保這一組帳號已登入。回傳 (page, session, problems, 還沒被借走的空白分頁)。

    session 是 SESSION_JS 探出來的 sessionStorage 內容，登入失敗時是 None；
    problems 非空代表這組帳號直接放棄（逾時或瀏覽器操作失敗）。

    store 見 new_store()。帶了它，同一組帳號會重複用自己的分頁，而且換人時先試
    「換回他的 cookie」，換不回去（逾時了）才重登。不帶就是每次都重新登入。

    一組帳號走到這裡有三條路，成本差很多：

        本來就是他的回合   直接回帳戶頁確認一下          最快
        換人、cookie 還在  換回他的 cookie 再確認一次    快（不必驗證碼）
        沒有 cookie 或逾時 整套登入流程跑一遍            慢

    spare 要一路傳下去、不能每組各自去拿 context.pages[0]：第一組登入完之後，
    那個「第一個分頁」就是第一組在用的分頁了，再借給第二組等於把第一組踢掉。
    """
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
            page = simulate.open_page(context, account, reuse_page or spare)
            if reuse_page is None:
                spare = None
            session = {
                "branch_id": account["branch_id"],
                "cust_id": account["cust_id"],
                "account": account["name"],
            }
            if store is not None:
                store["pages"][order] = page
            return page, session, problems, spare

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
            page = do_login(context, account["id"], account["password"], reuse_page or spare)
            spare = None
            _open_account_page(page)
            session = page.evaluate(SESSION_JS)

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

    return page, session, problems, spare


def ensure_logged_in(context, selected, store=None):
    """
    把每組帳號都登入好，回傳 {第幾組: (page, session, problems)}。給「登入」按鈕用。

    **要抓資料的話不要用這一支**（見 collect）：整個瀏覽器只有一組 cookie，全部
    登入完之後它是最後一組的，這時候回頭查前面幾組會拿到最後那一組的資料。
    """
    results = {}
    spare = spare_page(context, store)

    for order, account in selected:
        page, session, problems, spare = _ensure_one(context, order, account, store, spare)
        results[order] = (page, session, problems)

    return results


def login_only(context, selected, store=None):
    """
    只確保登入，不查資料。給「登入」按鈕用：先把瀏覽器開好、帳號登入好，
    之後按「讀取」（見 collect）就能直接重查，不必再等一次自動登入。

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

    return records


def collect(context, selected, store=None, need_bank=False):
    """
    逐一確保登入並抓資料。selected 是 [(第幾組, 帳號設定)]，回傳每組一筆記錄。

    need_bank 是「現金餘額這次要用銀行餘額推算」。只有那時候才多查銀行餘額與交割
    金額兩支 —— 20 個帳號就是 40 次多餘的往返，而用另一種算法的日子一次都用不到。

    **一組登入完就立刻抓完他的資料**，不是全部登入完再回頭抓。整個瀏覽器只有一組
    cookie，登入下一組就等於把上一組的身分換掉了（伺服器那邊的 session 還活著，
    只是沒有人帶著他的 cookie 去敲門，見 _swap_cookies）。全部登入完再回頭查，
    查到的會是最後那一組的資料 —— 每一列都帶 bhno/cseq，所以下面那道核對會擋下來，
    但那是「每次讀取都失敗」，不是「偶爾」。2026/08/21 修好，那之前只有一組真帳號，
    所以一直沒有人踩到。

    store 的用法見 new_store 與 _ensure_one。
    """
    records = []
    spare = spare_page(context, store)

    for order, account in selected:
        record = {"order": order, "problems": []}
        records.append(record)

        page, session, problems, spare = _ensure_one(context, order, account, store, spare)
        if problems:
            record["problems"].extend(problems)
            continue

        bid, cid = (session or {}).get("branch_id"), (session or {}).get("cust_id")
        if not bid or not cid:
            record["problems"].append(f"沒有登入成功（sessionStorage 沒有帳號資料），目前網址 {page.url}")
            continue

        record["account_code"] = account_code(session)
        record["sheet_name"] = (session.get("account") or "").strip()

        queries = {
            "未實現損益": ("queryInstantAccount_new", {
                "branchId": "1" + bid, "custId": cid,
                "range": "stksum,stkdat", "stock_no": "",
            }),
            "當日淨收付": ("queryInstantAccount_new", {
                "branchId": "1" + bid, "custId": cid,
                "his": "y", "queryType": "1",
                "range": "stksum,stkdat,matsum,matdat", "stock_no": "",
            }),
        }

        if need_bank:
            # 現金餘額第二種算法要的兩支（見 docs/現金餘額兩種算法.md）。
            #
            # 交割金額查詢就是網站上「交割金額查詢」那一頁（account/layoutRWD.jsp?type=3）
            # 自己打的那一支，回的正好是「接下來每一天各要交割多少」，一天一列。
            # 上一版是「撈十天的淨收付再逐筆看 cdate」算同一件事，2026/08/21 對過答案
            # 兩邊一樣（都是 -238），但這支只要一次往返、也不必自己算日期區間。
            queries["交割金額"] = ("query610", {"branchId": "1" + bid, "custId": cid})
            queries["銀行餘額"] = ("queryBankBalance",
                                   {"branchId": "1" + bid, "custId": cid})

        # 模擬帳號的資料是從它自己的假頁面即時算出來的，形狀跟真 API 的回應一樣，
        # 所以下面的 retcode 檢查、帳號回顯檢查全部照跑，不特別放水。
        simulated = None
        if account.get("fake"):
            try:
                simulated = simulate.read_page(page)
            except (PlaywrightError, RuntimeError) as exc:
                record["problems"].append(f"讀取模擬頁面失敗：{exc}")
                continue

        for title, (cmd, param_info) in queries.items():
            if simulated is not None:
                data, raw = simulated.get(title), "（模擬資料）"
            else:
                data, raw = query(page, cmd, param_info)
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

            if title in ("銀行餘額", "交割金額"):
                # 這兩支的回應形狀跟其他的不一樣（data 不是 arrays），而且裡面都沒有
                # bhno/cseq，上面那道核對等於沒做 —— 所以另外檢查一次。
                #
                # 交割金額連「銀行帳號裡含著客戶號」這種弱核對都沒有，身分完全靠
                # 同一輪的當日淨收付撐著：那支每一列都有 bhno/cseq，對不上就會被
                # 上面那道擋下來、record 裡不會有「當日淨收付」，而現金那一格
                # 沒有它就整格不寫（見 planner._cash_blocked）。所以 session 被別人
                # 頂掉時，這支就算悄悄回了別人的資料，也走不到 Excel 上。
                rows = data.get("data") or []
                problem = (bank_problem(rows, cid) if title == "銀行餘額"
                           else settle_problem(rows))
                if problem:
                    record["problems"].append(problem)
                    continue
                record[title] = rows
                continue

            record[title] = data.get("arrays") or []

    return records
