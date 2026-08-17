"""
登入並抓回「未實現損益」與「當日淨收付」的原始資料。只讀網頁，不碰 Excel。

抓資料的方式是在已登入的分頁裡呼叫網站自己的 AJAX（POST /tbb/MainController），
沿用頁面現成的 B64_XOR_Encode / XOR_KEY 與 cookie，不做畫面解析 —— 表格是 JS
動態產生的，解析畫面既慢又容易被改版弄壞。
"""

from playwright.sync_api import Error as PlaywrightError, TimeoutError as PlaywrightTimeoutError

from login import do_login
from recon import ACCOUNT_PAGE, SESSION_JS, account_codes, query


def collect(context, selected):
    """
    逐一登入並抓資料。selected 是 [(第幾組, 帳號設定)]，回傳每組一筆記錄。

    刻意「登入完立刻抓」而不是全部登入完再回頭抓：所有分頁共用同一個 JSESSIONID，
    後登入的會把前一個的 server session 頂掉，晚一步抓可能拿到別人的資料。
    """
    records = []
    spare_page = context.pages[0] if context.pages else None

    for order, account in selected:
        record = {"order": order, "problems": []}
        records.append(record)

        try:
            page = do_login(context, account["id"], account["password"], spare_page)
            spare_page = None
            page.goto(ACCOUNT_PAGE)
            try:
                page.wait_for_load_state("networkidle", timeout=15000)
            except PlaywrightTimeoutError:
                pass
        except PlaywrightTimeoutError:
            record["problems"].append("登入逾時，找不到欄位，網站版面可能已變更")
            continue
        except PlaywrightError as exc:
            record["problems"].append(f"瀏覽器操作失敗：{exc}")
            continue

        session = page.evaluate(SESSION_JS)
        bid, cid = session.get("branch_id"), session.get("cust_id")
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

        for title, param_info in queries.items():
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
