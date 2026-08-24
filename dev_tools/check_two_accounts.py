# -*- coding: utf-8 -*-
"""
驗證「多帳號讀取時，每一組拿到的是自己的資料」—— 用假瀏覽器跑，不連網、不碰 Excel。

    python dev_tools/check_two_accounts.py

為什麼需要這一支：整個瀏覽器只有一組 cookie，所以抓資料的那一刻，瀏覽器帶著的
必須正好是那一組的 cookie（見 fetch.collect 的 docstring）。這件事寫錯的話，
只有一組真帳號時完全看不出來 —— 前面沒有別人可以頂掉他 —— 20 組上線的第一天
才會每次讀取都壞。2026/08/21 真的踩過一次：collect() 當時是先把全部登入完
才回頭逐一抓，前面 19 組拿到的都是最後那一組的資料。

假伺服器是這支腳本的重點，它照真的規則回話：

    看 cookie 是誰的就回誰的資料    整個 context 只有一組 cookie
    分頁的 sessionStorage 不會變    換 cookie 不影響它（fetch._revisit 講的那個陷阱）

所以「順序寫錯」在這裡一定會被抓到，不必等到真的有 20 組帳號。

跑四種情境，每一種都要求「每組拿到自己的三份資料、沒有任何 problem」：

    第一次讀取        兩組都要重新登入
    第二次讀取        cookie 還在，一次都不該重登（重登＝慢、又要驗證碼）
    只更新某一位      要先把他的 cookie 換回來
    先登入再讀取      「登入」按鈕先把全部登好，「讀取」時每組要先換回自己的 cookie

最後再看一次分頁有沒有被兩組帳號共用（借空白分頁那條路寫錯就會，見 fetch.spare_page）。
"""

import json
import sys
from pathlib import Path

# 直接用 `python dev_tools/check_two_accounts.py` 執行時，sys.path 起點是
# dev_tools 這一層，補回專案根目錄才找得到 fetch / recon。
_ROOT = Path(__file__).resolve().parent.parent
if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))

import fetch
import recon

# 假帳號設定，形狀跟 .env 讀出來的真帳號一樣（沒有 fake 旗標 —— 這裡要測的正是
# 真帳號那條路：登入、換 cookie、身分核對。模擬帳號根本不碰 cookie）。
ACCOUNTS = [
    {"id": "A1234", "password": "x", "branch_id": "112", "cust_id": "0100001"},
    {"id": "B5678", "password": "y", "branch_id": "112", "cust_id": "0100002"},
]

_server = {"owner": None}     # 瀏覽器現在帶著誰的 cookie
_events = []                  # 依序記下發生過什麼（登入、查詢）


def _payload(cmd, owner):
    """假伺服器：不管是誰的分頁在問，一律回「cookie 現在是誰的」那個人的資料。"""
    bid, cid = "1" + owner["branch_id"], owner["cust_id"]
    if cmd == "query610":
        return {"retcode": "000000",
                "data": [{"trade": "20260821", "cdate": "20260825", "pay_amt": "-238"}]}
    if cmd == "queryBankBalance":
        # bnkacc 隨便給一個能通過 fetch.bank_problem 基本檢查（有回傳、只有一筆）的值即可，
        # 2026/08/24 起 fetch.bank_problem 已經不核對客戶號是否含在銀行帳號裡。
        return {"retcode": "000000",
                "data": [{"Amount": "0000000089300", "bnkacc": "7101" + cid}]}
    return {"retcode": "000000",
            "arrays": [{"bhno": bid, "cseq": cid, "stkno": "", "matdat": []}]}


class FakePage:
    """一個登入完的分頁。它記得自己是誰登入的 —— 那就是它的 sessionStorage。"""

    def __init__(self, owner):
        self.owner = owner
        self.url = recon.ACCOUNT_PAGE

    def is_closed(self):
        return False

    def goto(self, url, **_kwargs):
        self.url = url

    def wait_for_function(self, *_args, **_kwargs):
        return True

    def evaluate(self, js, arg=None):
        if js == recon.SESSION_JS:
            # 換 cookie 不會動到分頁自己的 sessionStorage，所以這裡照樣回原本那個人。
            return {"branch_id": self.owner["branch_id"],
                    "cust_id": self.owner["cust_id"],
                    "account": self.owner["id"]}
        if js == fetch.PAGE_READY_JS:
            return True
        if js == recon.FETCH_JS:
            _events.append(f"查 {self.owner['id']} 的分頁 {arg['cmd']}"
                           f"（cookie 是 {_server['owner']['id']} 的）")
            return {"status": 200, "text": json.dumps(_payload(arg["cmd"], _server["owner"]))}
        raise AssertionError(f"沒有預期到的 evaluate：{js[:60]}")


class FakeContext:
    """只做 fetch 會用到的那幾件事：借分頁、收 cookie、換 cookie。"""

    def __init__(self):
        self.pages = []

    def new_page(self):
        raise AssertionError("這個測試裡不該開新分頁（分頁都是 do_login 開的）")

    def cookies(self):
        return [{"name": "JSESSIONID", "value": _server["owner"]["id"]}]

    def clear_cookies(self):
        _server["owner"] = None

    def add_cookies(self, jar):
        _server["owner"] = next(a for a in ACCOUNTS if a["id"] == jar[0]["value"])


def _fake_login(_context, tbb_id, _password, page=None):
    """
    代替 login.do_login。fetch 是 `from login import do_login` 進來的，
    所以要換掉的是 fetch 自己命名空間裡那一個。
    """
    account = next(a for a in ACCOUNTS if a["id"] == tbb_id)
    _events.append(f"登入 {tbb_id}")
    _server["owner"] = account
    return FakePage(account)


def _check(title, records, accounts, no_login=False):
    """印出這一輪的經過，回傳有沒有通過。"""
    print(f"--- {title} ---")
    for line in _events:
        print("    " + line)

    ok = True
    if no_login and any(line.startswith("登入") for line in _events):
        print("    !! 不該重登卻重登了（換 cookie 那條路沒走到，會慢、又要驗證碼）")
        ok = False
    _events.clear()

    for record, account in zip(records, accounts):
        want = f"1{account['branch_id']}-{account['cust_id']}"
        got = [k for k in record
               if k not in ("order", "problems", "account_code", "sheet_name")]
        good = (not record["problems"] and len(got) == 3
                and record.get("account_code") == want)
        print(f"    第 {record['order']} 組 {account['id']}："
              f"身分 {record.get('account_code') or '（沒有）'}、抓到 {len(got)} 份"
              f" -> {'OK' if good else '!! 不對'}")
        for problem in record["problems"]:
            print(f"        {problem}")
        ok = ok and good

    print()
    return ok


def main():
    fetch.do_login = _fake_login
    numbered = list(enumerate(ACCOUNTS, start=1))
    passed = []

    context, store = FakeContext(), fetch.new_store()
    passed.append(_check("第一次讀取（兩組都要重新登入）",
                         fetch.collect(context, numbered, store, need_bank=True), ACCOUNTS))
    passed.append(_check("第二次讀取（cookie 還在，一次都不該重登）",
                         fetch.collect(context, numbered, store, need_bank=True), ACCOUNTS,
                         no_login=True))
    passed.append(_check("只更新第 1 組（要先把他的 cookie 換回來）",
                         fetch.collect(context, numbered[:1], store, need_bank=True),
                         ACCOUNTS[:1], no_login=True))

    # 「登入」按鈕的流程：先把全部登入好，之後按「讀取」才抓資料。
    context2, store2 = FakeContext(), fetch.new_store()
    fetch.login_only(context2, numbered, store2)
    _events.clear()
    passed.append(_check("先按登入、再按讀取",
                         fetch.collect(context2, numbered, store2, need_bank=True), ACCOUNTS,
                         no_login=True))

    pages = {order: id(page) for order, page in store["pages"].items()}
    shared = len(set(pages.values())) != len(pages)
    print(f"--- 分頁有沒有被共用 --- {'!! 有兩組帳號共用同一個分頁' if shared else 'OK：各用各的'}")
    passed.append(not shared)

    print()
    if all(passed):
        print("全部通過：每一組拿到的都是自己的資料。")
        return 0
    print("有情境沒過。抓資料的那一刻，瀏覽器帶著的 cookie 一定要是那一組的 —— "
          "看上面的經過，是不是又變成「全部登入完才回頭抓」了。")
    return 1


if __name__ == "__main__":
    sys.exit(main())
