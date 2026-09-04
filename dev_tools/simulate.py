"""
用假帳號模擬多帳號的情境：只有一組真帳號，也能測 20 個分頁的流程。

為什麼要有這支程式
------------------
真正要跑的規模是 20 組帳號、Excel 20 個分頁，但手上只有一組真帳號。
介面、紀錄檔、現金帳本在這個規模下順不順，不能等到拿到 20 組帳密才知道。

做法是「真的那一組照原本的流程登入，其餘 19 個在同一個瀏覽器開分頁，
塞一頁自己畫的假網頁」。假網頁不連任何網站（page.set_content，沒有任何請求），
資料放在頁面的 window.__SIM__ 這個函式裡，fetch.py 對假帳號改成呼叫它，
而不是打 MainController。回傳的形狀跟真的 API 一模一樣，所以 planner、ledger、
excel_io、ui 完全不必知道有假帳號這件事。

模擬不到的是什麼
----------------
20 個分頁裡只有 1 個真的在跟網站講話，所以「多帳號共用 JSESSIONID 會互踢」
這個真正的風險測不到 —— 那只有真的 20 組帳號才驗得出來。

假網頁上的數字可以直接改
------------------------
股數、成交均價、市價、當日成交都是輸入格，改完馬上生效（讀取時是即時從輸入格
算出來的，不必按任何「套用」按鈕，也就沒有「改了忘記按」的陷阱）。
這是刻意的：要測特定情境（某檔賣光、淨收付為 0、成本被改動…）用改頁面最快。
股數改成 0 就等於網頁庫存沒有這一檔，可以測 planner 的「網頁沒有」那條路。

分頁被關掉就重新產生一頁；還開著就沿用，不會把你改過的數字洗掉。
"""

import json
import os
import random

from playwright.sync_api import Error as PlaywrightError

# 假帳號的名字。Excel 的分頁名稱要跟這個一模一樣（分頁名 = 帳戶名，見 excel_io.find_sheet）。
FAKE_NAMES = [f"交易人{letter}" for letter in "ABCDEFGHIJKLMNOPQRSTUVWXYZ"]

# 模擬用的股票：代號、名稱、基準價。基準價只是隨機的中心點，不代表真實股價。
STOCKS = [
    ("2059", "川湖", 1150),
    ("2344", "華邦電", 25),
    ("2408", "南亞科", 65),
    ("6213", "聯茂", 130),
    ("6274", "台燿", 200),
]

# ============================================================
# 固定模擬帳號資料
# ============================================================

FIXED_ACCOUNTS = {
    "交易人A": {
        "cash": 80_000,
        "original_capital": 200_000,
        "holdings": [
            ("0050", "元大台灣50", 300, 104.60, 300),
            ("006208", "富邦台50", 200, 239.10, 200),
        ],
        "trades": [],
    },

    "交易人B": {
        "cash": 40_000,
        "original_capital":200,
        "holdings": [
            ("0050", "元大台灣50", 300, 104.60, 300),
            ("006208", "富邦台50", 200, 239.10, 200),
        ],
        "trades": [],
    },

    "交易人C": {
        "cash": 1_354_628,
        "original_capital": 956_000,
        "holdings": [
            ("2059", "川湖", 3_000, 1070.0, 1150),
            ("2344", "華邦電", 78_000, 26.7, 25),
            ("2408", "南亞科", 64_000, 68.1, 65),
            ("6213", "聯茂", 5_000, 147.0, 130),
            ("6274", "台燿", 19_000, 175.5, 200),
        ],
        "trades": [],
    },

    "交易人D": {
        "cash": 2_016_578,
        "original_capital": 3_367_000,
        "holdings": [
            ("2059", "川湖", 2_000, 1322.0, 1150),
            ("2344", "華邦電", 38_000, 28.6, 25),
            ("2408", "南亞科", 9_000, 57.8, 65),
            ("6213", "聯茂", 14_000, 122.5, 130),
            ("6274", "台燿", 14_000, 191.5, 200),
        ],
        "trades": [],
    },

    "交易人E": {
        "cash": 5_021_213,
        "original_capital": 2_774_000,
        "holdings": [
            ("2059", "川湖", 3_000, 1311.0, 1150),
            ("2344", "華邦電", 211_000, 23.3, 25),
            ("2408", "南亞科", 59_000, 58.2, 65),
            ("6213", "聯茂", 43_000, 113.0, 130),
            ("6274", "台燿", 11_000, 192.0, 200),
        ],
        "trades": [],
    },
        "交易人F": {
        "cash": 3_612_644,
        "original_capital": 2_118_000,
        "holdings": [
            ("2059", "川湖", 4_000, 1128.0, 1150),
            ("2344", "華邦電", 136_000, 27.9, 25),
            ("2408", "南亞科", 55_000, 71.9, 65),
            ("6213", "聯茂", 35_000, 141.5, 130),
            ("6274", "台燿", 5_000, 180.0, 200),
        ],
        "trades": [],
    },

    "交易人G": {
        "cash": 1_939_060,
        "original_capital": 1_360_000,
        "holdings": [
            ("2059", "川湖", 2_000, 1136.0, 1150),
            ("2344", "華邦電", 98_000, 21.5, 25),
            ("2408", "南亞科", 26_000, 59.0, 65),
            ("6213", "聯茂", 11_000, 139.5, 130),
            ("6274", "台燿", 5_000, 228.0, 200),
        ],
        "trades": [],
    },
    "交易人H": {
        "cash": 8_767_959,
        "original_capital": 4_484_000,
        "holdings": [
            ("2059", "川湖", 2_000, 1292.0, 1150),
            ("2344", "華邦電", 77_000, 23.8, 25),
            ("2408", "南亞科", 10_000, 62.8, 65),
            ("6213", "聯茂", 15_000, 117.5, 130),
            ("6274", "台燿", 19_000, 223.0, 200),
        ],
        "trades": [],
    },

    "交易人I": {
        "cash": -7_936_722,
        "original_capital": 614_000,
        "holdings": [
            ("2059", "川湖", 1_000, 1092.0, 1150),
            ("2344", "華邦電", 161_000, 26.7, 25),
            ("2408", "南亞科", 79_000, 62.6, 65),
            ("6213", "聯茂", 11_000, 146.0, 130),
            ("6274", "台燿", 4_000, 199.0, 200),
        ],
        "trades": [],
    },

    "交易人J": {
        "cash": 4_429_567,
        "original_capital": 4_453_000,
        "holdings": [
            ("2059", "川湖", 2_000, 1002.0, 1150),
            ("2344", "華邦電", 162_000, 22.6, 25),
            ("2408", "南亞科", 18_000, 65.7, 65),
            ("6213", "聯茂", 33_000, 117.0, 130),
            ("6274", "台燿", 10_000, 193.5, 200),
        ],
        "trades": [],
    },

    "交易人K": {
        "cash": -1_753_676,
        "original_capital": 1_582_000,
        "holdings": [
            ("2059", "川湖", 4_000, 1171.0, 1150),
            ("2344", "華邦電", 145_000, 28.1, 25),
            ("2408", "南亞科", 35_000, 66.1, 65),
            ("6213", "聯茂", 6_000, 149.5, 130),
            ("6274", "台燿", 16_000, 218.5, 200),
        ],
        "trades": [],
    },

    "交易人L": {
        "cash": 1_954_171,
        "original_capital": 2_537_000,
        "holdings": [
            ("2059", "川湖", 2_000, 1013.0, 1150),
            ("2344", "華邦電", 47_000, 28.4, 25),
            ("2408", "南亞科", 52_000, 65.3, 65),
            ("6213", "聯茂", 6_000, 120.0, 130),
            ("6274", "台燿", 5_000, 205.0, 200),
        ],
        "trades": [],
    },

    "交易人M": {
        "cash": 8_385_307,
        "original_capital": 2_998_000,
        "holdings": [
            ("2059", "川湖", 3_000, 1242.0, 1150),
            ("2344", "華邦電", 181_000, 22.0, 25),
            ("2408", "南亞科", 58_000, 58.9, 65),
            ("6213", "聯茂", 15_000, 113.5, 130),
            ("6274", "台燿", 4_000, 190.0, 200),
        ],
        "trades": [],
    },

    "交易人N": {
        "cash": 2_706_115,
        "original_capital": 3_328_000,
        "holdings": [
            ("2059", "川湖", 4_000, 1081.0, 1150),
            ("2344", "華邦電", 150_000, 24.0, 25),
            ("2408", "南亞科", 60_000, 73.2, 65),
            ("6213", "聯茂", 5_000, 121.0, 130),
            ("6274", "台燿", 13_000, 210.5, 200),
        ],
        "trades": [],
    },

    "交易人O": {
        "cash": 2_431_457,
        "original_capital": 504_000,
        "holdings": [
            ("2059", "川湖", 2_000, 1241.0, 1150),
            ("2344", "華邦電", 66_000, 21.7, 25),
            ("2408", "南亞科", 55_000, 67.3, 65),
            ("6213", "聯茂", 30_000, 139.5, 130),
            ("6274", "台燿", 21_000, 194.0, 200),
        ],
        "trades": [],
    },

    "交易人P": {
        "cash": 4_718_000,
        "original_capital": 4_718_000,
        "holdings": [
            ("2059", "川湖", 1_000, 1108.0, 1150),
            ("2344", "華邦電", 124_000, 23.3, 25),
            ("2408", "南亞科", 66_000, 65.2, 65),
            ("6213", "聯茂", 28_000, 135.5, 130),
            ("6274", "台燿", 17_000, 208.5, 200),
        ],
        "trades": [],
    },

    "交易人Q": {
        "cash": 3_652_000,
        "original_capital": 3_652_000,
        "holdings": [
            ("2059", "川湖", 1_000, 1156.0, 1150),
            ("2344", "華邦電", 69_000, 22.3, 25),
            ("2408", "南亞科", 10_000, 58.9, 65),
            ("6213", "聯茂", 7_000, 142.0, 130),
            ("6274", "台燿", 22_000, 189.5, 200),
        ],
        "trades": [],
    },

    "交易人R": {
        "cash": 1_814_371,
        "original_capital": 773_000,
        "holdings": [
            ("2059", "川湖", 4_000, 1184.0, 1150),
            ("2344", "華邦電", 168_000, 24.7, 25),
            ("2408", "南亞科", 15_000, 67.2, 65),
            ("6213", "聯茂", 11_000, 120.5, 130),
            ("6274", "台燿", 7_000, 184.5, 200),
        ],
        "trades": [],
    },

    "交易人S": {
        "cash": 3_274_494,
        "original_capital": 4_017_000,
        "holdings": [
            ("2059", "川湖", 1_000, 1316.0, 1150),
            ("2344", "華邦電", 155_000, 22.1, 25),
            ("2408", "南亞科", 63_000, 64.0, 65),
            ("6213", "聯茂", 10_000, 119.5, 130),
            ("6274", "台燿", 6_000, 185.0, 200),
        ],
        "trades": [],
    },
}

# 假帳號的分公司代碼與客戶號。用 999 / 9 開頭是刻意的：畫面上一眼就看得出不是真帳號
# （真的長得像 1112-0108640）。
FAKE_BRANCH = "999"

# 一個假帳號在 Excel B8 的起始現金範圍（元）。
CASH_MIN, CASH_MAX = 500_000, 5_000_000

# 一檔持股的市值大概落在這個範圍，張數由它跟股價回推 —— 直接亂數決定張數的話，
# 川湖（千元股）跟華邦電（二十幾元）會差到很誇張，表格看起來就不像真的。
POSITION_MIN, POSITION_MAX = 500_000, 5_000_000


def simulate_count(raw=None):
    """
    要模擬幾個假帳號。.env 的 SIMULATE_ACCOUNTS，沒設或設 0 就完全沒有假帳號。

    預設關閉是重點：正式部署的機器上這個變數不存在，整套模擬就等於不存在，
    不可能因為忘記關而把假資料寫進真的 Excel。

    raw 有傳就用那個值、不看 os.environ —— 給「直接讀 .env 檔案」的那條路用
    （見 login.accounts_on_disk），假帳號的數量改了也要算進「.env 變了沒」。
    """
    if raw is None:
        raw = os.getenv("SIMULATE_ACCOUNTS", "")
    raw = raw.strip()
    if not raw:
        return 0
    try:
        count = int(raw)
    except ValueError:
        return 0
    return max(0, min(count, len(FAKE_NAMES)))


def fake_accounts(count=None):
    """
    產生假帳號設定，形狀跟 load_accounts() 的真帳號一樣，只是多帶 fake 旗標。

    id 不是身分證字號而是「(模擬)交易人A」，萬一哪裡不小心印出來也看得出是假的。
    """
    if count is None:
        count = simulate_count()

    accounts = []
    for index, name in enumerate(FAKE_NAMES[:count], start=1):
        accounts.append({
            "id": f"(模擬){name}",
            "password": "",
            "fake": True,
            "name": name,
            "branch_id": FAKE_BRANCH,
            "cust_id": f"90000{index:02d}",
        })
    return accounts


def account_code(account):
    """假帳號的帳號代碼，格式跟真的一致（1 + 分公司 + 客戶號）。"""
    return f"1{account['branch_id']}-{account['cust_id']}"


def initial_cash(name):
    """取得固定帳號的原始資金。"""

    account = FIXED_ACCOUNTS.get(name)

    if account:
        return account["original_capital"]

    return 0


def bank_balance(name):
    """假帳號的銀行餘額。用現金餘額當預設值，理由見 render_html 裡的註解。"""
    return (FIXED_ACCOUNTS.get(name) or {}).get("cash", 0)


def tick_size(price):
    """台股的最小跳動單位。假資料也照這個規則，畫面上才不會出現 63.4271 這種價位。"""
    if price < 10:
        return 0.01
    if price < 50:
        return 0.05
    if price < 100:
        return 0.1
    if price < 500:
        return 0.5
    return 1.0


def tick_round(price):
    unit = tick_size(price)
    return round(round(price / unit) * unit, 2)


def seed_data(name, day):
    """
    取得固定模擬帳號資料。

    不再依照日期或亂數產生資料，
    每個帳號每天看到的資料都完全一樣。
    """

    account = FIXED_ACCOUNTS.get(name)

    if account is None:
        return {
            "holdings": [],
            "trades": [],
        }

    holdings = []

    for code, stock_name, qty, price, mkt in account["holdings"]:
        holdings.append({
            "code": code,
            "name": stock_name,
            "base": price,
            "qty": qty,
            "price": price,
            "mkt": mkt,
        })

    return {
        "holdings": holdings,
        "trades": account.get("trades", []),
    }


def open_page(context, account, page=None):
    """
    確保這個假帳號有一個假網頁分頁，回傳該分頁。

    已經是我們畫的那一頁（window.__SIM__ 在）就直接沿用，不重新產生 ——
    否則每次按「讀取」都會把使用者在頁面上改過的數字洗掉。
    """
    if page is None:
        page = context.new_page()

    try:
        ready = page.evaluate("() => typeof window.__SIM__ === 'function'")
    except PlaywrightError:
        ready = False

    if not ready:
        import datetime

        page.set_content(render_html(account, seed_data(account["name"], datetime.date.today())))

    return page


def read_page(page):
    """
    從假網頁讀回每一支查詢的資料，形狀跟真的 API 回應一致
    （retcode / retmsg，加上 arrays 或 data —— 交割金額與銀行餘額是 data）。

    即時從頁面上的輸入格算出來，所以看到什麼就讀到什麼。
    """
    data = page.evaluate("() => (typeof window.__SIM__ === 'function') ? window.__SIM__() : null")
    if not data:
        raise RuntimeError("這個分頁不是模擬頁面（找不到 window.__SIM__），可能被導去別的網址了")
    return data


def render_html(account, data):
    """畫出假網頁。資料與計算全在頁面裡，改輸入格就即時生效。"""
    meta = {
        "name": account["name"],
        # 資料列裡的 bhno 含開頭那個 1（真的長得像 1112），而 sessionStorage 的
        # branch_id 沒有（112）—— fetch.py 就是拿 "1"+branch_id 去比對每一列的 bhno，
        # 所以這裡也得帶那個 1，否則每組模擬帳號都會被判成「session 被別人頂掉」。
        "bhno": "1" + account["branch_id"],
        "cseq": account["cust_id"],
        "code": account_code(account),
        "cash": initial_cash(account["name"]),
        # 銀行餘額的預設值刻意等於這個帳號的現金餘額：兩種算法算出來就會一樣，
        # 畫面上不會平白冒出差額。要測「兩種算法對不上」的情境，就改這一格，
        # 或把某一筆成交改成「已交割」。
        "bank": bank_balance(account["name"]),
        # 真的那個是 71017108640（含客戶號 0108640）；2026/08/24 起
        # fetch.bank_problem 不再拿客戶號去比對這欄，這裡沿用同樣格式只是圖方便。
        "bnkacc": "7101" + account["cust_id"],
    }

    rows = "\n".join(
        f'<tr data-code="{h["code"]}" data-name="{h["name"]}" data-base="{h["base"]}">'
        f'<td class="stk">{h["name"]}<span class="no">{h["code"]}</span></td>'
        f'<td><input class="qty num" value="{h["qty"]}"></td>'
        f'<td><input class="price num" value="{h["price"]:g}"></td>'
        f'<td class="amount"></td>'
        f'<td><input class="mkt num" value="{h["mkt"]:g}"></td>'
        f'<td class="value"></td><td class="pnl"></td><td class="rate"></td></tr>'
        for h in data["holdings"]
    )

    trades = json.dumps(data["trades"], ensure_ascii=False)
    names = json.dumps({code: name for code, name, _base in STOCKS}, ensure_ascii=False)

    return f"""<!DOCTYPE html>
<html lang="zh-Hant"><head><meta charset="utf-8">
<title>{meta['name']}（模擬）</title>
<style>
  body {{ font-family: "Microsoft JhengHei UI", "Microsoft JhengHei", sans-serif;
         margin: 0; padding: 0 0 40px; background: #f6f7f9; color: #202124; }}
  .banner {{ background: #b3261e; color: #fff; padding: 10px 18px; font-weight: bold; }}
  .banner small {{ font-weight: normal; opacity: .9; margin-left: 10px; }}
  header {{ padding: 14px 18px 0; }}
  h1 {{ font-size: 20px; margin: 0 0 4px; }}
  .meta {{ color: #5f6368; font-size: 13px; }}
  section {{ margin: 18px; background: #fff; border: 1px solid #dadce0; border-radius: 8px; padding: 14px 16px; }}
  h2 {{ font-size: 15px; margin: 0 0 10px; }}
  table {{ border-collapse: collapse; width: 100%; font-size: 14px; }}
  th, td {{ border-bottom: 1px solid #e8eaed; padding: 6px 8px; text-align: right; }}
  th {{ background: #f1f3f4; color: #3c4043; font-weight: normal; white-space: nowrap; }}
  th:first-child, td:first-child, .stk {{ text-align: left; }}
  .no {{ color: #80868b; margin-left: 6px; font-size: 12px; }}
  input {{ width: 96px; text-align: right; font: inherit; padding: 3px 5px;
          border: 1px solid #c4c7c5; border-radius: 4px; background: #fffef2; }}
  input.code {{ width: 70px; text-align: left; }}
  select {{ font: inherit; padding: 3px; }}
  .neg {{ color: #b3261e; }}
  .pos {{ color: #1a7f37; }}
  .total {{ font-weight: bold; font-size: 15px; margin-top: 10px; }}
  button {{ font: inherit; padding: 5px 12px; margin-right: 8px; cursor: pointer;
           border: 1px solid #c4c7c5; border-radius: 4px; background: #fff; }}
  .hint {{ color: #5f6368; font-size: 12px; margin-top: 8px; line-height: 1.7; }}
</style></head><body>

<div class="banner">模擬資料<small>這一頁不是券商網站，是程式自己畫出來的假頁面，只用來測試多帳號流程</small></div>

<header>
  <h1>{meta['name']}</h1>
  <div class="meta">帳號代碼 {meta['code']}
    起始現金（Excel B8）{meta['cash']:,}
    日期 <span id="today"></span>
    交割日 <span id="settle"></span></div>
</header>

<section>
  <h2>銀行餘額</h2>
  <div>銀行帳號 {meta['bnkacc']}　餘額（元）
    <input id="bank" class="num" value="{meta['bank']}"></div>
  <div class="hint">
    現金餘額的第二種算法用的就是這個數字：<b>銀行餘額 + 淨收付(T+0) + 淨收付(T+1)</b>。<br>
    預設值等於這個帳號的現金餘額，所以兩種算法會算出一樣的結果。
    要測「兩種對不上」（真實情境是全額交割），就改這一格，
    或把下面某一筆成交改成「已交割」。
  </div>
</section>

<section>
  <h2>未實現損益</h2>
  <table id="pnl"><thead><tr>
    <th>股票</th><th>成交股數</th><th>成交均價</th><th>成交金額</th>
    <th>市價</th><th>市值</th><th>損益</th><th>報酬率</th>
  </tr></thead><tbody>
{rows}
  </tbody></table>
  <div class="hint">
    黃色格子都可以直接改，改完立刻生效（不用按任何按鈕）。<br>
    股數改成 0 = 這一檔在網頁庫存裡消失，可以測「Excel 有這一列、網頁沒有」那條路。
  </div>
</section>

<section>
  <h2>成交明細</h2>
  <table id="mat"><thead><tr>
    <th>代號</th><th>買賣</th><th>股數</th><th>單價</th>
    <th>價金</th><th>手續費</th><th>交易稅</th><th>淨收付</th><th></th>
  </tr></thead><tbody></tbody></table>
  <div class="total">淨收付合計：<span id="net">0</span></div>
  <div class="hint">
    這張表餵給未實現損益（股數/成本）跟交割金額查詢（今日淨收付、還沒交割的錢）
    兩支，是唯一的資料來源。全部刪掉 = 今天沒有成交，淨收付 0。<br>
    「新增一筆 T+1 成交」= 這一筆算成<b>前一個交易日</b>成交、明天才交割，錢還沒扣，
    跟一般成交（今天成交、T+2 交割）一樣都算「還沒交割的錢」——但交割金額查詢會分成
    兩列，用來測<b>銀行餘額推算</b>要把好幾天的還沒交割的錢加起來，不能只看今天那一列。<br>
    注意：未實現損益那邊只要有今天的成交明細（就是這張表的內容），交割金額查詢裡
    卻沒有今天那一列或今天金額是 0，程式會刻意不動現金那一格 —— 那是防「收盤結帳後
    查不到當日資料」的保護。
  </div>
  <div style="margin-top:10px">
    <button id="add">新增一筆成交</button>
    <button id="addT1">新增一筆 T+1 成交</button>
  </div>
</section>

<script>
const META = {json.dumps(meta, ensure_ascii=False)};
const STOCK_NAMES = {names};
const INIT_TRADES = {trades};

const today = new Date();
const STAMP = today.getFullYear() +
    String(today.getMonth() + 1).padStart(2, '0') +
    String(today.getDate()).padStart(2, '0');
document.getElementById('today').textContent = STAMP;

// 交割日：往後兩個交易日（碰到週末再往後挪）。真的網站是 T+2，
// 假頁面也照著給，方法二那條「cdate 比今天晚才要補」才測得到。
function stampOf(d) {{
  return d.getFullYear() + String(d.getMonth() + 1).padStart(2, '0') +
      String(d.getDate()).padStart(2, '0');
}}
function shiftDays(from, steps) {{
  const d = new Date(from);
  const step = steps < 0 ? -1 : 1;
  for (let left = Math.abs(steps); left > 0; ) {{
    d.setDate(d.getDate() + step);
    const week = d.getDay();
    if (week !== 0 && week !== 6) left -= 1;
  }}
  return d;
}}
const SETTLE = stampOf(shiftDays(today, 2));
// 往前兩個交易日。標成「已交割」的那幾筆算成那一天成交的：今天成交的錢要 T+2
// 才扣得掉，所以「今天成交、又已經交割」在真實世界裡不存在，假頁面也不該生出來
// （程式有一道專門擋這種矛盾，見 planner._cash_blocked）。
const PREV = stampOf(shiftDays(today, -2));
// T+1：標成「T+1」的那幾筆算成前一個交易日成交，明天才交割。
const NEXT1 = stampOf(shiftDays(today, 1));
const PREV1 = stampOf(shiftDays(today, -1));
document.getElementById('settle').textContent = SETTLE;

function tickUnit(p) {{
  if (p < 10) return 0.01;
  if (p < 50) return 0.05;
  if (p < 100) return 0.1;
  if (p < 500) return 0.5;
  return 1;
}}
function tick(p) {{ const u = tickUnit(p); return Math.round(Math.round(p / u) * u * 100) / 100; }}
function num(el) {{ return el ? (parseFloat(el.value) || 0) : 0; }}
function money(n) {{ return Math.round(n).toLocaleString('en-US'); }}
// 只加減顏色的 class，不能整個蓋掉 className —— 這幾格身上還有 pnl / rate / payn，
// 一蓋掉，下一次 refresh 就再也找不到它們（第一次改數字就整張表不動了）。
function sign(el, n) {{
  el.classList.remove('neg', 'pos');
  if (n) el.classList.add(n < 0 ? 'neg' : 'pos');
}}

// ---- 持股 ----

function holdings() {{
  return [...document.querySelectorAll('#pnl tbody tr')].map(tr => ({{
    code: tr.dataset.code, name: tr.dataset.name,
    qty: num(tr.querySelector('.qty')),
    price: num(tr.querySelector('.price')),
    mkt: num(tr.querySelector('.mkt')),
  }}));
}}

// ---- 成交 ----

function tradeRow(trade) {{
  const tr = document.createElement('tr');
  // 交割狀態不給選，靠「新增一筆成交」跟「新增一筆 T+1 成交」兩顆按鈕決定，
  // 存在 dataset 裡（不是畫面上的欄位）。
  tr.dataset.settle = trade.settled === true ? '1' : (trade.settled || '0');
  tr.innerHTML =
    '<td><input class="code" value="' + trade.code + '"></td>' +
    '<td><select class="bs"><option value="B">買</option><option value="S">賣</option></select></td>' +
    '<td><input class="tqty num" value="' + trade.qty + '"></td>' +
    '<td><input class="tprice num" value="' + trade.price + '"></td>' +
    '<td class="priceqty"></td><td class="fee"></td><td class="tax"></td><td class="payn"></td>' +
    '<td><button class="del">刪除</button></td>';
  tr.querySelector('.bs').value = trade.bs;
  tr.querySelector('.del').addEventListener('click', () => {{ tr.remove(); refresh(); }});
  document.querySelector('#mat tbody').appendChild(tr);
}}

function trades() {{
  return [...document.querySelectorAll('#mat tbody tr')].map(tr => {{
    const code = tr.querySelector('.code').value.trim();
    const trade = {{
      row: tr, code: code, name: STOCK_NAMES[code] || code,
      bs: tr.querySelector('.bs').value,
      qty: num(tr.querySelector('.tqty')),
      price: num(tr.querySelector('.tprice')),
      settled: tr.dataset.settle,
    }};
    const amount = Math.round(trade.qty * trade.price);
    trade.priceqty = amount;
    trade.fee = amount ? Math.max(1, Math.round(amount * 0.001425)) : 0;
    trade.tax = trade.bs === 'S' ? Math.round(amount * 0.003) : 0;
    trade.payn = trade.bs === 'B' ? -(amount + trade.fee) : amount - trade.fee - trade.tax;
    return trade;
  }}).filter(t => t.code && t.qty > 0);
}}

// ---- 組成跟真 API 一樣形狀的回應 ----

// 交割日是方法二的關鍵：cdate 比今天晚才代表這筆錢還沒離開銀行帳戶。
// 真的網站是 T+2；「已交割」整筆往前挪兩個交易日（前幾天成交、錢已經扣掉了）；
// 「T+1」往前挪一個交易日成交、明天才交割（還沒扣，但跟今天成交的 T+2 不同一天）。
function datesOf(t) {{
  if (t.settled === '1') return {{ tdate: PREV, cdate: STAMP }};
  if (t.settled === 'next1') return {{ tdate: PREV1, cdate: NEXT1 }};
  return {{ tdate: STAMP, cdate: SETTLE }};
}}

function detailOf(t, index) {{
  const d = datesOf(t);
  return {{
    tagName: 'stkdat', bhno: META.bhno, cseq: META.cseq,
    stkno: t.code, stkna: t.name, trade: '0', bs: t.bs,
    tdate: d.tdate, cdate: d.cdate,
    qty: String(t.qty), price: String(t.price), priceqty: String(t.priceqty),
    fee: String(t.fee), tax: String(t.tax), payn: String(t.payn),
    ordno: 'S' + String(index + 1).padStart(4, '0'),
  }};
}}

function build() {{
  const list = trades();

  const pnl = holdings().filter(h => h.qty > 0).map(h => {{
    const cost = Math.round(h.qty * h.price);
    const value = Math.round(h.qty * h.mkt);
    const detail = list.filter(t => t.code === h.code).map(detailOf);
    return {{
      tagName: 'stksum', bhno: META.bhno, cseq: META.cseq,
      stkno: h.code, stkna: h.name, trade: '0', stype: 'H',
      costqtyn: String(h.qty), priceavgn: String(h.price), costsumn: String(cost), costsum: String(cost),
      pricemkt: String(h.mkt), valuemktn: String(value), valuenown: String(value),
      makeasum: String(value - cost),
      makeaper: cost ? ((value - cost) / cost * 100).toFixed(2) : '0',
      stkdat: detail,
    }};
  }});

  const ok = {{ retcode: '000000', retmsg: '' }};
  // Amount 的單位是分（真的回應長這樣：0000000089300 = 893.00），
  // 假的也要照給，不然模擬跑得過、真帳號一上線就差 100 倍。
  // 負數要把負號留在最前面，不能連負號一起補零（補成 000-793672200 就沒有人讀得懂了）。
  // 真的網站會不會出現負的銀行餘額沒有人看過，這裡照「負號 + 補零」給，
  // 至少讀得回來；真的遇到再照實際格式改。
  const raw = Math.round(num(document.getElementById('bank')) * 100);
  const cents = (raw < 0 ? '-' : '') + String(Math.abs(raw)).padStart(13, '0');

  // 交割金額查詢（query610）：一天一列，金額是那一天要交割的淨額。真的網站不管
  // 有沒有成交，都會把最近幾個交易日各列一列（沒成交就是 0），所以這裡也是固定
  // 三列、金額 0 也照給 —— 「今天那一列在不在」是程式的一道檢查（見
  // planner._cash_blocked），不能因為今天沒成交就讓它消失。
  // 三列分別是：已交割（前兩天成交、今天交割，錢已經在銀行餘額裡了）、
  // T+1（前一天成交、明天交割，還沒交割）、未交割（今天成交、T+2 交割，還沒交割）——
  // 「銀行餘額推算」要把 T+1 跟未交割兩列都加進去，只看未交割那一列會少算。
  const sumOf = (settled) => list.filter(t => t.settled === settled)
      .reduce((s, t) => s + t.payn, 0);
  const due = [
    {{ trade: PREV, cdate: STAMP, pay_amt: String(sumOf('1')) }},
    {{ trade: PREV1, cdate: NEXT1, pay_amt: String(sumOf('next1')) }},
    {{ trade: STAMP, cdate: SETTLE, pay_amt: String(sumOf('0')) }},
  ];

  return {{
    '未實現損益': Object.assign({{ arrays: pnl }}, ok),
    '交割金額': Object.assign({{ data: due }}, ok),
    '銀行餘額': Object.assign({{ data: [{{
      qry_date: STAMP, qry_times: '090000',
      bnkno: '050', bnkacc: META.bnkacc,
      Amount: cents,
    }}] }}, ok),
  }};
}}

window.__SIM__ = build;

// ---- 畫面上的推導欄位 ----

function refresh() {{
  document.querySelectorAll('#pnl tbody tr').forEach(tr => {{
    const qty = num(tr.querySelector('.qty'));
    const cost = Math.round(qty * num(tr.querySelector('.price')));
    const value = Math.round(qty * num(tr.querySelector('.mkt')));
    tr.querySelector('.amount').textContent = money(cost);
    tr.querySelector('.value').textContent = money(value);
    tr.querySelector('.pnl').textContent = money(value - cost);
    tr.querySelector('.rate').textContent = cost ? ((value - cost) / cost * 100).toFixed(2) + '%' : '—';
    sign(tr.querySelector('.pnl'), value - cost);
    sign(tr.querySelector('.rate'), value - cost);
  }});

  let net = 0;
  document.querySelectorAll('#mat tbody tr').forEach(tr => {{
    ['priceqty', 'fee', 'tax', 'payn'].forEach(k => tr.querySelector('.' + k).textContent = '');
  }});
  trades().forEach(t => {{
    net += t.payn;
    t.row.querySelector('.priceqty').textContent = money(t.priceqty);
    t.row.querySelector('.fee').textContent = money(t.fee);
    t.row.querySelector('.tax').textContent = money(t.tax);
    t.row.querySelector('.payn').textContent = money(t.payn);
    sign(t.row.querySelector('.payn'), t.payn);
  }});
  const total = document.getElementById('net');
  total.textContent = money(net);
  sign(total, net);
}}

function reroll() {{
  document.querySelectorAll('#pnl tbody tr').forEach(tr => {{
    const base = parseFloat(tr.dataset.base);
    const price = tick(base * (0.85 + Math.random() * 0.3));
    const lots = Math.max(1, Math.round((500000 + Math.random() * 4500000) / (price * 1000)));
    tr.querySelector('.price').value = price;
    tr.querySelector('.qty').value = lots * 1000;
    tr.querySelector('.mkt').value = tick(price * (0.9 + Math.random() * 0.22));
  }});
  refresh();
}}

INIT_TRADES.forEach(tradeRow);
function addTrade(settled) {{
  const first = document.querySelector('#pnl tbody tr');
  tradeRow({{ code: first ? first.dataset.code : '2059', bs: 'B', qty: 1000,
             price: first ? num(first.querySelector('.mkt')) : 100, settled: settled }});
  refresh();
}}
document.getElementById('add').addEventListener('click', () => addTrade('0'));
document.getElementById('addT1').addEventListener('click', () => addTrade('next1'));
document.addEventListener('input', refresh);
document.addEventListener('change', refresh);
refresh();

// ---- 下單模擬：假帳號的委託／查詢／撤單，供 dev_tools/simulate_orders.py 呼叫 ----
//
// 只做「下單分頁／掛單分頁那層協調邏輯」（多輪收斂、等 6 秒撤零股、每輪重新
// 同步）測得到的最小行為——真的操作 select2／layer.js 那幾支操作真實網站 DOM
// 的程式碼（order_fill.py／order_query.py／order_cancel.py）完全不會被跑到，
// 假帳號一律繞過它們，那一層的正確性要靠真帳號實測，見對應模組的模組說明。
//
// 成不成交是固定機率決定的，不是查真實委買賣一比價——這幾檔雖然用的是真實
// 股票代號（見 STOCKS），盤中真的查得到報價，但拿它來判斷成交會讓測試結果
// 綁著「現在是不是開盤時間」，離峰時間永遠查不到、永遠走同一條路。IOC（整股．
// 盤中追價）下單當場知道成不成交；ROD（零股、買賣股票／全持股交易）掛著，
// 之後每次被查詢（掛單分頁查詢、出清零股等 6 秒撤單前的那次查詢）才「擲骰子」
// 決定要不要判定成交——ROD_FILL_CHANCE 刻意調低，貼近現實裡零股 6 秒內大多數
// 時候不會成交、得靠撤單清掉那個常態（docs/介面規劃.md 9.8）。
let ORDERS = [];
let orderSeq = 0;
const IOC_FILL_CHANCE = 0.7;
const ROD_FILL_CHANCE = 0.3;

function genOrdno() {{
  orderSeq += 1;
  return 'F' + String(orderSeq).padStart(4, '0');
}}

// 委託成交要連帶改兩個地方，跟人手動測試時「改 qty 輸入格 + 按新增一筆成交」
// 做的是同一件事：持股（#pnl 那一列的 .qty）真的減少／增加，成交明細
// （#mat）多一筆——後者才是「未實現損益」「交割金額查詢」算現金與股數的唯一
// 資料來源（見 build() 與檔案開頭的說明），只改 qty 不會反映到那兩支查詢上。
function applyFill(code, side, qty, price) {{
  const row = [...document.querySelectorAll('#pnl tbody tr')]
      .find((tr) => tr.dataset.code === code);
  if (row) {{
    const box = row.querySelector('.qty');
    const sign = side === 'B' ? 1 : -1;
    box.value = Math.max(0, num(box) + sign * qty);
  }}
  tradeRow({{ code, bs: side, qty, price, settled: '0' }});
  refresh();
}}

// orgqty 的單位跟真的網站一樣看盤別：整股（apcode '1'）填的是「張」，零股
// （apcode '5'）填的是「股」（見 order_fill.TAB1_LOT／TAB1_ODD 那段說明）。
// 但 #pnl 的 .qty／成交明細的股數一律是**股**（跟 Excel E 欄同一個單位），
// 整股成交要拿 orgqty 乘 1000 才是真正要從持股扣掉的股數，這裡漏了會少改
// 1000 倍，多輪出清會一直看到「還有一大堆股數沒賣掉」。
const SHARES_PER_LOT = 1000;

function sharesOf(order) {{
  return order.apcode === '1' ? Number(order.orgqty) * SHARES_PER_LOT : Number(order.orgqty);
}}

function placeOrder(opts) {{
  const ordno = genOrdno();
  const order = {{
    ordno, stockno: opts.code, buysell: opts.side, apcode: opts.apcode, trade: '0',
    priceflag: '0', odprice: String(opts.price),
    orgqty: String(opts.qty), matqty: '0', celqty: '0',
    ordstatus: '2', act: 'O', bs_flag: opts.bsFlag,
    orddate: STAMP, ordtime: '090000000', workdate: STAMP,
  }};
  ORDERS.push(order);

  if (opts.bsFlag === 'I') {{
    // 不是全有全無：真實 IOC 吃到多少對手單就成交多少張，剩下的交易所自動
    // 取消，不是「這一張委託全部成交或全部失敗」——大單一次全部成交在現實
    // 裡反而是少數。這裡也刻意不是全有全無，成交量落在委託量的 20%~100%
    // 之間隨機決定，這樣多輪測試才會真的需要跑好幾輪才清得完，跟按「多輪
    // 直到出清」想測的東西一致（一次就全部成交，多輪那段邏輯反而測不到）。
    const totalLots = Number(opts.qty);
    let filledLots = 0;
    if (Math.random() < IOC_FILL_CHANCE) {{
      filledLots = Math.max(1, Math.round(totalLots * (0.2 + Math.random() * 0.8)));
    }}
    order.matqty = String(filledLots);
    order.celqty = String(totalLots - filledLots);
    order.act = 'C';   // 沒吃到的部份，IOC 這一刻交易所就自動取消了

    if (filledLots === 0) {{
      return {{ ok: false, message: 'IOC. FOK 委託未能成交，委託失敗', ordno, matched: 0 }};
    }}
    applyFill(opts.code, opts.side, filledLots * SHARES_PER_LOT, opts.price);
    const note = filledLots < totalLots
        ? `（部分成交 ${{filledLots}}/${{totalLots}} 張，其餘 IOC 自動取消）` : '';
    // matched 是這一次呼叫「當場」確定成交的股數（IOC 沒有懸而未決這回事，
    // 送出去那一刻結果就確定了）——呼叫端（ui_order_exec._order_fill_job）
    // 拿它來判斷「這一輪是不是真的什麼都沒發生」，不是靠事後比對持股猜的，
    // 見那邊「沒有進展就停」那道保險怎麼用這個欄位。
    return {{ ok: true, message: '委託成功, 委託書編號: ' + ordno + note, ordno,
             matched: filledLots * SHARES_PER_LOT }};
  }}
  // ROD：先掛著，成不成交留到之後查詢那一刻才決定（見 resolvePending），
  // 下單當下 matched 一律是 0——不是「還不知道」，是「這一刻確定還沒有」。
  return {{ ok: true, message: '委託成功, 委託書編號: ' + ordno, ordno, matched: 0 }};
}}

function resolvePending() {{
  ORDERS.forEach((order) => {{
    if (order.matqty !== '0' || order.celqty !== '0') return;   // 已經有結果了，不重算
    if (Math.random() < ROD_FILL_CHANCE) {{
      order.matqty = order.orgqty;
      applyFill(order.stockno, order.buysell, sharesOf(order), Number(order.odprice));
    }}
  }});
}}

function snapshotOrder(order) {{
  // celable／errcode 不存進 ORDERS，是每次查詢當下依 matqty/celqty 現算的
  // 衍生欄位——省得兩處分別維護，改一個忘了改另一個。
  const open = order.matqty === '0' && order.celqty === '0';
  return Object.assign({{}}, order, {{ errcode: '00000000', celable: open ? '1' : '0' }});
}}

function queryOrders() {{
  resolvePending();
  return ORDERS.map(snapshotOrder);
}}

function cancelOrders(ordnos) {{
  const wanted = new Set(ordnos);
  const results = [];
  const locked = [];
  ORDERS.forEach((order) => {{
    if (!wanted.has(order.ordno)) return;
    if (order.matqty !== '0' || order.celqty !== '0') {{ locked.push(order.ordno); return; }}
    order.celqty = order.orgqty;
    order.act = 'C';
    results.push({{
      ordno: order.ordno, code: order.stockno,
      side: order.buysell === 'B' ? '買進' : '賣出', ok: true, message: '刪單成功',
    }});
  }});
  return {{ results, locked }};
}}

window.__SIM_ORDER__ = {{ place: placeOrder, query: queryOrders, cancel: cancelOrders }};
</script>
</body></html>"""
