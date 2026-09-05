"""
Excel 讀寫。只認得持股管理表的這幾格，其餘全是公式，一律不碰：

認得的是 **`持股管理-台美股-10家.xls` 那一版**（一頁 10 檔，2026/09/02 使用者
定案統一用它，見 STOCK_LIMIT）：

    D4:D13  股票名稱(代號)  只讀，用來對出每一列是哪一檔
    E4:E13  股數           <- 未實現損益的「成交股數」
    F4:F13  成本           <- 未實現損益的「成交均價」
    I4:I13  股價           只讀，跟 D 欄同一列對應同一檔股票；由使用者既有的
                           「更新股價」巨集（Module1.更新股價）填入，出清股票
                           多輪模式的「自動更新股價」用（見 orders.py 開頭說明）
    M19:N28 下單試算        只讀，加碼股數（正買負賣）與價格，由「自動計算」巨集
                           寫入；買賣股票作業的張數與價格就是這裡來的
    B8      現金餘額        <- 由紀錄檔的現金流水算出來
    A22/B22 今年報酬率      只讀。A22 是**整份表的錨點**（開檔就檢查，見
                           layout_problem），B22 是值——下單分頁的「執行帳戶」
                           照它由低到高排執行順序（見 orders.order_accounts）
    D1      程式維護提醒     每次寫入順便刷新，關閉程式時清掉，見 marker_enabled

位置寫死是刻意的：這支程式只認得這個特定格式的持股管理表，認錯了就是把數字
寫進別人的格子裡，寧可一開始就對不上而報錯——**「一開始就對不上而報錯」現在
真的做出來了**：開檔時檢查 A22 是不是「今年報酬率」，不是就擋下來（09/02 使用者
要求）。5 檔那一版的同一個標籤在 A17，而它的持股只到 D8、下單試算在 M14:N18，
每一個程式會讀寫的位置都不一樣。

版面完全維持原樣，沒有任何輔助欄位 —— 前日餘額、每日淨收付、最後更新日，
全部記在 ledger.py 管理的紀錄檔裡。
"""

import codecs
import contextlib
import os
import re
import threading
import time
from pathlib import Path

from login import app_dir
from util import env_int, to_num

ENV_FILE = ".env"
ENV_KEY = "EXCEL_PATH"

# 一張表最多幾檔持股，也就是巨集裡的 `Def_stock_limit`。**10**（2026/09/02
# 使用者定案：統一用 `持股管理-台美股-10家.xls` 這一版，不再支援 5 檔那一版）。
#
# 這份表的每一個位置都是從它算出來的，所以這裡只留這一個數字，HOLDING_ROWS、
# PLAN_ROWS、今年報酬率那一列全部由它推——以前是各寫死一次，改表的時候很容易
# 只改到一邊，而那種錯不會報錯，只會把 M 欄某一列的試算值配到別檔股票身上
# （見 PLAN_ROWS 的說明）。三個位置實測都對得上（2026/09/02 用那份表核過）：
# 持股 D4:D13 ＝ 4 ～ 3+limit、下單試算 M19:N28 ＝ limit+9 ～ limit*2+8、
# 今年報酬率 A22/B22 ＝ limit+12。
STOCK_LIMIT = 10

HOLDING_ROWS = range(4, 4 + STOCK_LIMIT)
COL_NAME, COL_QTY, COL_COST = 4, 5, 6
CELL_BALANCE = (8, 2)

# I 欄「股價」，跟 D 欄同一列對應同一檔股票——跟股數/成本是同一個列配置，
# 不是另外一套對應規則。只讀，程式從來不寫這一欄（見 read_sheet）。
COL_PRICE = 9

# 使用者既有的巨集（2026/08/28 使用者確認：Module1 的 Sub 更新股價()），
# 平常手動點的「更新」按鈕背後就是呼叫這個——多輪出清模式的「自動更新股價」
# 開關，就是每一輪開始前對「每一個要用到的分頁」各觸發一次這個巨集，等它
# 跑完再讀那一頁的 I4:I13（見 run_update_price_macro／orders.py 開頭的模式
# 說明）。2026/08/29 拿到巨集原始碼後確認它抓的是 Yahoo 的報價 API、而且
# 只對 ActiveSheet 動作，所以「一頁一次」是硬性要求，不是保守做法。
UPDATE_PRICE_MACRO = "Module1.更新股價"

# 下單試算：M 欄是股數（正買負賣）、N 欄是價格，一列對一檔股票，跟上面 D/E/F
# 那幾列是**同一批股票的另一組列**（10 檔版：D4 對 M19、D5 對 M20…）。列號在
# 巨集裡是算出來的（stock_limit + 9 ～ stock_limit * 2 + 8），這裡照同一條式子
# 從 STOCK_LIMIT 推，見 docs/介面規劃.md 9.1 那張表。
#
# 只讀，程式從來不寫這兩欄：要動它們是靠觸發使用者自己的「自動計算」巨集，
# 跟 I 欄那個「巨集寫、程式讀」的分工一樣。
PLAN_ROWS = range(STOCK_LIMIT + 9, STOCK_LIMIT * 2 + 9)
COL_PLAN_QTY, COL_PLAN_PRICE = 13, 14

# 使用者既有的另一支巨集（Module1）。跟「更新股價」一樣，只認 ActiveSheet
# （無限定的 Range()），一定要一個分頁 Activate 一次、跑一次；而且一樣會跳
# MsgBox——「檢查輸入錯誤」那段最多 23 個（見 docs/介面規劃.md 9.6 第 2 點
# 那張表），Application.Run 會卡在那裡等人按確定，背景執行緒就無聲掛住。
# 9.6 那套 Python 事前檢查刻意還沒做（真的卡住了再回頭照那裡寫的做），跟
# 「更新股價」的 #1 MsgBox 是同一種、一直以來就有的風險，不是這裡新加的。
AUTO_CALC_MACRO = "Module1.自動計算"

# 今年報酬率，下單功能排執行順序用（報酬率低的先執行）。只讀，不寫——
# 這一格本來就是 Excel 自己的公式算出來的，程式沒有理由覆蓋它。
#
# **這一格同時是整份表的錨點**（2026/09/02 使用者指定）：標籤在 A22、值在
# B22，對不上就是開錯版本的持股管理表，開檔那一刻就擋下來（見 layout_problem）。
# 09/02 早上曾經改成掃 A 欄找標籤，好讓 B17／B22／B32 三種格式都讀得到；同一天
# 下午使用者定案「統一用 10 家那一版」，掃描就退場了——會掃是因為當時要相容
# 多種格式，而相容多種格式本身正是危險的來源：其他兩種格式的下單試算根本不在
# M19:N28，讀得到報酬率只會讓一份對不上的表看起來像可以用。
RETURN_RATE_LABEL = "今年報酬率"
RETURN_RATE_ROW = STOCK_LIMIT + 12
COL_LABEL, COL_RETURN_RATE = 1, 2

# 分頁名稱就算錨點對得上，也不當交易人帳戶——目前只有這一個名字
# （2026/09/02 使用者要求：Excel 裡的「虛擬帳戶」分頁不要出現在下單的執行帳戶清單）。
EXCLUDED_ACCOUNT_SHEET_NAMES = {"虛擬帳戶"}

# D1 沒被上面三處佔用，理論上可以安全借來提醒「這份檔案有程式在管」。
# 語意是「這份檔案由程式維護」的持久標記，不是「程式現在正在跑」的即時燈號——
# _write_worker 每次同步只是短暫用 COM 開檔、寫入、存檔、關檔，平常大部分時間
# 程式根本沒碰著檔案，做不出真正的即時狀態，文字用詞也刻意避開「控制中」這種
# 容易被誤會成即時狀態的說法。
CELL_MARKER = (1, 4)
MARKER_TEXT = "此檔案由程式自動寫入現金餘額、股數與成本"
MARKER_COLOR = 0xFF0000  # COM 的 Font.Color 是 BGR，這個值是藍字
MARKER_ENV_KEY = "EXCEL_CONTROL_MARKER"


def marker_enabled():
    """.env 的 EXCEL_CONTROL_MARKER 沒設或設 1 就開著；設 0 就整個關掉這個功能。"""
    return env_int(MARKER_ENV_KEY, 1) != 0

# Excel 實例在半路死掉時會看到的 HRESULT。
#   0x800A01A8  物件不見了（Open 成功之後，下一句就撲空）
#   0x80010108  用戶端中斷了已啟動物件的連線
DEAD_OBJECT_CODES = (-2146827864, -2147417848)

DEAD_EXCEL_HINT = (
    "Excel 開起來之後又自己消失了。\n\n"
    "這通常是 Office 沒有啟用：啟用檢查失敗時 Excel 會把自己收掉、"
    "改用「產品啟動失敗」模式重開，程式手上的物件就跟著失效。\n\n"
    "兩個辦法：\n"
    "  1. 先自己用 Excel 把這個檔開著再跑 —— 程式會接上你那個視窗，不會另外開一個。\n"
    "  2. 把 Office 啟用起來，從根本解決。"
)


def is_dead_object(exc):
    """這個例外是不是「Excel 死在半路」。是的話錯誤訊息要多說一句，不然沒人看得懂。"""
    code = getattr(exc, "hresult", None)
    if code is None:
        args = getattr(exc, "args", ())
        code = args[0] if args and isinstance(args[0], int) else None
    return code in DEAD_OBJECT_CODES

# 從「台灣50(0050)」取出 0050。全形括號也吃，因為是人手打的欄位。
CODE_PATTERN = re.compile(r"[（(]\s*([0-9A-Za-z]+)\s*[)）]")


def excel_path():
    """
    要同步哪一份 Excel。只看 .env 的 EXCEL_PATH，沒設就回傳 None（還沒選過）。

    刻意不給預設檔名。給了預設的話，換一台機器或改個檔名，程式會安靜地去同步
    「另一個檔」或報「找不到某某檔」，而不是直接說「你還沒選」—— 對一支會改寫
    檔案的程式來說，猜錯對象是最不能接受的錯。

    相對路徑以 .env 所在資料夾為基準（打包成 exe 後就是 exe 旁邊）。
    """
    raw = os.getenv(ENV_KEY, "").strip().strip('"')
    if not raw:
        return None
    path = Path(os.path.expandvars(raw))
    return path if path.is_absolute() else app_dir() / path


def remember_excel_path(path):
    """
    把選好的檔案寫回 .env 的 EXCEL_PATH，下次開啟就直接用，不必再選一次。

    只動那一行：整個檔讀進來、找到就換掉、沒有就補在最後，其他行原封不動 ——
    .env 裡都是人寫的設定與註解，任何「重新產生一份」的做法都會把它們洗掉。
    BOM 也要留著：用記事本存過的 .env 檔頭會有 BOM，拿掉之後下次讀取
    第一個設定會失效（load_dotenv 是用 utf-8-sig 讀的）。
    """
    env = app_dir() / ENV_FILE
    raw = env.read_bytes() if env.is_file() else b""
    has_bom = raw.startswith(codecs.BOM_UTF8)
    text = raw.decode("utf-8-sig") if raw else ""
    newline = "\r\n" if "\r\n" in text else "\n"

    lines = text.splitlines()
    entry = f"{ENV_KEY}={path}"
    for i, line in enumerate(lines):
        stripped = line.strip()
        if stripped.startswith("#") or "=" not in stripped:
            continue
        if stripped.split("=", 1)[0].strip() == ENV_KEY:
            lines[i] = entry
            break
    else:
        if lines and lines[-1].strip():
            lines.append("")
        lines.append("# 要同步哪一份 Excel。介面上按「開啟EXCEL」選檔時會自動改這一行。")
        lines.append(entry)

    body = (newline.join(lines) + newline).encode("utf-8")
    env.write_bytes((codecs.BOM_UTF8 if has_bom else b"") + body)

    # 這個行程的 os.environ 也要跟著更新，否則同一次執行裡再呼叫 excel_path()
    # 拿到的還是舊的（load_dotenv 只在 import 時跑過一次）。
    os.environ[ENV_KEY] = str(path)


def stock_code_of(text):
    """「台灣50(0050)」-> 「0050」。沒有括號就回 None（空白列）。"""
    found = CODE_PATTERN.search(str(text or ""))
    return found.group(1).strip().upper() if found else None


def read_sheet(sheet, rows=None):
    """
    把一個分頁讀成純 Python 資料，之後的計算都不必再碰 COM。

    這樣做不只是為了乾淨：COM 每讀一格都是一次跨行程呼叫，而且介面版之後
    要在背景執行緒跑，讓計算完全脫離 COM 物件會少掉很多麻煩。

    `rows` 不給就是 HOLDING_ROWS（D4:D13，＝程式會寫回 E/F 的那幾列），
    也是所有呼叫端目前用的值。留這個參數是因為 09/02 早上下單那條路曾經掃得比
    寫得寬（那時候持股只認 5 列），統一成 10 檔版之後兩邊又是同一個範圍了。
    每一列回傳的 "row" 是真正的 Excel 列號，寫回去的時候照它定位。
    """
    scan = HOLDING_ROWS if rows is None else rows
    rows = []
    for row in scan:
        label = sheet.Cells(row, COL_NAME).Value
        code = stock_code_of(label)
        if not code:
            continue                 # 空白列，這張表允許留空
        rows.append({
            "row": row,
            "label": str(label).strip(),
            "code": code,
            "qty": sheet.Cells(row, COL_QTY).Value,
            "cost": sheet.Cells(row, COL_COST).Value,
            "price": to_num(sheet.Cells(row, COL_PRICE).Value, None),
        })

    return {
        "balance": to_num(sheet.Cells(*CELL_BALANCE).Value, None),
        "rows": rows,
    }


def anchor_ok(sheet):
    """
    這個分頁是不是「程式認得的那一版持股管理表」：**A22 要正好是「今年報酬率」**
    （2026/09/02 使用者指定的錨點檢查）。

    為什麼是這一格：整份表的兩個區塊都是照巨集的 `stock_limit` 算出來的，
    10 檔那一版把它排在 `stock_limit + 12` ＝ 第 22 列。換成 5 檔那一版，同一個
    標籤會落在 A17，而那一版的下單試算在 M14:N18、持股只有 D4:D8——**每一個
    程式會讀寫的位置都不一樣**。一格對得上，其餘位置就都對得上；對不上就整份
    表都不能用，不是只有報酬率讀不到。

    只讀一格（1 次 COM 往返），所以敢在開檔時對每一個分頁都問一次。
    """
    label = sheet.Cells(RETURN_RATE_ROW, COL_LABEL).Value
    return isinstance(label, str) and label.strip() == RETURN_RATE_LABEL


def layout_problem(book):
    """
    開檔時的錨點檢查（2026/09/02 使用者要求）：這份活頁簿有沒有任何一個看得見的
    分頁長得像程式認得的那一版。沒有就回一句話說明，有就回 None。

    **一個都對不上才算問題**，不是「有一個對不上就擋」：一份活頁簿裡可以有說明頁、
    工作用的空白頁，那些本來就不是交易人分頁（見 list_account_sheets）。全部都對
    不上才代表「開錯檔案／開到別的版本」——那時候不能讓程式繼續，它接下來要寫的
    E/F 欄、要讀的 M/N 欄全部會落在別人的格子上，而且不會報錯。
    """
    visible = [s for s in book.Worksheets if s.Visible == -1]
    if any(anchor_ok(sheet) for sheet in visible):
        return None
    names = "、".join(s.Name.strip() for s in visible[:5]) or "（沒有看得見的分頁）"
    return (f"這份活頁簿裡沒有一個分頁的 A{RETURN_RATE_ROW} 是「{RETURN_RATE_LABEL}」，"
            f"跟程式認得的持股管理表（{STOCK_LIMIT} 檔那一版：持股 D4:D"
            f"{3 + STOCK_LIMIT}、下單試算 M{PLAN_ROWS[0]}:N{PLAN_ROWS[-1]}、"
            f"今年報酬率 B{RETURN_RATE_ROW}）對不上。\n\n"
            f"看到的分頁：{names}\n\n"
            f"多半是開到別的版本（5 檔那一版的今年報酬率在 A17，"
            f"每一個位置都不一樣）。請改開 10 家那一版。")


def read_return_rate(sheet):
    """
    今年報酬率（B22），只給下單功能排執行順序用。

    位置寫死，而且**先驗錨點**（見 anchor_ok）：A22 不是「今年報酬率」就回 None，
    不去猜它搬到哪裡了。09/02 早上那版會掃 A 欄找標籤，好相容 B17／B22／B32 三種
    格式；同一天下午使用者定案統一用 10 家那一版之後，那種相容反而有害——別的
    格式讀得到報酬率，只會讓一份每個位置都對不上的表看起來像可以用。

    讀不到（錨點不對、空格、公式錯誤值那類）就回 None，不要回 0——0 在這裡
    看起來像一個真正的答案（今年打平），會被誤判成「報酬率最低，第一個執行」，
    跟 planner.bank_balance 讀不懂銀行餘額時的態度一樣。
    """
    if not anchor_ok(sheet):
        return None
    return to_num(sheet.Cells(RETURN_RATE_ROW, COL_RETURN_RATE).Value, None)


def list_account_sheets(book):
    """
    這份活頁簿裡有哪幾位交易人，以及各自的今年報酬率。
    回傳 [(分頁名, 報酬率或 None), ...]，順序照活頁簿裡分頁本來的順序。

    給下單分頁的「執行帳戶」用（2026/09/01）：那份清單原本要等登入拿到名字才
    長得出來，但一份持股管理表的分頁**本來就是一位交易人一頁**，開檔當下就
    知道有誰了，不必等登入。登入只跟「送得出委託」有關，跟「這份表裡有誰」無關。

    只讀 A22／B22（錨點與今年報酬率），不跑巨集、不碰 D/E/F/I/M——一位兩次 COM
    往返，20 位也只是幾十毫秒，這也是它敢在開檔／切分頁時自動跑的原因
    （見 ui_order.refresh_order_accounts）。

    **哪些分頁算數**：看得見的、而且**錨點對得上**（A22 是「今年報酬率」，見
    anchor_ok）。2026/09/02 之前是「報酬率讀得到 或 D 欄至少有一個股票代號」那種
    兩選一的特徵判斷；改用錨點之後，判斷的是「這一頁是不是程式認得的那一版持股
    管理表」——那才是真正該問的問題，而且順便把「一頁全空的新交易人」也收進來了
    （他一格股票都還沒填，舊的判斷會把他整個漏掉，人在畫面上找不到他）。

    不用分頁名稱去猜是誰（名字就是人名，猜不出規則），也不預設「所有分頁都是交易人」——
    唯一的例外是名字剛好等於 EXCLUDED_ACCOUNT_SHEET_NAMES 裡列的那幾個（目前只有
    「虛擬帳戶」）：那種分頁是特地留著當範本／測試用的，錨點會對得上但不該被當成
    可以下單的真帳戶，2026/09/02 使用者要求排除。
    """
    accounts = []
    for sheet in book.Worksheets:
        # -1 是 xlSheetVisible；0 隱藏、2 是「非常隱藏」（只能用 VBA 叫回來）。
        name = sheet.Name.strip()
        if sheet.Visible != -1 or name in EXCLUDED_ACCOUNT_SHEET_NAMES or not anchor_ok(sheet):
            continue
        accounts.append((name,
                         to_num(sheet.Cells(RETURN_RATE_ROW, COL_RETURN_RATE).Value, None)))
    return accounts


def first_visible_sheet(book):
    """
    活頁簿由左到右第一個「算數的交易人分頁」——條件跟 list_account_sheets
    完全一樣（看得見、名字不在 EXCLUDED_ACCOUNT_SHEET_NAMES、錨點對得上，見
    anchor_ok）。找不到（全部都藏起來、全部被排除、或全部錨點不對）回 None。

    2026/09/05 補上錨點檢查：一開始只排除了「虛擬帳戶」這個特定名字，但實測
    發現活頁簿裡還會有別的看得見、名字也不是範本、卻不是持股管理表的分頁
    （例如某交易人自己另外開的「OOO資產」總覽頁，A22 不是「今年報酬率」）。
    這種分頁一樣會被 list_account_sheets 排除在「執行帳戶」之外，但改之前
    這裡只看 Visible 跟名字，還是會把它當成「第一個分頁」，讀出它的
    D4~D13（形狀對不上，通常是空的）當股票候選——「讀取ＯＯ持股」因此讀到
    一個聽起來不像交易人、股票代號也是空的分頁。用錨點取代單純的名字排除，
    才是真正跟 list_account_sheets 同一個判斷：「這個分頁算不算一位交易人」，
    不是「這個分頁的名字認不認識」——名字沒辦法窮舉，錨點才是結構上的答案。

    獨立成一支是因為下單分頁「讀取ＯＯ持股」那顆按鈕要在畫面上講出這個分頁
    的名字（見 ui_order.refresh_order_stock_list），不能只跟 read_stock_list
    要一份沒有名字的股票清單。用 `for ... in book.Worksheets` 取，不用
    `book.Worksheets(1)`——理由同 read_stock_list 那段說明。
    """
    return next((s for s in book.Worksheets
                 if s.Visible == -1 and s.Name.strip() not in EXCLUDED_ACCOUNT_SHEET_NAMES
                 and anchor_ok(s)),
                None)


def read_stock_list(book):
    """
    **第一個分頁**的 D4~D13 有哪幾檔股票，回傳 [(代號, D 欄原文), ...]。
    下單分頁「指定股票」那個下拉的候選就是這一份（2026/09/02 使用者指定）。

    為什麼是第一個分頁而不是各帳戶自己那一頁：下單分頁改成「勾好幾個帳戶、
    一次跑完」之後，要選的是**這一輪要動哪幾檔股票**，不是「某一位手上有
    什麼」——每位手上的檔數本來就不一樣，拿其中一位的清單當候選會讓別人
    多出來的那幾檔選不到。哪一位沒有選中的那一檔，執行預覽那一列會寫
    「這一位沒有這檔」並略過（見 orders.plan_trade_orders）。

    「第一個分頁」定義見 first_visible_sheet。

    只讀不寫，掃到什麼算什麼——這個分頁如果不是持股管理表（將來有人在最
    前面插一頁說明），D4~D13 就一個股號都對不出來，回空清單，畫面上的下拉
    是空的，不會拿一份猜出來的清單讓人選。
    """
    sheet = first_visible_sheet(book)
    if sheet is None:
        return []
    stocks = []
    for row in HOLDING_ROWS:
        label = sheet.Cells(row, COL_NAME).Value
        code = stock_code_of(label)
        if code:
            stocks.append((code, str(label).strip()))
    return stocks


def read_order_plan(sheet):
    """
    讀這一頁的下單試算：M19:N28，配上 D4:D13 的股票代號。
    回傳 {股票代號: {"name", "qty", "price"}}，只讀不寫。

    照 read_return_rate 的樣子獨立一支，不塞進 read_sheet——那支是更新分頁在
    用的，每次讀取 20 個帳戶都會叫，沒有理由為了下單分頁的功能讓它多讀 10 格。

    配對是**照位置**的（D4↔M14、D5↔M15…），不是照股票代號比對：巨集本身就是
    這樣算列號的（見 PLAN_ROWS）。所以 D 欄空著的那一列，就算 M 欄有殘值也不會
    被撿進來——沒有股號的試算值沒有意義（那正是「自動計算」會跳「有股數無股號！」
    的情況，見 9.6）。

    qty 讀不到當 0（沒有試算＝這一檔這一輪不動），price 讀不到留 None，不要當
    0——0 在價格這裡看起來像一個真的答案，跟 read_return_rate 的態度一樣。
    """
    plan = {}
    for row, plan_row in zip(HOLDING_ROWS, PLAN_ROWS):
        code = stock_code_of(sheet.Cells(row, COL_NAME).Value)
        if not code:
            continue
        plan[code] = {
            "name": str(sheet.Cells(row, COL_NAME).Value or "").strip(),
            "qty": to_num(sheet.Cells(plan_row, COL_PLAN_QTY).Value, 0) or 0,
            "price": to_num(sheet.Cells(plan_row, COL_PLAN_PRICE).Value, None),
        }
    return plan


# 看門狗預設等幾秒才喊「可能卡住了」。巨集內部是逐檔同步 HTTP（見
# run_update_price_macro 的說明），5 檔正常也就 1~2 秒上下；抓寬一點是因為
# 網路本來就會抖，抓太緊只會在正常變慢的時候也喊，喊多了等於沒喊。
MACRO_STUCK_AFTER = 10


def _run_macro_watched(excel, macro, on_stuck, stuck_after):
    """
    真的呼叫 `excel.Run(macro)`，`on_stuck` 給的話另外起一顆
    `threading.Timer` 在旁邊倒數：超過 `stuck_after` 秒 `Run` 還沒回來就
    呼叫一次 `on_stuck`（多半是巨集跳了 9.6 那幾個 MsgBox，`Application.Run`
    卡在那裡等人按確定）。

    這顆計時器**擋不住、也不會去按掉**那個對話框——`DisplayAlerts` 管不到
    VBA 自己叫的 MsgBox（見 docs/介面規劃.md 9.6），Python 這邊沒有辦法主動
    打斷一個正在跑的 `Application.Run`。它純粹是「去 Excel 看一眼」的提醒：
    `Run` 正常跑完就在 `finally` 把計時器取消掉，`on_stuck` 不會被叫到。
    """
    if on_stuck is None:
        excel.Run(macro)
        return
    timer = threading.Timer(stuck_after, on_stuck)
    timer.daemon = True
    timer.start()
    try:
        excel.Run(macro)
    finally:
        timer.cancel()


def run_auto_calc_macro(excel, sheet, on_stuck=None, stuck_after=MACRO_STUCK_AFTER):
    """
    對 sheet 這一個分頁觸發「自動計算」巨集：依 M4:M13 目標比重反覆試算，
    結果寫進 M19:N28。

    跟 run_update_price_macro 同一條規矩：**一定要傳分頁進來、一頁跑一次**，
    只認 ActiveSheet；呼叫端一樣要自己包 keep_active_sheet()。

    過程中會**暫時改寫 E4:E13／F4:F13／B8 再還原**（docs/介面規劃.md 9.6 第 3
    點），跟更新分頁「E/F/B8 一律覆蓋」正面衝突，呼叫端要確保跟那邊互斥
    （`_excel_in_use()`）。也可能跳 MsgBox（見 AUTO_CALC_MACRO 上面的說明），
    卡住的話畫面會停在「跑巨集中」不動，要去 Excel 視窗把對話框按掉——
    `on_stuck` 是這件事的提醒（見 `_run_macro_watched`），不是防呆。

    **2026/09/02 起沒有任何一條活路呼叫這支**：下單分頁那顆「執行 更新→自動計算」
    拿掉了（使用者要求，Excel 上自己按就好，見 docs/介面規劃.md 9.2）。留著是因為
    程式要跑這支的場合是**多輪執行時自己跑**（勾了才跑）與全持股交易的「跑自動計算」，
    兩條都還沒接——手動按一顆按鈕那條路才是被否決的那一種。
    """
    sheet.Activate()
    _run_macro_watched(excel, AUTO_CALC_MACRO, on_stuck, stuck_after)


@contextlib.contextmanager
def keep_active_sheet(book):
    """
    把使用者原本停在的那一頁記下來，離開這個區塊時還回去。

    跑「更新股價」巨集一定要一頁一頁 Activate（見 run_update_price_macro），
    但 Excel 通常就開在使用者眼前——20 個帳戶跑完把他丟在最後一個分頁上，
    等於每按一次「讀取試算」畫面就被搬走一次。記一次、還一次，比每
    跑一頁就來回切兩次少掉一半的 COM 往返與畫面重繪。

    記不住或還不回去都不是錯誤（分頁被刪了、活頁簿被關了、Excel 正忙），
    整段吞掉就好：這只是「別動到使用者的畫面」這種禮貌，不該讓真正要做的
    讀取因為它失敗。
    """
    before = None
    try:
        before = book.Application.ActiveSheet.Name
    except Exception:
        pass
    try:
        yield
    finally:
        if before is not None:
            try:
                sheet, _ = find_sheet(book, before)
                if sheet is not None:
                    sheet.Activate()
            except Exception:
                pass


def run_update_price_macro(excel, sheet, on_stuck=None, stuck_after=MACRO_STUCK_AFTER):
    """
    對 sheet 這一個分頁觸發「更新股價」巨集（見 UPDATE_PRICE_MACRO 的說明）。

    **一定要傳分頁進來、一頁跑一次。** 巨集寫在標準模組裡，用的是無限定的
    `Range("D" & i)`／`Range("I" & i)`，VBA 把這種寫法解析成 ActiveSheet，
    所以它只更新「當下作用中的那一頁」。2026/08/29 之前這裡整批只呼叫一次，
    20 個帳戶裡只有使用者上次剛好停在的那一頁拿到新價格，其餘讀回來的是
    上一次的舊 I4:I13——不會報錯、不會少一欄、看起來完全正常，而盤中追價的
    基準價就建立在這上面。呼叫端要自己包 keep_active_sheet()，跑完把畫面
    還給使用者。

    excel 是 open_workbook() 回傳的第一個值（Application 物件，不是
    Workbook）——巨集掛在整個活頁簿上，用 Application.Run 呼叫；sheet 是
    find_sheet() 回傳的那個 Worksheet。

    巨集內部是 MSXML2.XMLHTTP 對 Yahoo 報價 API 打**同步** GET
    （`.Open "GET", url, False` 第三個參數 False 就是同步），所以
    Application.Run 回來的時候 I4:I13 確實已經是這一次抓到的值，不必再等、
    也不必懷疑讀到舊數字。代價是「一檔股票一次 HTTP 往返」：5 檔 × N 個
    分頁的請求會逐一發出去，分頁多的時候這一步本來就慢，那是慢不是卡住。

    **抓不到價格時巨集會跳 `MsgBox "xxxx 股價更新失敗"`**，那是 VBA 的
    modal 對話框，Application.Run 會停在那裡等人按確定——背景執行緒就這樣
    無聲掛住，畫面停在「更新股價、讀取中…」。Excel 的 DisplayAlerts 管不到
    VBA 自己叫的 MsgBox，Python 這邊擋不掉，只能改巨集本身。`on_stuck` 給的
    話會在卡住 `stuck_after` 秒後喊一聲（見 `_run_macro_watched`），把無聲
    掛住換成一句看得懂的提示，不是真的解決掉這個坑。
    """
    sheet.Activate()
    _run_macro_watched(excel, UPDATE_PRICE_MACRO, on_stuck, stuck_after)


def write_cells(sheet, writes):
    """
    writes 是 (row, col, value) 的清單。只會寫進記憶體，存檔是另一回事。

    現金餘額那格 value 會是「=1000-107」這種公式字串（見 util.cash_formula），
    要走 Formula 屬性才會被 Excel 當成公式算，走 Value 屬性只會存成一串文字。
    """
    for row, col, value in writes:
        cell = sheet.Cells(row, col)
        if isinstance(value, str) and value.startswith("="):
            cell.Formula = value
        else:
            cell.Value = value


def write_marker(sheet):
    """在 D1 補上藍字提醒。每次同步寫入都順手刷新，反正已經在開檔，不多一次 COM 往返。"""
    cell = sheet.Cells(*CELL_MARKER)
    cell.Value = MARKER_TEXT
    cell.Font.Color = MARKER_COLOR


def clear_all_markers(path):
    """
    程式正常關閉時呼叫：把 Excel 裡所有分頁的 D1 提醒清掉。

    範圍是整份活頁簿，不只這次執行寫過的分頁——帳戶數不多，全部掃一遍成本低，
    還能順便清掉上次非正常結束（當機、被工作管理員砍掉）遺留下來的提醒。

    Best-effort：任何一步失敗都默默放棄，不能卡住關閉流程。真正的風險是相反的
    不對稱——非正常結束不會走到這裡，D1 會留著過期的提醒，那個風險是可以接受的，
    但絕不能因為清除失敗就讓程式關不掉。

    活頁簿的鎖只等 3 秒就放棄（opened 的 wait），不跟其他呼叫端一樣無限等：這裡
    跑在關閉流程的主執行緒上，背景真的卡住的話（例如巨集跳了 MsgBox 停在那裡，
    見 docs/介面規劃.md 9.6）無限等就等於關不掉，而 D1 留著過期提醒是可以接受的。
    """
    if not marker_enabled() or path is None or not path.is_file():
        return
    try:
        with opened(path, True, wait=3) as (_excel, workbook, _attached):
            changed = False
            for sheet in workbook.Worksheets:
                cell = sheet.Cells(*CELL_MARKER)
                if str(cell.Value or "").strip() == MARKER_TEXT:
                    cell.Value = None
                    changed = True
            if changed:
                workbook.Save()
    except Exception:
        return


def find_sheet(book, name):
    """用分頁名稱找分頁。同名不只一個就不猜，直接報錯。"""
    wanted = (name or "").strip()
    if not wanted:
        return None, "沒有分頁名稱可以對應（登入後沒讀到帳戶名）"
    matches = [s for s in book.Worksheets if s.Name.strip() == wanted]
    if not matches:
        return None, f"找不到名稱為「{wanted}」的分頁"
    if len(matches) > 1:
        return None, f"有 {len(matches)} 個分頁都叫「{wanted}」，無法判斷要寫哪一個"
    return matches[0], None


def is_open_in_excel(path):
    """檔案是不是正被 Excel 開著。"""
    if (path.parent / ("~$" + path.name)).exists():
        return True
    try:
        with open(path, "r+b"):
            return False
    except OSError:
        return True


# 同一份活頁簿一次只讓一條執行緒操作。
#
# 為什麼需要：程式接上的是使用者眼前那個 Excel 實例（見 _open_once 的 GetObject
# 分支），不是各開各的一份。而「更新分頁寫入」「下單分頁的讀取ＯＯ持股／讀取
# 試算／新增股票／多輪之間重讀」各有各的忙碌旗標、各跑各的執行緒，彼此不知道
# 對方存在。
#
# 程式自己的讀寫都是 sheet.Cells(...) 這種限定寫法，不受別人 Activate 影響；
# 但巨集用的是無限定的 Range()，只認 ActiveSheet（見 run_update_price_macro）。
# 兩條執行緒交錯 Activate 的話，巨集會跑在別人剛切過去的那一頁上——那一頁被更新
# 兩次、自己這一頁從來沒更新過，而接著讀回來的是舊的 I4:I13。不報錯、不缺欄位，
# 只是靜靜地拿舊價格當盤中追價的基準，跟 2026/08/29 修掉的那個 ActiveSheet
# bug 是同一種失敗，只是成因換成並行。
#
# 畫面那一層（ui_background._excel_in_use）已經把按鈕擋掉了，這把鎖是替「以後
# 新加一個 COM 入口、忘了問那個述詞」兜底。兩層都要。
_EXCEL_LOCK = threading.Lock()


class ExcelBusy(RuntimeError):
    """另一條執行緒正在用這份活頁簿，而呼叫端不願意等（見 opened() 的 wait）。"""


@contextlib.contextmanager
def opened(path, write, wait=None):
    """
    open_workbook ＋ close_workbook 的成對版本，順便鎖住「一次只有一條執行緒
    在動這份活頁簿」：

        with excel_io.opened(path, True) as (excel, book, attached):
            ...

    wait=None 是一直等——背景工作都該用這個，它們本來就在背景，等一下沒關係。
    給數字就是最多等幾秒、等不到丟 ExcelBusy，只有「不做也沒關係、但絕對不能
    卡住」的呼叫端才用（見 clear_all_markers）。

    新的 COM 入口一律走這裡，不要再自己配對 open_workbook／close_workbook——
    漏掉鎖不會報錯，只會在很久以後變成一個算錯的數字。
    """
    if not _EXCEL_LOCK.acquire(timeout=-1 if wait is None else wait):
        raise ExcelBusy(f"另一項作業正在使用這份 Excel：{path}")
    try:
        excel, book, attached = open_workbook(path, write)
    except Exception:
        _EXCEL_LOCK.release()
        raise
    try:
        yield excel, book, attached
    finally:
        try:
            close_workbook(excel, book, attached)
        finally:
            _EXCEL_LOCK.release()


def open_workbook(path, write, attempts=3):
    """
    取得 Workbook，回傳 (excel, book, attached)。

    檔案已經開在 Excel 裡時就直接接上那個實例，不必為了跑這支程式先關檔。
    接上的好處不只是方便：讀到的是你畫面上的即時內容（含還沒存檔的修改）。
    寫入之後一樣會存檔 —— 留給人自己按 Ctrl+S 聽起來像多一道確認，實際上是
    「紀錄檔記成寫過了、檔案卻沒存」這個破口的來源。

    attached=True 時絕對不能 Close 或 Quit —— 那是使用者的視窗，關掉他會很錯愕。

    為什麼開完要再碰一下、失敗還要重開
    ----------------------------------
    Office 沒啟用時，Excel 起來之後那道啟用檢查失敗會把自己收掉重來，
    Open 明明成功，下一句碰 Worksheets 就是 OLE error 0x800A01A8。
    錯誤發生在半路最麻煩，所以開完先碰一下確認物件是活的，
    死的就整個丟掉、換一個新的 Excel 再試。
    """
    trouble = ""
    for attempt in range(attempts):
        excel = book = None
        attached = False
        try:
            excel, book, attached = _open_once(path, write)
        except Exception as exc:
            trouble = str(exc)
        else:
            if _still_alive(book):
                return excel, book, attached
            trouble = DEAD_EXCEL_HINT
            close_workbook(excel, book, attached)
        if attempt < attempts - 1:
            time.sleep(1.0)

    raise RuntimeError(f"開不了 Excel：{path}\n\n{trouble}")


def _still_alive(book):
    """碰一下 COM 物件，確認它還在。死掉的物件在這裡就會丟例外。"""
    try:
        book.Worksheets.Count
        return True
    except Exception:
        return False


def _open_once(path, write):
    import win32com.client

    if is_open_in_excel(path):
        try:
            book = win32com.client.GetObject(str(path))
            if Path(str(book.FullName)).resolve() == path.resolve():
                return book.Application, book, True
        except Exception:
            pass
        raise RuntimeError(
            f"這個檔案正被 Excel 開著，但接不上那個 Excel：{path}\n"
            f"請先在 Excel 存檔並關閉，再執行一次。"
        )

    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    return excel, excel.Workbooks.Open(str(path), ReadOnly=not write), False


def close_workbook(excel, book, attached):
    """
    接上使用者自己開的 Excel 時，關檔與結束都不是我們該做的事。

    這裡的錯誤一律吞掉。收尾失敗沒有任何補救動作可做，而 Excel 已經死掉時
    Close/Quit 只會再丟一個例外，把真正的原因蓋在「During handling of the
    above exception」底下 —— 那正是最需要看清楚的那一行。
    """
    if attached:
        return
    for action in (lambda: book.Close(False), excel.Quit):
        try:
            action()
        except Exception:
            pass
