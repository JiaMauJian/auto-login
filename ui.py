"""
持股同步的桌面介面。

    python ui.py

兩個分頁分別對應兩件事：

    同步      這次要改哪幾格、哪幾格程式不准碰
    歷程      誰在什麼時候改了哪一格（程式、人工、交接）

同步分頁是左右兩半：左邊一份交易人名單（誰要處理、現金剩多少），右邊只畫
選中的那一位。20 個帳號攤平成一張表是兩百多列，捲到一半連標頭都看不到是誰；
而實際的工作方式是依序輪詢、一次只處理一個人，所以「換下一位」被做成一個
動作（按鈕或 Ctrl+↑ / Ctrl+↓），右邊那張表永遠一頁看得完。

右邊那張表排得跟 Excel 的「持股資料」一樣：一檔股票一列，股數與成本並排，
現金（B8）像在 Excel 裡一樣擺在表格外面、自己一條（負的數字自己紅字）。畫面跟檔案同形狀，
對照的時候才不必在心裡翻譯一次。值平常只寫現在的數字，有變化才寫成
「舊 → 新」—— 20 檔裡通常只有一兩格要動，那一兩個箭頭才跳得出來。

一次讀全部，還是只更新一位
--------------------------
上面那個「範圍」跟左邊那份名單是同一個選擇的兩個入口：名單上點誰，範圍就換成誰，
「讀取網頁資料」那顆也跟著改名成「更新（王小明）」，按下去只查他一個、只寫他那一頁。
要重讀全部就把範圍切回「全部」。

一天下來按最多次的是「只更新一位」（盯著某一位的部位在動），一次讀全部反而只有
開盤前那一次。預設仍然是全部，因為交易人的名字要登入之後才從網站拿得到
（.env 裡只有帳密），沒讀過一輪之前左邊名單是空的，也就沒有人可以點。

只更新一位有兩件事要守住：

一、畫面上的資料新舊不一。別人那幾列是上一輪讀的，所以每一位都記著自己的讀取時間，
右邊標頭寫「讀取於 10:32」—— 沒有它，半小時前的數字跟剛讀的長得一模一樣。

二、寫入、落帳、接管只能碰這一輪讀到的那幾位（round_scope）。名單上別人也可能有
「要寫」的格子，那是用上一輪的網頁資料算出來的，順手寫出去就是拿舊資料改 Excel。

還有一個藏在瀏覽器裡的限制：整個瀏覽器只有一組 cookie，同時只帶得動一個人的身分
—— 而被頂掉的那個分頁自己不會知道，探起來還像活著。但伺服器那邊 20 個帳號的
session 是各自獨立的，B 登入不會殺掉 A 的，所以換人不必重登：把 A 登入時收下來的
cookie 換回去就回到他登入完的那一刻（見 fetch.new_store），只有他在伺服器上逾時了
才需要真的重登一次。

原本還有第三個「現金帳本」分頁（基準、逐日流水、補登、重新校正）。
登入即初始化之後那些操作全部沒有存在的必要了：基準每次登入自己重設，
要修正就直接改 Excel 的 B8 —— 下次登入程式就以那個數字為準。
現金的流水還是照記，只是那是程式內部的帳，不再是一個要人看、要人操作的畫面。

「改 B8 就好」只在隔天成立。同一天第二次以後登入，基準已經設過、不會再跟著 B8 走，
手改只會被判成人工改動然後凍住一整天。所以基準本身就寫在畫面上：現金那一條底下
一行「今日初始現金餘額」，旁邊一顆「修改」（見 _fill_opening、edit_opening）。

    現金餘額 = 今日初始現金餘額 + 今日淨收付

右邊那項是網頁抄來的、不會錯，所以餘額不對的時候要改的一定是左邊那項。

這裡原本是另一種做法：程式自己判斷「今天可能已經開過了」，讀完網頁資料就跳一個
對話框、一次問完所有分頁。改掉的理由是猜錯的兩個方向都很貴 —— 不該問的時候問，
20 個分頁按到最後沒人在看內容；該問的時候沒問（或當下填錯），當天就再也沒有入口。
數字一直在畫面上、按鈕一直按得動，就不必由程式決定什麼時候該問誰。

Excel 由誰維護，只由上面那個「程式自動更新」開關決定
--------------------------------------------------
勾起來：「讀取網頁資料」讀完就自己備份、寫進 Excel，一顆按鈕從頭做到尾。
取消勾選：程式一格都不碰，同步分頁只是對照表，數字由人自己在 Excel 改。

原本每一格前面都有勾選框、外加全選／全不選、還有一顆「交還給程式」，
等於同一件事有三層開關。20 個帳號、上百格的規模下沒有人按得完，
真正的決定只有「這次讓不讓程式寫」，所以只留這一個開關。
右下角原本還留著一顆「寫入」當人工維護時的逃生口，後來也拿掉了：
真的想讓程式寫，把開關勾起來再讀一次就是了，不必有第二條路。

登入完成的當下，Excel 上的股數、成本、現金餘額會直接被收成程式的起點
（見 _initialize）。那個時間點今天要買賣什麼都還沒發生，Excel 就是唯一真相，
所以不必問任何問題 —— 連現金也不必問「含不含今天的淨收付」，登入時它一定
還沒含，今天成交了什麼等「讀取網頁資料」時再往上加。

開關記在紀錄檔裡，下次打開沿用。一個決定「程式會不會動你的檔案」的開關，
不該每次啟動就自己跳回自動。

畫面只做顯示與操作，所有判斷都來自 planner.py —— 介面與命令列走同一段程式碼，
才不會出現「介面接管的結果跟命令列不一樣」這種最難查的問題。

執行緒
------
登入、抓資料、開 Excel 都很慢，全部丟到背景執行緒，否則視窗會整個凍住。
Tk 的元件只能在主執行緒碰，所以背景做完只把純資料丟進 queue，
由主執行緒定時取出來畫 —— 不在背景執行緒動任何 widget。

背景執行緒要用 COM（Excel）之前一定要先 CoInitialize，這是 Windows 的規定，
少了它 win32com 會直接丟例外。
"""

import datetime
import json
import os
import queue
import threading
import traceback
import tkinter as tk
from pathlib import Path
from tkinter import filedialog
from tkinter import font as tkfont
from tkinter import messagebox, ttk

from playwright.sync_api import Error as PlaywrightError
from playwright.sync_api import sync_playwright

import excel_io
import ledger as ledger_mod
import planner
import profile_tools
import fetch as fetch_mod
from fetch import collect, login_only
from login import app_dir, configure_browsers_path, load_accounts, open_context
from util import cell_name, env_int, show, to_num, values_match

# 「補登」這個來源已經沒有入口了（現金帳本分頁拿掉時一起收掉），
# 但舊的歷程檔裡還有這種紀錄，名字要留著才不會顯示成一個英文代號。
SOURCE_NAMES = {"program": "程式", "human": "人工", "adopt": "交接", "backfill": "補登"}

# 背景做的三件事，講給人聽的名字。收尾出錯時要說得出是哪一步壞掉的。
STEP_NAMES = {"logged_in": "登入", "fetched": "讀取網頁資料", "written": "寫入"}

# 瀏覽器起不來時，錯誤視窗最上面那段人話。traceback 講的是 Playwright 的內部狀況，
# 對使用者沒有意義，真正能動手的只有下面這兩件事。
BROWSER_HINT = (
    "瀏覽器沒能開起來，所以這次的動作沒有做。常見的原因有兩個：\n"
    "・這台電腦上找不到指定的瀏覽器（.env 的 BROWSER_CHANNEL 指到沒裝的版本），"
    "或 Playwright 的 Chromium 還沒下載完\n"
    "・USER_DATA_DIR 那個資料夾正被另一個 Chrome 視窗開著 —— 同一個資料夾不能同時"
    "給兩個 Chrome 用，請把 Chrome 全部關掉再按一次"
)

# 三個歷程篩選選單的「不篩選」那一項。用一段文字當值而不是 None，
# 選單裡本來就會有一列寫著它，兩邊用同一個東西才不會對不起來。
ALL_CHOICE = "全部"

# 期間只給這兩個級距。歷程是拿來回答「剛才那次執行做了什麼」「這幾天有沒有人動過」，
# 再細的區間（挑日期、選月份）沒有人真的會用，卻要多一個日曆元件跟一整套錯誤處理。
WHEN_TODAY, WHEN_WEEK = "今天", "近 7 天"

# 外觀設定，全部從 .env 讀，改完重開介面生效。

FONT_ENV_KEY = "UI_FONT_SIZE"
FONT_DEFAULT = 12
FONT_MIN, FONT_MAX = 8, 24

WIDTH_ENV_KEY, HEIGHT_ENV_KEY = "UI_WIDTH", "UI_HEIGHT"
# 視窗再小就不叫小視窗，而是壞掉的版面：左邊名單擠成一條、右邊那幾格全是省略號。
WINDOW_MIN_W, WINDOW_MIN_H = 640, 400


def font_size_from_env():
    """
    一般文字的字級。上下限是版面撐得住的範圍：再小按鈕上的字會糊成一團，
    再大則欄寬跟著長大到視窗塞不下（wide() 放得大欄寬，放不大螢幕）。
    超出範圍就夾到邊界，不當成錯誤。
    """
    return max(FONT_MIN, min(FONT_MAX, env_int(FONT_ENV_KEY, FONT_DEFAULT)))


def window_size_from_env():
    """
    視窗的長寬（像素）。沒設就回 None，表示照字級與交易人數自己算（見 _build）——
    自己算出來的尺寸會跟著字級與名單長度變，是比任何固定數字都好的預設值，
    所以這裡只在使用者明確填了數字時才蓋過它。

    填了就當作實際像素、不再乘上字級：使用者量的是螢幕上的視窗，不是某個
    基準字級下的寬度，把他填的 1400 放大成 1680 只會讓人以為程式沒讀到設定。
    上限不在這裡管 —— 桌面可用區多大要等視窗建出來才問得到，一併在 _build 夾。
    """
    width, height = env_int(WIDTH_ENV_KEY, None), env_int(HEIGHT_ENV_KEY, None)
    return (max(WINDOW_MIN_W, width) if width else None,
            max(WINDOW_MIN_H, height) if height else None)


FONT_SIZE = font_size_from_env()   # 一般文字：按鈕、表格、對話框
HINT_SIZE = FONT_SIZE - 1          # 次要文字：路徑、提醒、欄位說明
WINDOW_W, WINDOW_H = window_size_from_env()


def wide(pixels):
    """
    把「照 10 級字調出來的像素寬」換算成目前字級要的寬度。

    欄寬跟列高不會自己跟著字級長大。放大字卻沒放大欄寬的結果是
    「成本（2059 川湖）」被截成「成本（2059…」—— 使用者只會看到一個莫名其妙
    的省略號，不會知道那是版面問題。
    """
    return int(pixels * FONT_SIZE / 10)


def build_columns(tree, spec):
    """
    照 spec 把一張表的欄位設好，並讓寬度跟著表格伸縮。

    spec 是 (欄名, 標題, 比重, 下限, 對齊) 的序列，比重與下限都照 10 級字的像素給。
    """
    for key, title, _weight, floor, anchor in spec:
        tree.heading(key, text=title)
        tree.column(key, width=wide(floor), minwidth=wide(floor), anchor=anchor, stretch=False)
    tree.bind("<Configure>", lambda _event: fit_columns(tree, spec))


def fit_columns(tree, spec):
    """
    把欄位寬度按比重攤進表格當下的寬度，每一欄不低於自己的下限。

    為什麼不交給 ttk 的 stretch：它只有在「一開始就放得下」的時候才會分配。
    欄寬總和一旦超過表格實際拿到的寬度，ttk 會整組凍住 —— 之後把視窗拉多寬、
    把分隔線拖多開都不再重算。兩張表原本都踩在這條線上（明細表的固定寬加起來
    wide(837)，預設字級下就超出 34px；歷程表超出 85px），於是最右邊的欄位被切在
    畫面外，而 Treeview 沒有橫向捲軸，捲也捲不出來。自己算就沒有那個狀態，
    任何寬度都是當場分配。

    下限是「還讀得出來」的下限，不是好看的下限：擠到極限時寧可看到
    「175,000 → 18…」，也不要一整欄消失 —— 明細表被切掉的偏偏是「狀態」跟
    「說明」，這一格程式會不會覆蓋、為什麼不覆蓋，答案就在那兩欄裡。
    """
    room = tree.winfo_width()
    if room <= 1:
        return                            # 還沒排版完，等下一次 Configure

    floors = [wide(floor) for _key, _title, _weight, floor, _anchor in spec]
    weights = [weight for _key, _title, weight, _floor, _anchor in spec]
    spare = max(room - sum(floors), 0)
    widths = [floor + spare * weight // sum(weights)
              for floor, weight in zip(floors, weights)]
    if spare:
        # 整數除法會掉幾個像素，全部補給最後一欄（兩張表都是「說明」），才會剛好填滿。
        widths[-1] += room - sum(widths)

    for (key, _title, _weight, _floor, _anchor), width in zip(spec, widths):
        # 值沒變就別寫回去 —— 改欄寬會再引來一次 Configure，兩邊互相觸發會抖。
        if tree.column(key, "width") != width:
            tree.column(key, width=width)


def work_area(root):
    """
    桌面上真正能用的那一塊：主螢幕扣掉工作列。

    winfo_screenwidth/height 給的是含工作列的整個螢幕，照它算出來的中心點會比
    眼睛看到的中心低半個工作列 —— 視窗最下面那行狀態列（寫入結果就報告在那裡）
    常常正好壓在工作列上緣。

    問不到就退回整個螢幕：非 Windows 沒有這個 API，而 DPI 縮放沒對齊時 Windows
    給的是實體像素、Tk 給的是縮放後的，兩套數字混用會把視窗丟到螢幕外，所以
    看起來不合理的答案寧可整組不要。
    """
    screen_w, screen_h = root.winfo_screenwidth(), root.winfo_screenheight()
    try:
        import ctypes
        from ctypes import wintypes

        rect = wintypes.RECT()
        if ctypes.windll.user32.SystemParametersInfoW(0x0030, 0, ctypes.byref(rect), 0):
            width, height = rect.right - rect.left, rect.bottom - rect.top
            if 0 < width <= screen_w and 0 < height <= screen_h:
                return rect.left, rect.top, rect.right, rect.bottom
    except Exception:
        pass
    return 0, 0, screen_w, screen_h


def _browser_alive(context):
    """瀏覽器是不是還開著（使用者有沒有自己把它關掉）。"""
    try:
        return any(not pg.is_closed() for pg in context.pages)
    except PlaywrightError:
        return False


def _read_excel_after_fetch(records, path):
    """背景執行緒用：登入抓完網頁資料後，順便把 Excel 現值讀出來。只回傳純資料，不回傳任何 COM 物件。"""
    import pythoncom

    excel = workbook = None
    pythoncom.CoInitialize()
    try:
        excel, workbook, attached = excel_io.open_workbook(path, False)
        try:
            sheets, sheet_errors = {}, {}
            for record in records:
                name = record.get("sheet_name")
                if not name or record["problems"]:
                    continue
                sheet, error = excel_io.find_sheet(workbook, name)
                if sheet is None:
                    sheet_errors[name] = error
                else:
                    sheets[name] = excel_io.read_sheet(sheet)
        finally:
            excel_io.close_workbook(excel, workbook, attached)
    finally:
        # COM 參考要在 CoUninitialize 之前放掉。反過來的話，物件會在
        # 「已經沒有 COM 的執行緒」上被回收，那是未定義行為。
        sheet = excel = workbook = None
        pythoncom.CoUninitialize()

    return {"records": records, "sheets": sheets, "sheet_errors": sheet_errors, "attached": attached}


def _error_text(payload):
    """錯誤視窗的內容。看得懂的說明放最上面，原始 traceback 留在下面備查。"""
    detail = payload["error"][-1500:]
    hint = payload.get("hint")
    return f"{hint}\n\n────────────────\n{detail}" if hint else detail


def pick_font():
    """挑一個有中文字的字型。挑不到就交給 Tk 自己決定。"""
    families = set(tkfont.families())
    for name in ("Microsoft JhengHei UI", "Microsoft JhengHei", "Segoe UI"):
        if name in families:
            return name
    return "TkDefaultFont"


def center_on(win, parent):
    """
    把對話框擺在主視窗中間偏上。

    不設位置的話由視窗管理員決定，常常跑到螢幕角落或蓋住主視窗的邊，
    使用者得先找它在哪。偏上是因為對話框通常比主視窗矮，正中間會顯得偏低。
    """
    win.update_idletasks()
    width, height = win.winfo_width(), win.winfo_height()
    x = parent.winfo_rootx() + (parent.winfo_width() - width) // 2
    y = parent.winfo_rooty() + (parent.winfo_height() - height) // 3
    win.geometry(f"+{max(x, 0)}+{max(y, 0)}")


# 明細表的六欄：(欄名, 標題, 比重, 下限, 對齊)。比重與下限都照 10 級字的像素給，
# 實際寬度由 wide() 乘上字級，再由 _fit_columns 按比重攤進表格當下的寬度。
DETAIL_COLUMNS = (
    ("cells", "格子", 72, 52, "w"),
    ("stock", "股票", 140, 90, "w"),
    ("qty", "股數", 175, 110, "e"),
    ("cost", "成本", 150, 95, "e"),
    ("status", "狀態", 150, 90, "center"),
    ("note", "說明", 150, 60, "w"),
)

# 歷程表的七欄，格式同上。
HISTORY_COLUMNS = (
    ("at", "時間", 140, 100, "w"),
    ("sheet", "交易人", 90, 60, "w"),
    ("cell", "格子", 60, 40, "w"),
    ("label", "項目", 200, 120, "w"),
    ("change", "變化", 230, 130, "w"),
    ("by", "來源", 60, 40, "center"),
    ("note", "說明", 420, 80, "w"),
)

def ask_opening_balance(parent, family, name, current, item):
    """
    改一個人的「今日初始現金餘額」。回傳新的開盤前金額，取消或留空就回 None。

    這裡曾經有另一個對話框：程式自己判斷「今天可能已經開過了」，讀完網頁資料就
    跳出來、一次問完所有分頁。它被這顆按鈕取代了 —— 同一個問題，程式問是「猜
    你需要」，而猜錯的兩個方向都很貴：不該跳的時候跳，20 個分頁按到最後沒人在看
    內容；該跳的時候沒跳，或跳的時候填錯，當天就再也沒有入口。基準現在一直寫在
    畫面上，看到不對再按，不必由程式決定什麼時候該問誰。

    填的是「開盤前」而不是「現在正確的餘額」：開盤前的現金是今天唯一不會再變的
    量，往上加多少由程式自己算（見 planner.apply_cash_reset）。
    """
    win = tk.Toplevel(parent)
    win.title("修改今日初始現金餘額")
    win.transient(parent)
    win.resizable(False, False)

    answer = {}

    outer = ttk.Frame(win, padding=12)
    outer.pack(fill="both", expand=True)

    ttk.Label(outer, justify="left", text=(
        f"「{name}」今天開盤前有多少錢 —— 今天第一筆成交之前的現金。\n"
        f"餘額是它一路加上每天的淨收付算出來的，所以餘額不對的時候，要改的是這個數字。"
    )).grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 10))

    lines = (("現在記著的今日初始現金餘額", show(current)),
             (f"今日淨收付（{item['net_rows']} 筆成交）", show(item["net"])),
             ("Excel 上的現金餘額 " + item["cell"], show(item["current"])))
    for index, (title, value) in enumerate(lines, start=1):
        ttk.Label(outer, text=title, style="Hint.TLabel").grid(row=index, column=0,
                                                              sticky="w", pady=2)
        ttk.Label(outer, text=value).grid(row=index, column=1, sticky="e",
                                          padx=(24, 0), pady=2)

    text = tk.StringVar(value="" if current is None else show(current))
    row = len(lines) + 1

    # 會走到這個對話框的路幾乎只有一條：人自己在盤中改過 B8。他改進去的是「現在
    # 正確的餘額」，那個數字已經含了今天的成交 —— 照抄進來就等於把今天的淨收付
    # 算進起點裡，待會兒再加一次。扣掉淨收付才是開盤前，算出來擺著並且給一顆
    # 「帶入」，比請人自己心算可靠：這一格填錯，畫面上不會有任何徵兆。
    if item["current"] is not None:
        suggested = round(item["current"] - item["net"], 2)
        ttk.Label(outer, text=f"{item['cell']} 已經含了今天的成交的話，開盤前是",
                  style="Hint.TLabel").grid(row=row, column=0, sticky="w", pady=2)
        pick = ttk.Frame(outer)
        pick.grid(row=row, column=1, sticky="e", padx=(24, 0), pady=2)
        ttk.Label(pick, text=show(suggested)).pack(side="left")
        ttk.Button(pick, text="帶入", width=5,
                   command=lambda: (text.set(show(suggested)), entry.focus_set(),
                                    entry.select_range(0, "end"))).pack(side="left", padx=(8, 0))
        row += 1

    ttk.Label(outer, text="改成（今天開盤前的現金）").grid(row=row, column=0,
                                                          sticky="w", pady=(10, 2))
    entry = ttk.Entry(outer, width=16, font=(family, FONT_SIZE), justify="right",
                      textvariable=text)
    entry.grid(row=row, column=1, sticky="e", padx=(24, 0), pady=(10, 2))

    # 結果邊打邊算。要核對的是「按下去會變成什麼」，自己看得到就不必先在心裡
    # 算一次再賭它跟程式算的一樣。
    ttk.Label(outer, text=f"{item['cell']} 會變成", style="Hint.TLabel").grid(
        row=row + 1, column=0, sticky="w", pady=2)
    result = ttk.Label(outer, text="維持原樣", style="Hint.TLabel")
    result.grid(row=row + 1, column=1, sticky="e", padx=(24, 0), pady=2)

    def update(*_args):
        raw = text.get().strip().replace(",", "")
        if not raw:
            result.configure(text="維持原樣", style="Hint.TLabel")
            return
        opening = to_num(raw, None)
        if opening is None:
            result.configure(text="看不懂", style="Manual.TLabel")
            return
        result.configure(text=show(round(opening + item["net"], 2)), style="Auto.TLabel")

    text.trace_add("write", update)
    update()

    def confirm(*_args):
        raw = text.get().strip().replace(",", "")
        if not raw:
            win.destroy()
            return
        value = to_num(raw, None)
        if value is None:
            messagebox.showerror("看不懂這個數字",
                                 f"「{text.get().strip()}」不是一個數字。", parent=win)
            entry.focus_set()
            return
        answer["opening"] = value
        win.destroy()

    # 填錯只有一種形狀，就直接寫出來。這裡沒有任何檢查擋得住它：
    # 兩個數字都是合法的金額，錯的那個要到明天的餘額才看得出來。
    ttk.Label(outer, justify="left", style="Hint.TLabel", text=(
        "盤中自己改過 " + item["cell"] + " 的話要小心：改進去的通常是「已經含了今天成交」的餘額，\n"
        "不能直接當成開盤前 —— 那樣今天的淨收付會被算兩次。"
    )).grid(row=row + 2, column=0, columnspan=2, sticky="w", pady=(10, 0))

    buttons = ttk.Frame(outer)
    buttons.grid(row=row + 3, column=0, columnspan=2, sticky="e", pady=(12, 0))
    ttk.Button(buttons, text="取消", command=win.destroy).pack(side="left", padx=(0, 8))
    ttk.Button(buttons, text="確定", command=confirm).pack(side="left")

    win.protocol("WM_DELETE_WINDOW", win.destroy)
    win.bind("<Escape>", lambda _e: win.destroy())
    win.bind("<Return>", confirm)
    center_on(win, parent)
    win.grab_set()
    # 整段選起來：來改的人心裡已經有一個數字，直接打就換掉，不必先清空。
    entry.focus_set()
    entry.select_range(0, "end")
    parent.wait_window(win)
    return answer.get("opening")


class SyncApp:
    def __init__(self, root):
        self.root = root
        self.path = excel_io.excel_path()
        self.accounts = load_accounts()
        self.today = datetime.date.today()

        # 模擬帳號的分頁名。畫面上要標出來 —— 一次看 20 個分頁時，
        # 分不出哪個是真的就很容易把模擬的結果當成真帳務。
        self.fake_sheets = {a["name"] for a in self.accounts if a.get("fake")}

        self.records = {}        # 分頁名 -> 網頁資料
        self.sheet_data = {}     # 分頁名 -> Excel 現值
        self.before = {}         # (分頁名, 格子) -> 這批網頁資料讀進來時 Excel 上的舊值
        self.proposals = {}      # 分頁名 -> 提案清單
        self.warnings = {}       # 分頁名 -> 提醒
        self.problems = []       # 這一輪畫在提醒框裡的失敗原因（由 problem_of 攤平而來）
        self.current_sheet = None  # 右邊正在看哪一位交易人

        # 「第幾組帳號」與「哪一位交易人」的對照。帳號設定裡只有帳密沒有名字，
        # 名字要登入之後才從網站的 sessionStorage 拿得到，所以這份對照是一邊做
        # 一邊長出來的。模擬帳號例外 —— 它的名字本來就寫在 .env 裡，
        # 一開機就填得進去，逐一交易人更新在模擬模式下不必先讀一輪。
        self.trader_of = {i: a["name"] for i, a in enumerate(self.accounts, start=1)
                          if a.get("fake")}
        # 每一位的網頁資料是什麼時候讀的。改成一次只更新一位之後，畫面上同時
        # 存在好幾個時間點的資料 —— 不寫出來，別人那幾列看起來就跟剛讀的一樣新。
        self.read_at = {}        # 分頁名 -> datetime
        # 失敗原因改成用「第幾組」當 key，不再是一整串重來一次的清單：只更新一位的
        # 時候，別人上一輪的失敗還沒被解決，不能因為這一輪沒讀到他就當作沒事了。
        self.problem_of = {}     # 第幾組 -> 失敗原因
        # 這一輪動到哪幾位。寫入、落帳、接管都只能在這個範圍裡做 —— 別人手上那份
        # 是上一輪的舊資料，拿舊資料去寫 Excel 是這個改動最大的風險。
        self.round_scope = set()
        # 這一輪按下去的時候要做誰（None = 全部）。報告用的，不是判斷用的。
        self.round_target = None
        self.busy = False
        # 「修改今日初始現金餘額」現在能不能按。_fill_opening 判定，_sync_buttons 套用
        # ——「忙不忙」跟「有沒有資料」是兩件事，分開記才不會互相蓋掉。
        self.opening_ready = False
        self.write_count = 0     # 這次要寫幾格，寫完報告時要用
        self.queue = queue.Queue()
        self.browser_cmd_queue = queue.Queue()
        self.browser_thread = None
        # 已經丟給背景、還沒收到回話的指令有幾個。畫面解除「登入中…」只靠那則回話，
        # 所以「執行緒不在了」等於「這個數字永遠減不回去」——_check_browser_thread
        # 就是在盯這件事。
        self.browser_waiting = 0

        # 「憑證」分頁：交易人 -> {"text": 網頁抄來的原文, "expiry": 解析出來的到期日}。
        # 只在登入的當下抓一次（見 fetch._fetch_cert_status），所以還沒登入過的人不會出現。
        self.cert_status = {}
        # 這次工作階段已經提醒過的人 —— 快到期不必每讀一次資料就再跳一次視窗。
        self.cert_alerted = set()
        self._migrate_candidates = {}   # 遷移憑證那張表的列 id -> profile_tools.scan_cert_sources() 的一筆
        self.profile_busy = False       # 「建立 Profile」進行中，避免重複點
        self.cert_tab_scanned = False   # 「憑證」分頁只在第一次切過去時自動掃描一次，之後靠手動「掃描」

        self.ledger = None
        self.ledger_error = None
        # 這份 Excel 的紀錄檔是不是這次才生出來的。是的話，今天的現金起點是憑
        # 「B8 現在的數字」定的，而程式沒有任何辦法看出它含不含今天的成交
        # —— 唯一能發現的人是使用者，所以要在提醒欄講一句（見 _fill_notes）。
        # 記在這裡而不是每次去問 ledger.existed：存過一次檔它就變成 True 了，
        # 而「這一天是從一份空帳本開始的」這件事不會因為存過檔就不成立。
        self.ledger_fresh = False
        # 這份 Excel 現在有沒有真的開在 Excel 裡。登入與讀取都卡在這個旗標上，
        # 由 _poll_excel 每幾秒重新確認一次 —— 使用者中途把 Excel 關掉也算數。
        self.excel_open = False
        if self.path is not None:
            try:
                self.ledger = ledger_mod.Ledger(self.path)
                self.ledger_fresh = not self.ledger.existed
            except RuntimeError as exc:
                self.ledger_error = str(exc)

        auto = True if self.ledger is None else bool(self.ledger.setting("auto_write", True))
        self.auto_write = tk.BooleanVar(value=auto)
        # 只看有差異的：20 位裡通常只有幾位要動，這個開關把名單縮到那幾位。
        # 刻意不記進紀錄檔 —— 它是「這一輪想少看幾個人」，不是一個設定。
        self.only_diff = tk.BooleanVar(value=False)

        self._build()
        self._drain()

        self._poll_excel()

        if self.path is None:
            self._say("還沒選檔案。按左上角「開啟EXCEL」挑一份持股管理表，選過就會記住。")
        elif not self.path.is_file():
            self._say(f"找不到 Excel：{self.path}")
        elif self.ledger_error:
            messagebox.showerror("紀錄檔有問題", self.ledger_error)
        elif not self.accounts:
            self._say("找不到帳號設定，請先在 .env 填入 TBB_ID_1 / TBB_PASSWORD_1")
        elif not self.excel_open:
            # 開機這一次只在狀態列講一句，不彈視窗。每次開程式都彈一個，看兩天就
            # 變成閉著眼睛按掉的東西，真的有事要說的時候也會被一起按掉。
            self._say("這份 Excel 還沒開著 —— 按左上角「開啟EXCEL」把它打開，登入才會亮起來。")
        else:
            self._say("按「登入」開瀏覽器並自動登入，之後要更新資料時再按「讀取網頁資料」。")

        # 開機也要畫一次。右半邊的「還沒有資料 —— 按上面的『讀取網頁資料』」
        # 是在 _fill_head／_fill_notes 裡準備好的，而那兩個只有 fill_sync_tree 會叫到
        # ——不在這裡叫一次的話，那句話要等到第一次讀取（或換檔、切開關）才出現，
        # 也就是永遠等不到它真正該出現的那一刻：程式剛打開、人正在想「然後呢」。
        self.fill_sync_tree()
        self.refresh_history()

    # ---------- 版面 ----------

    def _build(self):
        family = pick_font()
        self.root.title("持股同步" + (f"（模擬模式：另有 {len(self.fake_sheets)} 個假帳號）"
                                     if self.fake_sheets else ""))
        # 字放大了版面就要跟著大，但不能大到超出桌面可用區 —— 視窗比可用區還高
        # 的話，最下面那行狀態列會被工作列蓋住，而那是報告寫入結果的地方。
        # .env 填了 UI_WIDTH/UI_HEIGHT 就用他填的，可用區這一關照樣要過。
        left, top_edge, right, bottom = work_area(self.root)
        room_w, room_h = right - left, bottom - top_edge
        width = min(WINDOW_W or wide(1180), room_w - 80)
        # 高度不跟著帳號數長：名單本來就有捲軸，滑鼠滾兩下的成本遠低於「20 個帳號
        # 就開一個佔滿整個桌面的視窗」。要一次看完整份名單就自己把視窗拉高，
        # 或在 .env 設 UI_HEIGHT。
        height = min(WINDOW_H or wide(720), room_h - 60)

        # 先擺一個大概的中心點，等元件都建好、標題列量得到了再擺準一次
        # （見 _build 最後的 _center）。不指定位置的話由視窗管理員決定，
        # 常常開在角落或跨到第二個螢幕上，每次開都得先把它拖回來。
        x = left + max((room_w - width) // 2, 0)
        y = top_edge + max((room_h - height) // 2, 0)
        self.root.geometry(f"{width}x{height}+{x}+{y}")
        self.root.minsize(min(wide(900), width), min(wide(560), height))

        style = ttk.Style()
        style.configure("Treeview", font=(family, FONT_SIZE), rowheight=wide(28))
        style.configure("Treeview.Heading", font=(family, FONT_SIZE, "bold"))
        style.configure("TButton", font=(family, FONT_SIZE))
        style.configure("TLabel", font=(family, FONT_SIZE))
        # 勾選框不吃 TLabel 的設定，要自己來。少了這一行，UI_FONT_SIZE 調大之後
        # 整個畫面只有「程式自動更新」「只看有差異的」兩個還是原本的小字 ——
        # 而前者正是決定程式會不會動你 Excel 的那個開關。
        style.configure("TCheckbutton", font=(family, FONT_SIZE))
        style.configure("TNotebook.Tab", font=(family, FONT_SIZE))
        style.configure("Hint.TLabel", font=(family, HINT_SIZE), foreground="#666666")
        style.configure("Big.TButton", font=(family, FONT_SIZE, "bold"))
        style.configure("Auto.TLabel", font=(family, HINT_SIZE), foreground="#1a7f37")
        style.configure("Manual.TLabel", font=(family, HINT_SIZE), foreground="#a34a00")
        self.family = family

        top = ttk.Frame(self.root, padding=(12, 10, 12, 6))
        top.pack(fill="x")

        # 三顆按鈕疊成左邊那一直行，由上而下就是做事的順序：
        # 開啟EXCEL -> 登入 -> 讀取網頁資料。
        #
        # 原本「讀取網頁資料」擺在最右邊，動線就變成左上、左下、再橫跨整個視窗
        # ——而它是這裡最常按的一顆（登入一天一次，讀取一天很多次）。垂直對齊
        # 還多帶一個好處：後兩顆在前一步沒完成時是灰的，排成一行才看得出那是
        # 「還沒輪到」而不是「壞了」。sticky="ew" 讓三顆一樣寬，看起來才是一疊
        # 步驟而不是三顆大小不一的按鈕。
        self.excel_button = ttk.Button(top, text="開啟EXCEL", command=self.open_excel)
        self.excel_button.grid(row=0, column=0, sticky="ew")

        self.path_label = ttk.Label(top, text=self._path_text(), style="Hint.TLabel")
        self.path_label.grid(row=0, column=1, columnspan=3, sticky="w", padx=(16, 0))

        self.login_button = ttk.Button(top, text="登入", command=self.start_login)
        self.login_button.grid(row=1, column=0, sticky="ew", pady=(6, 0))

        # 「範圍」不只是登入哪幾組 —— 讀取也照它走，而且它跟左邊那份名單是同一個
        # 選擇的兩個入口：名單上點一位，這裡就換成那一位；這裡換一位，名單也跟著跳。
        # 一次只更新一位是常態（一整天下來按最多次的就是它），一次讀全部反而是
        # 開盤前那一次，所以入口做成「預設全部、點了誰就只做誰」。
        ttk.Label(top, text="範圍").grid(row=1, column=1, sticky="w", padx=(16, 0), pady=(6, 0))
        # 名字不能叫 width —— 上面那個 width 是視窗寬度，_center 最後還要用它。
        choice_width = 22 if self.fake_sheets else 16
        self.account_choice = ttk.Combobox(top, values=self._account_choices(), state="readonly",
                                           width=choice_width, font=(family, FONT_SIZE))
        self.account_choice.current(0)
        self.account_choice.grid(row=1, column=2, sticky="w", padx=(8, 0), pady=(6, 0))
        self.account_choice.bind("<<ComboboxSelected>>", self._on_scope_changed)

        ttk.Label(top, text="左邊名單點一位，這裡就跟著換；要重讀全部就切回「全部」",
                  style="Hint.TLabel").grid(row=1, column=3, sticky="w",
                                            padx=(12, 0), pady=(6, 0))

        # 按鈕上的字跟著範圍走：全部是「讀取網頁資料」，選了一位就是「更新（王小明）」
        # —— 按下去會動到誰，寫在按鈕上，不必回頭去看那個下拉選單（見 _refresh_fetch_button）。
        self.fetch_button = ttk.Button(top, text="讀取網頁資料", style="Big.TButton", command=self.start_fetch)
        self.fetch_button.grid(row=2, column=0, sticky="ew", pady=(6, 0))

        # 開關就貼在「讀取網頁資料」旁邊：它改變的正是按下那顆按鈕之後會發生的事，
        # 擺在一起才看得出因果，放到別的分頁或選單裡就沒人記得自己是哪一邊。
        self.auto_check = ttk.Checkbutton(top, text="程式自動更新", variable=self.auto_write,
                                          command=self._on_auto_changed)
        self.auto_check.grid(row=2, column=1, columnspan=2, sticky="w", padx=(16, 0), pady=(6, 0))

        # 四個字的開關看不出後果，底下再寫一句「現在按下去會發生什麼事」。
        self.mode_hint = ttk.Label(top, text="")
        self.mode_hint.grid(row=3, column=1, columnspan=3, sticky="w", padx=(16, 0), pady=(2, 0))

        # 右邊留白那一欄負責吃掉多餘寬度，左邊那一直行才不會被拉開。
        top.columnconfigure(3, weight=1)

        self.progress = ttk.Progressbar(self.root, mode="indeterminate")

        self.tabs = ttk.Notebook(self.root)
        self.tabs.pack(fill="both", expand=True, padx=12, pady=(6, 0))
        self.tabs.bind("<<NotebookTabChanged>>", self._on_tab_changed)

        self._build_sync_tab()
        self._build_history_tab()
        self._build_cert_tab()

        self.status = ttk.Label(self.root, text="", style="Hint.TLabel", anchor="w", padding=(12, 6))
        self.status.pack(fill="x")

        self._refresh_mode_hint()
        self._refresh_fetch_button()
        # 按鈕的初始亮暗也要照規則來。少了這一次，Excel 沒開著時「登入」會亮到
        # 第一次狀態改變為止 —— 而「一直沒開」正好就是不會有改變的那種情況。
        self._sync_buttons()

        self._center(width, height)

    def _center(self, width, height):
        """
        把視窗擺到桌面可用區的正中央。

        為什麼不是在上面設完 geometry 就算數：geometry 的 height 不含標題列，
        但標題列確實佔著螢幕，少扣那一段視窗就會比正中央低半個標題列。標題列
        多高只有視窗真的建出來之後才量得到（winfo_rooty 是內容區、winfo_y 是
        含框的位置，差值就是它），所以這一步留到最後、元件都建好才做。
        """
        self.root.update_idletasks()
        left, top_edge, right, bottom = work_area(self.root)
        bar = max(self.root.winfo_rooty() - self.root.winfo_y(), 0)
        x = left + max((right - left - width) // 2, 0)
        y = top_edge + max((bottom - top_edge - height - bar) // 2, 0)
        self.root.geometry(f"{width}x{height}+{x}+{y}")

    def _build_sync_tab(self):
        """
        左邊挑人、右邊看那個人的那幾格。

        20 個交易人、每人十幾格，全部攤平成一張表是兩百多列 —— 捲到一半連標頭
        都看不到是誰。而實際的工作方式是「一次只處理一個人，處理完換下一個」，
        所以版面就照那個方式切：左邊一份名單負責「該找誰」，右邊只畫選中的那位，
        永遠一頁看得完，不必捲、不必展開群組。
        """
        frame = ttk.Frame(self.tabs, padding=8)
        self.tabs.add(frame, text="  同步  ")

        # 用 PanedWindow 是因為交易人的名字長短差很多，讓使用者自己拖比我猜寬度準。
        split = ttk.PanedWindow(frame, orient="horizontal")
        split.grid(row=0, column=0, sticky="nsew")

        self._build_people(split)
        self._build_detail(split)

        frame.rowconfigure(0, weight=1)
        frame.columnconfigure(0, weight=1)

    def _build_people(self, split):
        """左欄：所有交易人一次看完，誰要處理一眼就知道。"""
        box = ttk.Frame(split, padding=(0, 0, 8, 0))
        split.add(box, weight=0)

        head = ttk.Frame(box)
        head.grid(row=0, column=0, columnspan=2, sticky="ew")
        ttk.Label(head, text="交易人").pack(side="left")
        self.people_count = ttk.Label(head, text="", style="Hint.TLabel")
        self.people_count.pack(side="right")

        # 20 個人裡真正有差異的通常只有幾個，勾起來名單就只剩要動的那幾位。
        ttk.Checkbutton(box, text="只看有差異的", variable=self.only_diff,
                        command=self.fill_sync_tree).grid(row=1, column=0, columnspan=2,
                                                          sticky="w", pady=(2, 4))

        self.people = ttk.Treeview(box, columns=("cash", "flag"), show="tree headings",
                                   selectmode="browse")
        self.people.heading("#0", text="姓名")
        self.people.column("#0", width=wide(120), minwidth=wide(80), stretch=True)
        # 現金餘額直接列在名單上 —— 下單前最想知道的就是這個人還有多少錢，不必點進去。
        self.people.heading("cash", text="現金餘額")
        self.people.column("cash", width=wide(100), minwidth=wide(80), anchor="e", stretch=False)
        self.people.heading("flag", text="狀態")
        self.people.column("flag", width=wide(72), minwidth=wide(56), anchor="center", stretch=False)

        bar = ttk.Scrollbar(box, orient="vertical", command=self.people.yview)
        self.people.configure(yscrollcommand=bar.set)
        self.people.grid(row=2, column=0, sticky="nsew")
        bar.grid(row=2, column=1, sticky="ns")

        ttk.Label(box, text="點一下換人，或用 ↑ ↓；Ctrl+↑ / Ctrl+↓ 在哪都能換",
                  style="Hint.TLabel").grid(row=3, column=0, columnspan=2,
                                            sticky="w", pady=(4, 0))

        # 底色管「要不要處理」、前景色管「錢是不是負的」。兩件事用不同屬性，
        # 才能同時成立而不會互相蓋掉。
        self.people.tag_configure("attention", background="#eaf4ea")
        self.people.tag_configure("warned", foreground="#a34a00")
        self.people.tag_configure("negative", foreground="#c00000")
        self.people.bind("<<TreeviewSelect>>", self._on_person_selected)

        box.rowconfigure(2, weight=1)
        box.columnconfigure(0, weight=1)

    def _build_detail(self, split):
        """右欄：選中那位的每一格，外加提醒與按鈕。"""
        box = ttk.Frame(split)
        split.add(box, weight=1)

        # 標頭把「是誰、第幾位、現金多少、要寫幾格」寫在同一行，而且固定在上面。
        # 原本這些寫在群組列上，一捲就不見了。
        self.detail_head = ttk.Label(box, text="", font=(self.family, FONT_SIZE, "bold"))
        self.detail_head.grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 6))

        # 版面照 Excel 那張「持股資料」表：一檔股票一列，股數與成本並排。
        #
        # 原本是一格一列（E4 一列、F4 另一列），5 檔股票就攤成 10 列，眼睛得自己
        # 把上下兩列配對回同一檔 —— 但在 Excel 上它們本來就是同一列的兩欄，
        # 對照的時候還要在心裡做一次轉換。計算層仍然是一格一格算（planner 那邊
        # 自動／手動是分開判定的），併成一列純粹是顯示層的事。
        #
        # 值平常只寫一個數字，有話要說的時候才寫成兩個（「舊 → 新」，或不會寫的
        # 「現值（網頁 …）」）。20 檔裡通常只有一兩格要動，這樣掃一眼就找得到。
        # 說明是唯一會伸縮的一欄，剩多少寬度給它多少。
        # 六欄全部可伸縮。原本只有「說明」會伸縮，其餘五欄寬度寫死，加起來
        # wide(837) 比表格實際拿得到的寬度還大 —— 於是最右邊的欄位被切在畫面外，
        # 而且這張表沒有橫向捲軸，捲不出來。被切掉的偏偏是「狀態」跟「說明」：
        # 這一格程式會不會覆蓋、為什麼不覆蓋，答案全在那兩欄裡。字級調大或把中間
        # 那條分隔線往右拖，切掉的就更多。
        #
        # minwidth 是「還讀得出來」的下限，不是好看的下限 —— 擠到極限時寧可看到
        # 「175,000 → 18…」，也不要一整欄消失。六欄的下限加起來 wide(497)，比視窗
        # 縮到最小時右半邊拿得到的寬度還小，所以任何尺寸下六欄都留在畫面上。
        self.tree = ttk.Treeview(box, columns=[c[0] for c in DETAIL_COLUMNS],
                                 show="headings", selectmode="browse")
        # 寬度一變就重新攤一次（拖分隔線、改視窗大小都算）。
        build_columns(self.tree, DETAIL_COLUMNS)

        bar = ttk.Scrollbar(box, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=bar.set)
        self.tree.grid(row=1, column=0, sticky="nsew")
        bar.grid(row=1, column=1, sticky="ns")

        # 顏色只用來輔助，不是唯一線索 —— 狀態欄本身就寫著字。
        self.tree.tag_configure("write", background="#eaf4ea")
        self.tree.tag_configure("done", background="#eaf0fb")
        self.tree.tag_configure("manual", foreground="#a34a00")
        self.tree.tag_configure("untracked", foreground="#777777")

        # 現金自己一條，不塞進上面那張表。Excel 裡它本來就在「持股資料」那張表
        # 外面（B8），而且它的網頁值是「今日淨收付」不是餘額 —— 跟股數成本不同義，
        # 擺進同一欄底下會讓人以為那是網頁上的餘額。
        #
        # 用 Text 而不是 Label：Label 只能整行一個顏色，於是「553,161 → -13,877,373」
        # 這種一正一負的行只能整行紅（正數被染紅在說謊）或整行不紅（負數看不出來）。
        # Text 可以一段一段上色，負的那個數字自己紅，其餘照常。
        # 看起來要像一行字、不像輸入框：跟著面板底色、拿掉外框、關掉游標。
        panel = ttk.Style().lookup("TFrame", "background") or self.root.cget("background")
        self.cash_line = tk.Text(box, height=1, wrap="word", font=(self.family, FONT_SIZE),
                                 background=panel, relief="flat", highlightthickness=0,
                                 insertwidth=0, cursor="arrow", state="disabled")
        self.cash_line.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(6, 0))
        self.cash_line.tag_configure("dim", foreground="#666666")
        self.cash_line.tag_configure("neg", foreground="#c00000")
        # 「餘額轉負」跟負數同一個紅 —— 它講的就是那個負數，兩種顏色會讓人
        # 以為是兩件事。
        self.cash_line.tag_configure("turned", foreground="#c00000")
        # 拖動中間那條分隔線會改變寬度，說明就可能從一行變兩行 —— 寬度變了就重量一次。
        self.cash_line.bind("<Configure>", lambda _event: self._fit_cash_line())

        # 現金那一條底下再貼一條：今天是從多少錢開始的。
        #
        # 餘額不是網頁抄來的，是「今日初始現金餘額 + 今日淨收付」算出來的，所以
        # 餘額不對的時候要改的是這個數字。它一直只存在紀錄檔裡，畫面上看不到，
        # 而唯一改得到它的入口是程式自己判斷該不該跳的那個對話框 —— 沒跳、
        # 或跳的時候填錯，就得等明天。擺一個固定的位置給它就沒有這種時段了。
        #
        # 貼在現金那一條正下方，因為要看的是兩者的關係（初始 + 今日 = 現在）。
        opening = ttk.Frame(box)
        opening.grid(row=3, column=0, columnspan=2, sticky="ew", pady=(2, 0))
        ttk.Label(opening, text="今日初始現金餘額", style="Hint.TLabel").pack(side="left")
        self.opening_value = ttk.Label(opening, text="")
        self.opening_value.pack(side="left", padx=(8, 0))
        self.opening_button = ttk.Button(opening, text="修改", width=6, command=self.edit_opening)
        self.opening_button.pack(side="left", padx=(12, 0))
        # 按鈕變灰的時候，理由就寫在旁邊 —— 灰掉而不說為什麼，看起來就像壞了。
        self.opening_hint = ttk.Label(opening, text="", style="Hint.TLabel")
        self.opening_hint.pack(side="left", padx=(12, 0))

        self.warn_box = tk.Text(box, height=5, wrap="word", font=(self.family, HINT_SIZE),
                                background="#fbfbfb", relief="flat", state="disabled")
        self.warn_box.grid(row=4, column=0, columnspan=2, sticky="ew", pady=(8, 0))

        # 換人只靠名單本身。原本這裡還有「上一位／下一位」兩顆按鈕，但名單就在
        # 左邊、點下去更快，那兩顆只是同一件事的第二個入口；「寫入」也拿掉了
        # ——「程式自動更新」勾了就會自己寫，沒勾就是人自己維護，不需要第三條路。
        # 快速鍵綁在視窗上而不是名單上：手放在右邊那張表的時候焦點不在名單裡，
        # 這時候還是要能換人。
        self.root.bind("<Control-Down>", lambda _event: self._step_person(1))
        self.root.bind("<Control-Up>", lambda _event: self._step_person(-1))

        # 持股撐死五六檔，表格卻吃掉整片高度的話，現金那一條會被推到視窗最底下
        # ——它是每天最先要看的一個數字，不該離持股表那麼遠。所以表格改成照列數
        # 決定高度（見 _fill_detail），多出來的空白由最下面那一列吸收。
        box.rowconfigure(1, weight=0)
        box.rowconfigure(5, weight=1)
        box.columnconfigure(0, weight=1)

    def _build_history_tab(self):
        frame = ttk.Frame(self.tabs, padding=8)
        self.tabs.add(frame, text="  歷程  ")

        # 篩選列。20 個帳號、每個帳號上百格，整份歷程一路捲下去是找不到東西的；
        # 而人真正要問的永遠是「某個人的某一格，今天發生了什麼」，
        # 所以選單就照那句話拆成兩個：交易人、項目。
        picks = ttk.Frame(frame)
        picks.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 8))

        ttk.Label(picks, text="交易人").pack(side="left")
        self.history_who = ttk.Combobox(picks, state="readonly", width=20,
                                        font=(self.family, FONT_SIZE), values=[ALL_CHOICE])
        self.history_who.current(0)
        self.history_who.pack(side="left", padx=(6, 20))
        self.history_who.bind("<<ComboboxSelected>>", self._on_history_who)

        ttk.Label(picks, text="項目").pack(side="left")
        self.history_item = ttk.Combobox(picks, state="readonly", width=28,
                                         font=(self.family, FONT_SIZE), values=[ALL_CHOICE])
        self.history_item.current(0)
        self.history_item.pack(side="left", padx=(6, 20))
        self.history_item.bind("<<ComboboxSelected>>", lambda _event: self._fill_history())

        ttk.Label(picks, text="期間").pack(side="left")
        self.history_when = ttk.Combobox(picks, state="readonly", width=10,
                                         font=(self.family, FONT_SIZE),
                                         values=[ALL_CHOICE, WHEN_TODAY, WHEN_WEEK])
        self.history_when.current(0)
        self.history_when.pack(side="left", padx=(6, 0))
        self.history_when.bind("<<ComboboxSelected>>", lambda _event: self._fill_history())

        self.history_tree = ttk.Treeview(frame, columns=[c[0] for c in HISTORY_COLUMNS],
                                         show="headings")
        build_columns(self.history_tree, HISTORY_COLUMNS)

        bar = ttk.Scrollbar(frame, orient="vertical", command=self.history_tree.yview)
        self.history_tree.configure(yscrollcommand=bar.set)
        self.history_tree.grid(row=1, column=0, sticky="nsew")
        bar.grid(row=1, column=1, sticky="ns")

        self.history_tree.tag_configure("human", foreground="#a34a00")
        self.history_tree.tag_configure("adopt", foreground="#1a5fb4")

        row = ttk.Frame(frame)
        row.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(8, 0))
        self.history_hint = ttk.Label(row, text="", style="Hint.TLabel")
        self.history_hint.pack(side="left")
        ttk.Button(row, text="重新整理", command=self.refresh_history).pack(side="right")
        # 清除擺在重新整理左邊，中間留一段空白：兩顆按鈕一顆是「再讀一次」、
        # 一顆是「全部收走」，手滑按錯的代價差太多，不能並排貼在一起。
        self.history_clear = ttk.Button(row, text="清除歷程", command=self.clear_history)
        self.history_clear.pack(side="right", padx=(0, 16))

        frame.rowconfigure(1, weight=1)
        frame.columnconfigure(0, weight=1)

    def _build_cert_tab(self):
        """
        把 setup-profile.ps1（建立/重建使用者資料夾）與 migrate-cert.ps1（掃描、複製憑證）
        整合進來，外加登入時順便抓到的憑證到期日 —— 這三件事本來都要開 PowerShell 手動跑，
        現在收進同一個分頁，按鈕按下去就是了。
        """
        frame = ttk.Frame(self.tabs, padding=8)
        self.tabs.add(frame, text="  憑證  ")

        self._build_profile_section(frame)
        self._build_migrate_section(frame)
        self._build_cert_status_section(frame)

        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(2, weight=1)

    def _build_profile_section(self, parent):
        """建立／重建自動登入用的 Chrome 使用者資料夾（對應 setup-profile.ps1）。"""
        box = ttk.LabelFrame(parent, text="Profile", padding=8)
        box.grid(row=0, column=0, sticky="ew", pady=(0, 8))

        self.profile_status = ttk.Label(box, text="", style="Hint.TLabel")
        self.profile_status.grid(row=0, column=0, columnspan=4, sticky="w")

        ttk.Label(box, text="Profile 名稱").grid(row=1, column=0, sticky="w", pady=(6, 0))
        self.profile_name = tk.StringVar(value=profile_tools.current_raw())
        ttk.Entry(box, textvariable=self.profile_name, width=28,
                  font=(self.family, FONT_SIZE)).grid(row=1, column=1, sticky="w",
                                                       padx=(8, 8), pady=(6, 0))

        self.profile_button = ttk.Button(box, text="建立 Profile", command=self.create_profile)
        self.profile_button.grid(row=1, column=2, sticky="w", pady=(6, 0))

        ttk.Label(box, text="沒填就用 chrome-profile。會開一個 Chrome 視窗把資料夾初始化，"
                            "完成後自動關掉；資料夾已存在的話（沒有憑證就直接清空重建，"
                            "有憑證會先問過你）。",
                  style="Hint.TLabel", wraplength=wide(760)).grid(
            row=2, column=0, columnspan=4, sticky="w", pady=(4, 0))

        box.columnconfigure(3, weight=1)
        self._refresh_profile_status()

    def _build_migrate_section(self, parent):
        """從平常在用的 Chrome/Edge 找出已經申請過的憑證，複製到自動登入用的 Profile（對應 migrate-cert.ps1）。"""
        box = ttk.LabelFrame(parent, text="遷移憑證", padding=8)
        box.grid(row=1, column=0, sticky="ew", pady=(0, 8))

        head = ttk.Frame(box)
        head.grid(row=0, column=0, sticky="ew")
        ttk.Button(head, text="掃描", command=self.scan_cert_sources).pack(side="left")
        self.migrate_copy_button = ttk.Button(head, text="複製到自動登入用的 Profile",
                                              command=self.copy_selected_cert, state="disabled")
        self.migrate_copy_button.pack(side="left", padx=(8, 0))
        self.migrate_status = ttk.Label(head, text="", style="Hint.TLabel")
        self.migrate_status.pack(side="left", padx=(16, 0))

        columns = ("browser", "name", "found", "path")
        titles = {"browser": "瀏覽器", "name": "Profile", "found": "憑證", "path": "路徑"}
        widths = {"browser": 70, "name": 90, "found": 60, "path": 380}
        self.migrate_tree = ttk.Treeview(box, columns=columns, show="headings",
                                         height=6, selectmode="browse")
        for key in columns:
            self.migrate_tree.heading(key, text=titles[key])
            anchor = "center" if key == "found" else "w"
            self.migrate_tree.column(key, width=wide(widths[key]), minwidth=wide(widths[key] // 2),
                                     anchor=anchor, stretch=(key == "path"))
        self.migrate_tree.grid(row=1, column=0, sticky="ew", pady=(6, 0))
        self.migrate_tree.bind("<<TreeviewSelect>>", self._on_migrate_select)
        self.migrate_tree.tag_configure("found", background="#eaf4ea")

        box.columnconfigure(0, weight=1)

    def _build_cert_status_section(self, parent):
        """登入過的每一位交易人，憑證什麼時候到期（見 fetch._fetch_cert_status）。"""
        box = ttk.LabelFrame(parent, text="登入帳號的憑證到期日", padding=8)
        box.grid(row=2, column=0, sticky="nsew")

        columns = ("name", "expiry", "state")
        titles = {"name": "交易人", "expiry": "憑證到期日", "state": "狀態"}
        widths = {"name": 120, "expiry": 170, "state": 100}
        self.cert_tree = ttk.Treeview(box, columns=columns, show="headings", selectmode="browse")
        for key in columns:
            self.cert_tree.heading(key, text=titles[key])
            anchor = "center" if key == "state" else "w"
            self.cert_tree.column(key, width=wide(widths[key]), minwidth=wide(widths[key] // 2),
                                  anchor=anchor, stretch=(key == "name"))
        bar = ttk.Scrollbar(box, orient="vertical", command=self.cert_tree.yview)
        self.cert_tree.configure(yscrollcommand=bar.set)
        self.cert_tree.grid(row=0, column=0, sticky="nsew")
        bar.grid(row=0, column=1, sticky="ns")
        self.cert_tree.tag_configure("expired", foreground="#c00000")
        self.cert_tree.tag_configure("soon", foreground="#a34a00")

        ttk.Label(box, text="登入時才會抓到；還沒登入過的人這裡不會出現。",
                  style="Hint.TLabel").grid(row=1, column=0, columnspan=2, sticky="w", pady=(4, 0))

        box.rowconfigure(0, weight=1)
        box.columnconfigure(0, weight=1)

    # ---------- 憑證分頁 ----------

    def _refresh_profile_status(self):
        raw = profile_tools.current_raw()
        path = profile_tools.resolve_path(raw)
        if path is None:
            self.profile_status.configure(text="目前 .env 沒有設定 USER_DATA_DIR —— 憑證存不住。")
        elif not path.is_dir():
            self.profile_status.configure(text=f"目前的 Profile：{path}（還沒建立）")
        elif profile_tools.has_cert(path):
            self.profile_status.configure(text=f"目前的 Profile：{path}（已經有 tbbstock 憑證）")
        else:
            self.profile_status.configure(text=f"目前的 Profile：{path}（還沒有 tbbstock 憑證）")

    def create_profile(self):
        """
        建立（或清空重建）自動登入用的 Chrome 使用者資料夾：開一個一般模式的 Chrome
        把資料夾初始化出來，確認初始化完成就自動關掉視窗，不必像 setup-profile.ps1
        那樣手動把視窗關掉。對應 2.1～2.3。
        """
        if self.profile_busy:
            return

        raw = self.profile_name.get().strip() or profile_tools.DEFAULT_NAME
        path = profile_tools.resolve_path(raw)
        if path is None:
            messagebox.showerror("名稱不能是空的", "請輸入 Profile 名稱。", parent=self.root)
            return
        if profile_tools.is_default_chrome_dir(path):
            messagebox.showerror(
                "不能用這個資料夾",
                "不能指向 Chrome 的預設使用者資料夾，Chrome 136 之後禁止自動化連上它。\n"
                "請換一個名稱（例如 chrome-profile）。", parent=self.root)
            return

        if path.is_dir():
            if profile_tools.chrome_pids_for_profile(path):
                messagebox.showerror(
                    "資料夾正在使用中",
                    f"這個資料夾正被 Chrome 開著，請先把它關掉再試一次：\n{path}", parent=self.root)
                return
            if profile_tools.has_cert(path) and not messagebox.askyesno(
                "資料夾裡已經有憑證",
                f"這個資料夾裡有 tbbstock 的數位憑證：\n{path}\n\n"
                "刪掉重建的話，那張憑證就沒了，之後要重新申請一次。\n確定要刪除重建嗎？",
                icon="warning", default="no", parent=self.root,
            ):
                return
            try:
                profile_tools.delete_profile(path)
            except OSError as exc:
                messagebox.showerror("刪不掉", f"刪除資料夾失敗：\n{exc}", parent=self.root)
                return

        chrome_exe = profile_tools.find_chrome()
        if chrome_exe is None:
            messagebox.showerror("找不到 Chrome", "這台電腦找不到 Google Chrome，請先安裝。",
                                 parent=self.root)
            return

        try:
            profile_tools.remember_user_data_dir(raw)
        except OSError as exc:
            messagebox.showwarning("沒寫進 .env", f"這次可以用，但沒能寫進 .env：\n{exc}",
                                   parent=self.root)

        try:
            profile_tools.launch_manual_chrome(chrome_exe, path)
        except OSError as exc:
            messagebox.showerror("開不起來", f"沒辦法開啟 Chrome：\n{exc}", parent=self.root)
            return

        self.profile_busy = True
        self.profile_button.configure(state="disabled")
        self.profile_status.configure(text=f"正在建立 Profile：{path}（Chrome 開起來了，請稍候…）")
        self._poll_profile_init(path, 0)

    def _poll_profile_init(self, path, tries):
        """每半秒確認一次資料夾初始化完成沒，完成就自動把那個 Chrome 視窗關掉（2.3）。"""
        if profile_tools.profile_initialized(path):
            profile_tools.kill_pids(profile_tools.chrome_pids_for_profile(path))
            self.profile_busy = False
            self.profile_button.configure(state="normal")
            self._refresh_profile_status()
            messagebox.showinfo(
                "Profile 建好了",
                f"資料夾已就緒，視窗已經自動關掉：\n{path}\n\n"
                "接下來可以用下面的「遷移憑證」把憑證複製進來；掃不到的話就是還沒申請過，"
                "自己開這個資料夾登入 tbbstock 申請一次即可。", parent=self.root)
            return
        if tries >= 60:   # 30 秒
            self.profile_busy = False
            self.profile_button.configure(state="normal")
            self.profile_status.configure(text=f"還沒看到 Profile 初始化完成：{path}")
            messagebox.showwarning(
                "沒看到初始化完成",
                f"30 秒過去了，還沒看到 Chrome 把資料夾初始化好：\n{path}\n\n"
                "視窗可能還在開啟中，請自己看一下，關掉後再按一次「建立 Profile」確認結果。",
                parent=self.root)
            return
        self.root.after(500, lambda: self._poll_profile_init(path, tries + 1))

    def scan_cert_sources(self):
        """掃描這台電腦上 Chrome／Edge 的每個 profile，看誰有 tbbstock 憑證痕跡（2.4）。"""
        target = profile_tools.resolve_path(profile_tools.current_raw())
        self.migrate_tree.delete(*self.migrate_tree.get_children())
        self._migrate_candidates = {}
        self.migrate_copy_button.configure(state="disabled")

        if target is None:
            self.migrate_status.configure(text="還沒設定 USER_DATA_DIR。")
            return

        candidates = profile_tools.scan_cert_sources(target)
        # 沒憑證痕跡的 profile 只是雜訊 —— 這張表是給人挑「要從哪一個複製」，
        # 不是給人看「這台電腦裝了幾個 profile」，所以只列有痕跡的那幾個。
        found_candidates = [c for c in candidates if c["found"]]
        for candidate in found_candidates:
            item = self.migrate_tree.insert(
                "", "end",
                values=(candidate["browser"], candidate["name"], "有", str(candidate["path"])),
                tags=("found",),
            )
            self._migrate_candidates[item] = candidate

        if found_candidates:
            self.migrate_status.configure(
                text=f"掃了 {len(candidates)} 個 profile，{len(found_candidates)} 個有憑證痕跡。")
        elif candidates:
            self.migrate_status.configure(text=f"掃了 {len(candidates)} 個 profile，沒有找到憑證痕跡。")
        else:
            self.migrate_status.configure(text="這台電腦上沒找到 Chrome／Edge 的 profile。")

    def _on_migrate_select(self, _event=None):
        selection = self.migrate_tree.selection()
        ok = bool(selection) and self._migrate_candidates.get(selection[0], {}).get("found")
        self.migrate_copy_button.configure(state="normal" if ok else "disabled")

    def copy_selected_cert(self):
        """把選中的來源 profile 的憑證複製到自動登入用的 Profile。"""
        selection = self.migrate_tree.selection()
        if not selection:
            return
        source = self._migrate_candidates.get(selection[0])
        if not source or not source["found"]:
            return

        target = profile_tools.resolve_path(profile_tools.current_raw())
        if target is None or not target.is_dir():
            messagebox.showerror(
                "目標 Profile 還不存在",
                "請先用上面的「建立 Profile」把資料夾建出來，再回來複製憑證。", parent=self.root)
            return

        if profile_tools.browser_running(source["exe"]):
            messagebox.showerror(
                "來源瀏覽器還開著",
                f"{source['browser']} 還在執行，複製到的可能是還沒寫進磁碟的舊資料。\n"
                f"請把 {source['browser']} 所有視窗（含背景程序）都關掉再試一次。", parent=self.root)
            return
        if profile_tools.chrome_pids_for_profile(target):
            messagebox.showerror(
                "目標 Profile 正在使用中",
                f"這個資料夾正被 Chrome 開著，請先關掉：\n{target}", parent=self.root)
            return

        # target 是 USER_DATA_DIR 本身，但 Chrome 實際把資料存在它底下的 Default（或
        # BROWSER_PROFILE_DIR）子資料夾裡 —— copy_cert 要對到那一層，不是 USER_DATA_DIR 自己。
        profile_dir = target / profile_tools.profile_subdir_name()
        if not profile_dir.is_dir():
            messagebox.showerror(
                "Profile 還沒初始化",
                f"這個資料夾裡還沒有 {profile_tools.profile_subdir_name()} 子資料夾：\n{target}\n\n"
                "請先用上面的「建立 Profile」把資料夾初始化，再回來複製憑證。", parent=self.root)
            return

        if not messagebox.askyesno(
            "複製憑證",
            f"把「{source['browser']} / {source['name']}」的憑證複製到自動登入用的 Profile？\n\n"
            "目標現有的資料會先備份。來源那張憑證不受影響（這只是複製檔案，不是重新申請）。",
            parent=self.root,
        ):
            return

        try:
            backup = profile_tools.copy_cert(source["path"], profile_dir)
        except OSError as exc:
            messagebox.showerror("複製失敗", str(exc), parent=self.root)
            return

        self._refresh_profile_status()
        self.scan_cert_sources()
        note = f"（原本的已備份到 {backup.name}）" if backup else ""
        self.migrate_status.configure(text=f"已複製完成{note}")
        messagebox.showinfo(
            "複製完成",
            f"憑證已複製到自動登入用的 Profile{note}。\n"
            "接下來按「登入」實際驗證，不再跳「瀏覽器查無有效數位憑證」就是成功了。", parent=self.root)

    def _update_cert_status(self, records):
        """
        把這批登入結果裡的憑證到期日記起來，畫進「憑證」分頁；快到期或已過期的人跳一次提醒
        （同一個工作階段只提醒一次，不必每讀一次資料就再煩一次，見 2.5）。
        """
        alerts = []
        for record in records:
            name = record.get("sheet_name")
            if not name or record.get("cert_text") is None:
                continue
            raw_expiry = record.get("cert_expiry")
            expiry = datetime.datetime.fromisoformat(raw_expiry) if raw_expiry else None
            self.cert_status[name] = {"text": record["cert_text"], "expiry": expiry}
            level = profile_tools.cert_alert_level(expiry)
            if level and name not in self.cert_alerted:
                self.cert_alerted.add(name)
                alerts.append((name, level, expiry))

        self._refresh_cert_tree()

        if alerts:
            lines = [f"・{name}：{'已過期' if level == 'expired' else '快到期'}"
                    f"（{expiry.strftime('%Y/%m/%d') if expiry else '?'}）"
                    for name, level, expiry in alerts]
            messagebox.showwarning(
                "憑證快到期了",
                "以下交易人的 tbbstock 數位憑證：\n\n" + "\n".join(lines) +
                "\n\n請提醒本人去 tbbstock 重新申請憑證。", parent=self.root)

    def _refresh_cert_tree(self):
        self.cert_tree.delete(*self.cert_tree.get_children())
        now = datetime.datetime.now()
        for name in sorted(self.cert_status):
            info = self.cert_status[name]
            expiry = info.get("expiry")
            level = profile_tools.cert_alert_level(expiry, now)
            expiry_text = expiry.strftime("%Y/%m/%d %H:%M") if expiry else (info.get("text") or "抓不到")
            state = {"expired": "已過期", "soon": "即將到期"}.get(level, "正常" if expiry else "未知")
            self.cert_tree.insert("", "end", values=(name, expiry_text, state),
                                  tags=(level,) if level else ())

    # ---------- 背景工作 ----------

    def _selected_accounts(self):
        choice = self.account_choice.current()
        numbered = list(enumerate(self.accounts, start=1))
        return numbered if choice <= 0 else [numbered[choice - 1]]

    def _account_choices(self):
        """範圍選單上的每一列。名字知道了就寫名字 —— 20 組時光看「第 7 組」不知道是誰。"""
        choices = ["全部"]
        for i, account in enumerate(self.accounts, start=1):
            name = self.trader_of.get(i)
            label = f"第 {i} 組"
            if name:
                label += f"　{name}" + ("（模擬）" if account.get("fake") else "")
            choices.append(label)
        return choices

    def _refresh_account_choices(self):
        """登入或讀取之後名字才知道，選單上那幾列要跟著補上去。選中的那一列不動。"""
        keep = max(self.account_choice.current(), 0)
        self.account_choice.configure(values=self._account_choices())
        self.account_choice.current(keep)
        self._refresh_fetch_button()

    def _scope_order(self):
        """這次要做第幾組；None 代表全部。"""
        choice = self.account_choice.current()
        return choice if choice > 0 else None

    def _scope_name(self):
        """這次要做的是哪一位交易人；全部、或名字還不知道時是 None。"""
        order = self._scope_order()
        return None if order is None else self.trader_of.get(order)

    def _refresh_fetch_button(self):
        """
        按鈕上的字就是「按下去會動到誰」。

        名字還不知道（那一組沒登入過）時寫「第 3 組」而不是硬掰一個名字：
        這種時候按下去確實只做那一組，只是程式還說不出他是誰。
        """
        order = self._scope_order()
        if order is None:
            text = "讀取網頁資料"
        else:
            name = self.trader_of.get(order)
            text = f"更新（{name}）" if name else f"讀取網頁資料（第 {order} 組）"
        self.fetch_button.configure(text=text)

    def _on_scope_changed(self, _event=None):
        """上面換了範圍，左邊名單也跳到那一位 —— 兩邊是同一個選擇的兩個入口。"""
        self._refresh_fetch_button()
        name = self._scope_name()
        if name and name != self.current_sheet and name in self._shown():
            self.current_sheet = name
            self.people.selection_set(name)
            self.people.see(name)
            self._fill_detail()

    def _sync_scope_to_person(self):
        """
        左邊名單換人，上面的範圍跟著換成他。

        對不到帳號時範圍不動（名單是網頁資料長出來的，正常情況一定對得到）
        —— 硬把範圍留在別人身上也不會說謊，按鈕上寫的一直是範圍裡的那個人。
        """
        order = next((i for i, name in self.trader_of.items() if name == self.current_sheet), None)
        if order is not None and order <= len(self.accounts):
            if self.account_choice.current() != order:
                self.account_choice.current(order)
        self._refresh_fetch_button()

    def _ensure_browser_thread(self):
        if self.browser_thread is None or not self.browser_thread.is_alive():
            self.browser_thread = threading.Thread(target=self._browser_worker, daemon=True)
            self.browser_thread.start()

    def start_login(self):
        if self.busy or not self._require_excel():
            return
        if not self.accounts:
            messagebox.showerror("沒有帳號", "請先在 .env 填入 TBB_ID_1 / TBB_PASSWORD_1。")
            return
        if self.ledger_error:
            messagebox.showerror("紀錄檔有問題", self.ledger_error)
            return

        self._ensure_browser_thread()
        self._set_busy(True, "登入中，瀏覽器會自己開起來，請不要關掉它…")
        self.browser_waiting += 1
        self.browser_cmd_queue.put(("login", (self._selected_accounts(), self.path)))

    def start_fetch(self):
        if self.busy or not self._require_excel():
            return
        if not self.accounts:
            messagebox.showerror("沒有帳號", "請先在 .env 填入 TBB_ID_1 / TBB_PASSWORD_1。")
            return
        if self.ledger_error:
            messagebox.showerror("紀錄檔有問題", self.ledger_error)
            return

        # 這一輪要做的是誰，按下去的當下就記起來：等結果回來的這幾十秒裡，
        # 使用者隨時可能在左邊名單上點別人（範圍會跟著換），報告卻是在講剛才那一輪。
        who = self.round_target = self._scope_name()
        self._ensure_browser_thread()
        self._set_busy(True, f"讀取{f'（{who}）' if who else ''}中，"
                             f"還沒登入的話瀏覽器會自己開起來，請不要關掉它…")
        self.browser_waiting += 1
        self.browser_cmd_queue.put(("fetch", (self._selected_accounts(), self.path)))

    def _browser_worker(self):
        """
        背景：整個瀏覽器 session 的生命週期都在這個執行緒裡，一直活到使用者自己把
        瀏覽器關掉，或整個介面關閉為止。

        每組帳號的分頁與 cookie 都收在這個執行緒手上的 store 裡（見 fetch.new_store）
        —— 「只更新某一位」能不重登就查得到資料，靠的就是它活得跟瀏覽器一樣久。

        不能每次按鈕都開新執行緒各開各的瀏覽器 —— Playwright 的同步 API 底層用
        greenlet 綁死建立它的那個執行緒，換一個執行緒去操作同一個 context 會直接
        壞掉（cannot switch to a different thread）。所以瀏覽器只能養在「一個專屬、
        活得夠久」的執行緒裡，「登入」「讀取網頁資料」都只是丟一個指令進 queue
        給它處理，差別只在要不要順便查資料、更新 Excel。

        每一個指令都一定要回一則結果，這個執行緒也一定不能因為某次失敗就結束。
        主執行緒是收到那則結果才解除「登入中…」的 —— 不回話等於介面永遠卡在那裡，
        登入與讀取兩顆按鈕都是灰的，使用者只能關掉重開。所以 Playwright 改用
        start() 而不是 with：連「它自己起不來」（瀏覽器沒裝好、profile 正被另一個
        Chrome 佔著）都要變成一則看得到的「登入失敗」，而不是一個在背景執行緒裡
        炸掉、誰也看不到的例外。裝好瀏覽器之後再按一次就會重試，不必重開程式。

        登入與讀取只差在呼叫哪一個函式、回話時說自己是哪一種，所以走同一段程式碼
        —— 錯誤處理只有一份，不會有「登入那條路修好了、讀取那條還留著舊寫法」。
        """
        playwright = context = browser = None
        # 每組帳號的分頁、cookie、身分，還有「現在瀏覽器帶著誰的 cookie」。
        # 「一次只更新一位」全靠它：換人時把那一組登入時收下來的 cookie 換回去，
        # 不必重登（見 fetch.new_store）。跟著瀏覽器一起生、一起死。
        store = fetch_mod.new_store()

        def ensure_browser():
            nonlocal playwright, context, browser, store
            if playwright is None:
                # 瀏覽器的位置要在 driver 起來之前決定好，它是靠環境變數傳下去的。
                configure_browsers_path()
                playwright = sync_playwright().start()
            if context is not None and not _browser_alive(context):
                context = browser = None      # 使用者自己把它關掉了，重開一個
            if context is None:
                context, browser = open_context(playwright)
                store = fetch_mod.new_store()

        # 指令 -> (要呼叫誰, 回話時說自己是哪一種)
        jobs = {"login": (login_only, "logged_in"), "fetch": (collect, "fetched")}

        try:
            while True:
                cmd, arg = self.browser_cmd_queue.get()
                if cmd == "stop":
                    break

                fetch_records, kind = jobs[cmd]
                selected, path = arg
                try:
                    ensure_browser()
                    # 登入完也順便把 Excel 現值讀出來，主執行緒要拿它當程式的起點。
                    records = fetch_records(context, selected, store)
                    payload = _read_excel_after_fetch(records, path)
                except Exception as exc:
                    payload = {"error": traceback.format_exc()}
                    if excel_io.is_dead_object(exc):
                        payload["hint"] = excel_io.DEAD_EXCEL_HINT
                    elif context is None:
                        # 連瀏覽器都還沒開起來就失敗，那是這台電腦的環境問題，
                        # 不是網站或帳號的問題，講清楚才不會有人去重打密碼。
                        payload["hint"] = BROWSER_HINT
                self.queue.put((kind, payload))
        finally:
            try:
                if context is not None:
                    context.close()
                if browser is not None:
                    browser.close()
                if playwright is not None:
                    playwright.stop()
            # 收尾失敗沒有任何補救的餘地，而這裡通常是「介面正在關閉」——
            # 為了關不乾淨的瀏覽器再丟一個例外出去只會蓋掉真正的死因。
            except Exception:
                pass

    def _collect_writes(self):
        """
        把提案整理成「分頁 -> 要寫哪幾格」。

        要寫哪幾格完全由 planner 的 will_write 決定，介面不再插手。以前介面會把
        沒勾的就地改成不寫，那是紀錄檔跟 Excel 對不起來的唯一破口 —— 現在沒有
        這條路了，寫進 Excel 的跟記進紀錄檔的必定是同一批。

        只收 round_scope 裡那幾位。名單上其他人可能也有「要寫」的格子，但那是
        用上一輪的網頁資料算出來的 —— 按「更新（王小明）」只會去查王小明，
        這時候順手把別人那幾格也寫進去，寫的是舊資料，而且沒有人要求過。
        """
        writes, total = {}, 0
        for name, items in self.proposals.items():
            if name not in self.round_scope:
                continue
            cells = [(item["row"], item["col"], item["proposed"])
                     for item in items if item["will_write"]]
            if cells:
                writes[name] = cells
                total += len(cells)
        return writes, total

    def _begin_write(self, writes, total):
        self.write_count = total
        self._set_busy(True, f"寫入 {total} 格…")
        threading.Thread(target=self._write_worker, args=(self.path, writes), daemon=True).start()

    def _write_worker(self, path, writes):
        import pythoncom

        payload = {}
        excel = workbook = sheet = None
        pythoncom.CoInitialize()
        try:
            # 備份在開 Excel 之前做：這時候檔案還沒有人開著，複製到的一定是
            # 乾淨的內容；而且萬一後面整段失敗，備份已經先拿到手了。
            saved = excel_io.backup(path)

            excel, workbook, attached = excel_io.open_workbook(path, True)
            try:
                for name, cells in writes.items():
                    sheet, error = excel_io.find_sheet(workbook, name)
                    if sheet is None:
                        raise RuntimeError(error)
                    excel_io.write_cells(sheet, cells)
                # 一律存檔，接上使用者開著的 Excel 時也一樣。
                #
                # 原本接上時刻意不存、留給人自己按 Ctrl+S 當作多一道確認，但那道
                # 確認換來的是一個更糟的破口：紀錄檔在寫入「成功」之後就記成寫過了，
                # 人只要沒按 Ctrl+S（或關檔時選「不要儲存」），帳本就跟檔案分家，
                # 而且畫面上不會有任何徵兆。反悔的路已經有更好的一條 —— 寫入前
                # 一定會備份。
                workbook.Save()
            finally:
                excel_io.close_workbook(excel, workbook, attached)
            payload = {"backup": str(saved), "attached": attached}
        except Exception as exc:
            payload = {"error": traceback.format_exc()}
            if excel_io.is_dead_object(exc):
                payload["hint"] = excel_io.DEAD_EXCEL_HINT
        finally:
            sheet = excel = workbook = None
            pythoncom.CoUninitialize()

        self.queue.put(("written", payload))

    def _drain(self):
        """
        主執行緒每 120ms 看一次背景有沒有結果。widget 只在這裡被碰。

        每一筆各自包一層 try，而且「排下一次」放在 finally 裡。這個迴圈是背景與
        畫面之間唯一的通道：處理某一筆時丟出例外的話，Tk 只會把 traceback 印在
        console（打包成 exe 之後連 console 都沒有），然後最後那行排程就跳過了 ——
        取件迴圈從此停擺，之後每一次登入、讀取、寫入都會停在「進行中…」，
        而背景其實早就做完、結果就躺在 queue 裡沒有人取。與其這樣，寧可把出錯的
        那一筆報成失敗，讓迴圈活下去。
        """
        handlers = {"logged_in": self._on_logged_in, "fetched": self._on_fetched,
                    "written": self._on_written}
        try:
            while True:
                kind, payload = self.queue.get_nowait()
                try:
                    handlers[kind](payload)
                except Exception:
                    self._on_handler_error(kind)
        except queue.Empty:
            pass
        finally:
            self._check_browser_thread()
            self.root.after(120, self._drain)

    def _on_handler_error(self, kind):
        """
        收到背景結果之後，自己在處理時出錯（多半是程式的臭蟲）。

        一定要解除 busy：出錯的是「收尾」那一步，使用者按下的那顆按鈕還在等回音，
        不解除的話兩顆按鈕會一直是灰的，跟當掉沒有兩樣。

        也一定要說出來。這種錯發生在「已經做了一半」的時候 —— 網頁資料讀完了、
        Excel 可能也寫過了 —— 所以除了 traceback，還要請人去對一下歷程，
        那是唯一能看出實際做了什麼的地方。
        """
        detail = traceback.format_exc()
        step = STEP_NAMES.get(kind, kind)
        self._set_busy(False)
        self._say(f"「{step}」的後續處理出錯")
        messagebox.showerror(
            "程式出錯",
            f"「{step}」的結果處理到一半出錯，這一批沒有做完。\n"
            f"Excel 與紀錄檔可能只完成了一部分，請切到「歷程」看實際做了哪幾格。\n\n"
            f"────────────────\n{detail[-1500:]}")

    def _check_browser_thread(self):
        """
        還在等背景回話、負責瀏覽器的那個執行緒卻已經不在了，就自己收尾。

        _browser_worker 已經把每一種失敗都包成一則結果回報了，這裡擋的是它自己
        也擋不住的那種（例如連 except 都還沒走完就整個結束）。沒有這道網的話，
        畫面會停在「登入中…」、登入與讀取都是灰的，使用者只能關掉重開 ——
        而關的時候還會被問一次「還在忙，確定要關嗎」。

        沒被取走的指令要一起倒掉。留著的話，下次按登入會先跑到那個舊指令
        （舊的檔案路徑、舊的帳號選擇），做了一件使用者這次沒有要求的事。
        """
        if not self.browser_waiting:
            return
        if self.browser_thread is not None and self.browser_thread.is_alive():
            return

        self.browser_waiting = 0
        while True:
            try:
                self.browser_cmd_queue.get_nowait()
            except queue.Empty:
                break

        self._set_busy(False)
        self._say("背景作業中斷，這次的動作沒有做完")
        messagebox.showerror(
            "背景作業中斷",
            "負責瀏覽器的背景作業結束了，這次的動作沒有做完。\n\n"
            f"{BROWSER_HINT}\n\n再按一次「登入」會重新開一個瀏覽器。")

    def _on_logged_in(self, payload):
        self.browser_waiting = max(0, self.browser_waiting - 1)
        self._set_busy(False)

        if "error" in payload:
            self._say("登入失敗")
            messagebox.showerror("登入失敗", _error_text(payload))
            return

        names, problems = [], []
        for record in payload["records"]:
            # 名字是登入才拿得到的東西，拿到就記著 —— 上面那個範圍選單、
            # 「更新（某某）」那顆按鈕都靠這份對照。
            if record.get("sheet_name"):
                self.trader_of[record["order"]] = record["sheet_name"]
            if record["problems"]:
                problems.append(f"第 {record['order']} 組：" + "；".join(record["problems"]))
            elif record.get("sheet_name"):
                names.append(record["sheet_name"])
        problems.extend(payload.get("sheet_errors", {}).values())
        self._refresh_account_choices()
        self._update_cert_status(payload["records"])

        # 只登入一組時，其他人上次讀到的 Excel 現值要留著，不能整份換掉。
        self.sheet_data.update(payload.get("sheets", {}))
        self.today = datetime.date.today()
        count = self._initialize(payload["records"])

        if problems:
            messagebox.showerror("登入失敗", "\n".join(problems))
            self._say("登入失敗")
            return

        done = f"，並以 Excel 現在的數字初始化了 {count} 格" if count else ""
        self._say(f"已登入：{'、'.join(names)}{done}。要更新資料時再按「讀取網頁資料」。")

    def _initialize(self, records):
        """
        登入成功的當下，把 Excel 上的股數、成本、現金餘額收成程式的起點。
        回傳收了幾格，狀態列要報這個數字。

        敢一句話都不問，是因為時間點：登入完成、還沒讀網頁資料之前，Excel 上的
        數字必定是「今天買賣之前」的狀態 —— 今天成交了什麼還在網頁那邊沒查。
        所以現金基準直接取 B8、今天的流水先記 0，等「讀取網頁資料」再往上加。
        「B8 含不含今天的淨收付」這個最容易答錯的問題，在這個時間點根本不存在。

        判斷全在 planner.initialize()：一天只設一次現金基準、已經是自動的格子
        不再重收，都在那裡，介面不另外複製一份規則。這裡只負責挑出「這一組登入
        成功、而且 Excel 那一頁也真的讀到了」的分頁 —— 登入失敗或找不到分頁的
        一律跳過，沒讀到的東西不能拿來當起點。
        """
        if self.ledger is None:
            return 0

        at = datetime.datetime.now().isoformat(timespec="seconds")
        events = []
        for record in records:
            name = record.get("sheet_name")
            data = self.sheet_data.get(name)
            if record["problems"] or not data:
                continue
            book = self.ledger.sheet(name)
            # 帳號代號在登入的當下就知道了，順手記進紀錄檔 ——
            # 使用者有可能登入完就沒有再按讀取。
            book["account_code"] = record.get("account_code", "")
            events.extend(planner.initialize(data, book, name, self.today, at))

        if not events:
            return 0

        # 順序跟寫入那邊一樣：紀錄檔先落地，再追加歷程。
        self.ledger.save()
        self.ledger.append_history(events)
        self.refresh_history()
        return len(events)

    def _cash_item(self, name):
        """某個分頁的現金那一列提案，沒有就 None。"""
        return next((item for item in self.proposals.get(name, [])
                     if item["kind"] == "cash"), None)

    def _on_fetched(self, payload):
        self.browser_waiting = max(0, self.browser_waiting - 1)
        self._set_busy(False)

        if "error" in payload:
            self._say("讀取失敗")
            messagebox.showerror("讀取失敗", _error_text(payload))
            return

        # 這一輪讀到的一律是「補上去」，不是「整份換掉」：一次只更新一位的時候，
        # 名單上其他人手上那份是上一輪的資料，清掉他們等於整份名單只剩一個人。
        # 留著的代價是畫面上同時有好幾個時間點的資料，所以每一位都記下讀取時間，
        # 右邊標頭寫得出「讀取於 10:32」（見 _fill_head）。
        now = datetime.datetime.now()
        errors = payload["sheet_errors"]
        fresh = []
        for record in payload["records"]:
            order, name = record["order"], record.get("sheet_name")
            if name:
                self.trader_of[order] = name

            if record["problems"]:
                problem = f"第 {order} 組：" + "；".join(record["problems"])
            elif not name:
                problem = f"第 {order} 組：讀不出這是哪一位交易人"
            elif name in errors:
                problem = errors[name]
            else:
                problem = None

            # 失敗原因跟著組別走：這一組這次成功就把上次的原因收掉，
            # 沒讀到的那幾組維持原樣 —— 別人的問題不會因為我這次讀成功就消失。
            if problem:
                self.problem_of[order] = problem
                continue
            self.problem_of.pop(order, None)
            self.records[name] = record
            self.read_at[name] = now
            fresh.append(name)

        # 這一輪只准碰這幾位。寫入、落帳、接管全部照它 —— 別人手上那份是舊資料，
        # 拿舊資料去寫 Excel 是「一次只更新一位」最貴的一種錯。
        self.round_scope = set(fresh)
        self.sheet_data.update(payload["sheets"])
        self._refresh_problems()
        self._refresh_account_choices()
        self._update_cert_status(payload["records"])
        self.today = datetime.date.today()
        self.replan()
        # 舊值要在任何寫入之前收好。寫入成功後 sheet_data 會被換成新數字，
        # 那時候再問「原本是多少」就沒有人記得了。只換這一輪讀到的那幾位：
        # 別人畫面上的「舊 → 新」是上一輪剛寫進去的結果，不該被這一輪抹掉。
        self.before = {key: value for key, value in self.before.items() if key[0] not in fresh}
        self.before.update({(name, item["cell"]): item["current"]
                            for name in fresh for item in self.proposals.get(name, [])})

        who = self.round_target
        note = self._problem_note()
        # 一位都沒讀成功的時候絕對不能說「已讀取」——「更新（王小明）」按下去、
        # 他那一組登入逾時，畫面上其他人的數字全都還在，最像結論的那一句要是
        # 寫著「已讀取」，看的人不會知道自己看的是半小時前的東西。
        if not self.round_scope:
            self._say((f"{who} 這一次沒讀到，什麼都沒做。" if who
                       else "這一輪沒有一位對照得起來，什麼都沒做。") + note)
            return

        head = f"已讀取（{who}）。" if who else "已讀取。"
        if payload.get("attached"):
            self._say(head + "這個 Excel 正開著，程式會直接接上那個視窗寫入並存檔。" + note)
        else:
            self._say(head + note)

        # 自動模式：讀完直接接著寫，中間不再問一次。按「讀取網頁資料」之前使用者
        # 就已經在那個開關上表達過意願了，再跳一個確認只是重複問同一件事。
        if self.auto_write.get():
            self._auto_adopt()
            writes, total = self._collect_writes()
            if total:
                self._begin_write(writes, total)
            else:
                # 一格都不必寫，不代表沒事發生：剛接管的格子、剛偵測到的人工改動、
                # 今天的淨收付，都是在這條路上落帳的。
                recorded = self._commit_round()
                self.replan()
                self.refresh_history()

                # 有人沒完成的時候，「一致」的範圍只到對照得起來的那幾位為止。
                # 主詞跟著縮小，這句話才不會替沒完成的那幾位背書。
                #
                # 看的是 proposals 不是 records：網頁讀到了、Excel 卻找不到那個
                # 分頁時 records 有東西、畫面上卻一位都沒有，這種時候說「一致」
                # 是把「沒得比」講成了「比過了」。
                scope = who if who else ("對照得起來的那幾位" if self.problems else "Excel 的數字")
                kept = f"紀錄檔更新了 {recorded} 筆（見歷程）。" if recorded else ""
                self._say(f"{head}{scope}跟網頁一致，沒有需要寫的格子。{kept}{note}")

    def _commit_round(self):
        """
        把這一輪的結果落實到紀錄檔：程式寫過的格子、偵測到的人工改動、現金的
        基準與流水。回傳追加了幾筆歷程。

        寫入成功之後一定要跑，而且只能在成功之後（見 _on_written）。但「一格都
        不必寫」的時候也一樣要跑 —— 那一輪照樣可能有話要記，最重要的一種是
        使用者剛在「重設現金餘額」填的答案：他填的開盤前金額加上今日淨收付，
        算出來剛好等於 B8 上的數字時（也就是 B8 早就含了今天的成交，正是那個
        對話框最該接住的情況），沒有任何一格需要寫。那個答案這時只存在提案裡，
        不落帳就會跟著整輪一起丟掉，下一次讀取拿沒被修正的基準再算一次，
        今天的淨收付就被加了第二次 —— 而畫面上不會有任何徵兆。
        """
        if self.ledger is None:
            return 0

        at = datetime.datetime.now().isoformat(timespec="seconds")
        events = []
        for name, items in self.proposals.items():
            # 跟 _collect_writes 同一個範圍。落帳記的是「這一輪發生了什麼」，
            # 沒去查的那幾位這一輪什麼也沒發生，尤其不能替他們記今天的淨收付。
            if name not in self.round_scope:
                continue
            book = self.ledger.sheet(name)
            events.extend(planner.commit(items, book, name, self.today, at))
        self.ledger.save()
        self.ledger.append_history(events)
        return len(events)

    def _on_written(self, payload):
        self._set_busy(False)

        if "error" in payload:
            self._say("寫入失敗")
            messagebox.showerror("寫入失敗", _error_text(payload))
            return

        # 紀錄檔一定在 Excel 寫成功之後才更新。順序反過來的話，寫入失敗會留下
        # 一份「以為自己寫過了」的帳本，之後每次比對都判定成人工改動。
        self._commit_round()

        # Excel 已經被改過了，手上的現值是舊的，重新讀一次才會準。
        # 只回填這一輪真的寫過的那幾位（範圍跟 _collect_writes 同一個）——
        # 別人那些「要寫」的格子這次沒寫進去，跟著改就會變成畫面說寫了、檔案沒有。
        for name, items in self.proposals.items():
            data = self.sheet_data.get(name)
            if not data or name not in self.round_scope:
                continue
            for item in items:
                if item["will_write"]:
                    if item["kind"] == "cash":
                        data["balance"] = item["proposed"]
                    else:
                        for line in data["rows"]:
                            if line["row"] == item["row"]:
                                line["qty" if item["which"] == "qty" else "cost"] = item["proposed"]

        self.replan()
        self.refresh_history()

        where = "（就寫進你開著的那個 Excel 視窗）" if payload["attached"] else ""
        # 刻意不跳視窗：一顆按鈕從頭做到尾，不要在最後又叫人按確定。程式做的
        # 變更沒辦法用 Ctrl+Z 復原，要反悔只能用備份，所以備份的位置要寫在
        # 狀態列上，人才找得到。
        # 提醒擺在備份路徑前面：路徑很長，接在後面的字會被擠出狀態列外面。
        self._say(f"已自動寫入 {self.write_count} 格並存檔{where}。"
                  f"{self._problem_note()}備份在 {payload['backup']}")

    def _refresh_problems(self):
        """
        把「第幾組 -> 失敗原因」攤平成畫面上那一串。

        原因記在組別上而不是每讀一次就整串重來：一次只更新一位的時候，別人上一輪
        沒完成的事並沒有因此解決 —— 那些 ⚠ 要留在畫面上，直到那一組自己再讀一次
        成功為止。照組別排序，畫面上的順序才不會每讀一次就跳一次。
        """
        self.problems = [self.problem_of[order] for order in sorted(self.problem_of)]

    def _problem_note(self):
        """
        這一輪有幾個帳戶沒完成。沒有就回空字串，直接串在狀態列後面。

        狀態列上任何一句「已讀取」「跟網頁一致」「已寫入」都要帶著它。20 組裡有
        一組登入逾時、或 Excel 少一個分頁時，那一位會直接從左邊名單上消失，而
        畫面上其他每一個數字看起來都正常 —— 提醒框裡雖然列著 ⚠，但狀態列同時
        說「跟網頁一致」的話，那句話會蓋過它：最像結論的那一句才是人會相信的。

        只報數量、不報細節。是哪幾組、為什麼失敗都在提醒框裡，狀態列只有一行，
        20 組全爛的時候塞不下，而且真正要看的細節本來就不該擠在那裡。
        """
        if not self.problems:
            return ""
        return f"⚠ 有 {len(self.problems)} 個帳戶沒完成，看下面的提醒。"

    def _set_busy(self, busy, message=""):
        self.busy = busy
        # 寫到一半改模式或換檔都沒有意義，這一輪要動哪個檔、寫不寫早就決定了。
        self.auto_check.configure(state="disabled" if busy else "normal")
        self.excel_button.configure(state="disabled" if busy else "normal")
        self._sync_clear_button()
        if busy:
            self.progress.pack(fill="x", padx=12, pady=(4, 0), before=self.tabs)
            self.progress.start(12)
            self._say(message)
        else:
            self.progress.stop()
            self.progress.pack_forget()
        self._sync_buttons()

    def _say(self, message):
        self.status.configure(text=message)

    def _path_text(self):
        if self.path is None:
            return "還沒選檔案 —— 按左邊「開啟EXCEL」挑一份"
        # 路徑後面直接講「現在開著沒」。登入是灰的時候，人第一個問題就是為什麼，
        # 答案擺在他正在看的那一行，比藏在狀態列或跳出來的視窗裡好找。
        return f"{self.path}　—— " + ("已開在 Excel 裡" if self.excel_open
                                      else "還沒開著，按「開啟EXCEL」")

    def open_excel(self):
        """
        選一份要同步的 Excel，用 Excel 把它開起來，路徑寫回 .env 下次直接用。

        為什麼是「開啟」不是「載入」
        --------------------------
        程式自己也開得起來（沒開著時 excel_io 會在背景 DispatchEx 一個隱形的），
        但 Office 沒啟用的機器上，那個背景 Excel 會在啟用檢查失敗後把自己收掉，
        同步是跑到一半才炸的 —— 那時備份做了、搞不好還寫了幾格。所以這裡改成
        由使用者親手把檔開起來，程式接上他那個視窗：看得見、讀到的是畫面上的
        即時內容、寫完他也馬上看得到。沒開起來就不給登入，把問題擋在最前面。

        紀錄檔（現金基準、每格自動／手動）是跟著檔名走的，所以換一份檔等於換掉
        整個狀態來源 —— 手上這批網頁資料與提案全是上一份檔算出來的，一律清掉
        重來，不能讓 A 檔的提案留在畫面上等著寫進 B 檔。
        """
        if self.busy:
            return

        start = self.path.parent if (self.path and self.path.parent.is_dir()) else app_dir()
        chosen = filedialog.askopenfilename(
            parent=self.root,
            title="選擇要同步的持股管理表",
            initialdir=str(start),
            initialfile=self.path.name if self.path else "",
            filetypes=[("Excel 活頁簿", "*.xls *.xlsx *.xlsm"), ("所有檔案", "*.*")],
        )
        if not chosen:
            return

        path = Path(chosen)
        # 選同一份檔不必重來一遍，但還是要確認它開著 —— 使用者按這顆按鈕，
        # 想要的就是「把它打開」，不是「什麼都沒發生」。
        if path == self.path:
            self._open_in_excel(path)
            return

        try:
            ledger = ledger_mod.Ledger(path)
        except RuntimeError as exc:
            messagebox.showerror("紀錄檔有問題", str(exc))
            return

        try:
            excel_io.remember_excel_path(path)
        except OSError as exc:
            messagebox.showwarning("沒寫進 .env",
                                   f"這次可以用，但路徑沒記起來，下次要再選一次：\n{exc}")
        self.path = path
        self.ledger = ledger
        self.ledger_error = None
        self.ledger_fresh = not ledger.existed
        # excel_open 記的還是上一份檔的答案（通常是 True），而輪詢要三秒後才會發現
        # 換人了。中間那幾秒畫面在說謊：路徑列寫著新檔的路徑、後面卻接「已開在
        # Excel 裡」，登入也還亮著。那時候按下去，程式會自己開一個看不見的 Excel
        # 來讀這個檔，而 os.startfile 正在把同一個檔開進使用者看得見的那個 Excel
        # —— 兩邊搶同一個檔案，總有一邊拿到唯讀。
        # 亮回來是 _open_in_excel 的事：它每半秒確認一次，真的看到開起來才放行。
        self._set_excel_open(False)
        self.path_label.configure(text=self._path_text())

        self.records, self.sheet_data, self.proposals = {}, {}, {}
        self.warnings, self.problems = {}, []
        self.before = {}
        # 這幾樣都是「上一份檔的這一輪」，跟著整批作廢。trader_of 例外：
        # 那是帳號與交易人的對照，跟開哪一份 Excel 無關。
        self.problem_of, self.read_at, self.round_scope = {}, {}, set()
        self.round_target = None
        self.current_sheet = None
        self.auto_write.set(bool(ledger.setting("auto_write", True)))

        self._refresh_mode_hint()
        self.fill_sync_tree()
        self.refresh_history()
        if not ledger.existed:
            self._say(f"已換成 {path.name}（這份檔還沒有紀錄檔，所有格子都要重新接管）")
        self._open_in_excel(path)

    def _open_in_excel(self, path, tries=0):
        """
        把檔案交給 Excel 打開，然後等它真的開起來。

        用 os.startfile 而不是 COM：這等於使用者自己雙擊那個檔，Excel 是他的、
        視窗是看得見的，程式手上不留任何 COM 物件。真要同步時 excel_io 會用
        GetObject 接上這個視窗。

        不能開下去就當作成功 —— startfile 只是把檔案丟給系統，Excel 起來要幾秒，
        中間還可能卡在受保護的檢視或啟用巨集的提示上。所以這裡每半秒問一次檔案鎖，
        真的看到它被鎖住才把登入放行。等待期間不能用 sleep 卡住主執行緒，
        不然整個介面會凍在那裡，連「正在開啟」那行字都畫不出來。
        """
        if path != self.path:
            return                        # 等待期間又換了檔，這條等待就作廢
        if excel_io.is_open_in_excel(path):
            self._set_excel_open(True)
            self._say(f"{path.name} 已經開在 Excel 裡，可以按「登入」了。")
            return

        if tries == 0:
            try:
                os.startfile(str(path))
            except OSError as exc:
                self._set_excel_open(False)
                messagebox.showerror("開不起來", f"沒辦法用 Excel 打開這個檔：\n{path}\n\n{exc}")
                return
            self._say(f"正在用 Excel 開啟 {path.name}…")
        if tries >= 60:                   # 30 秒
            self._set_excel_open(False)
            self._say(f"還沒看到 {path.name} 開起來。")
            messagebox.showwarning(
                "沒看到 Excel 開起來",
                f"{path.name} 交給 Excel 了，但 30 秒過去還沒看到它被開啟。\n\n"
                f"Excel 可能卡在「受保護的檢視」或啟用提示上，也可能還在啟動。\n"
                f"請看一下 Excel 那邊；等它真的開好，登入就會自己亮起來。",
                parent=self.root,
            )
            return
        self.root.after(500, lambda: self._open_in_excel(path, tries + 1))

    def _poll_excel(self):
        """
        每三秒確認一次「這份 Excel 現在開著沒」，然後排下一次。

        為什麼要一直問：使用者隨時可能把 Excel 關掉，而按鈕的亮暗說的是「現在」
        能不能做，不是「按開啟那一刻」的狀態。檔案鎖很便宜（開一個檔案代號就關掉），
        也不碰 COM，所以固定輪詢比在十幾個地方各補一次檢查可靠。
        """
        try:
            here = self.path is not None and self.path.is_file() and excel_io.is_open_in_excel(self.path)
        except OSError:
            here = self.excel_open        # 判斷不出來就沿用上次，不要亂閃
        self._set_excel_open(here)
        self.root.after(3000, self._poll_excel)

    def _set_excel_open(self, value):
        """狀態有變才動畫面 —— 每三秒重畫一次按鈕會讓游標懸停的效果一直閃。"""
        if value == self.excel_open:
            return
        self.excel_open = value
        self.path_label.configure(text=self._path_text())
        self._sync_buttons()

    def _require_excel(self):
        """登入／讀取前的最後一道關卡。按鈕平常是灰的，這裡擋的是鍵盤觸發那種漏網。"""
        if self.path is None:
            messagebox.showerror("還沒選檔案", "請先按左上角「開啟EXCEL」選一份持股管理表。")
            return False
        if not self.path.is_file():
            messagebox.showerror("找不到 Excel", f"{self.path}\n\n可以在 .env 用 EXCEL_PATH 指定位置。")
            return False
        if not self.excel_open:
            messagebox.showwarning("Excel 沒開著",
                                   f"請先按左上角「開啟EXCEL」把 {self.path.name} 打開。\n\n"
                                   f"程式要接上你那個 Excel 視窗才會同步。",
                                   parent=self.root)
            return False
        return True

    def _on_auto_changed(self):
        # 存進紀錄檔，下次打開沿用。
        if self.ledger is not None:
            try:
                self.ledger.set_setting("auto_write", bool(self.auto_write.get()))
            except OSError as exc:
                messagebox.showwarning("設定沒存起來", f"開關這次有效，但沒能寫進紀錄檔：\n{exc}")
        self._refresh_mode_hint()
        self.fill_sync_tree()

    def _refresh_mode_hint(self):
        if self.auto_write.get():
            self.mode_hint.configure(
                text="目前是「程式自動更新」：讀完網頁資料就會自動備份並寫進 Excel",
                style="Auto.TLabel")
        else:
            self.mode_hint.configure(
                text="目前是「人工維護」：程式不會動 Excel，下面只列給你對照著自己改",
                style="Manual.TLabel")

    # ---------- 同步分頁 ----------

    def replan(self):
        """重新計算提案並重畫。plan 是純函式，可以隨便重跑。"""
        self.proposals, self.warnings = {}, {}
        for name, record in self.records.items():
            data = self.sheet_data.get(name)
            if not data:
                continue
            book = self.ledger.sheet(name)
            book["account_code"] = record.get("account_code", "")
            items, warns = planner.plan(data, record, book, self.today)
            self.proposals[name] = items
            self.warnings[name] = warns

        self.fill_sync_tree()

    def _value_text(self, name, item, compare_web=True):
        """
        一格在畫面上要寫什麼字，外加它是不是這一輪剛被寫過。

        沒事就只寫現在的數字 —— 一格擺出舊值、網頁值、新值三個數字，真正要看的
        那一個反而被埋掉了。有話要說的時候才寫兩個：

            1,000 → 2,000        等著寫的，或這一輪剛寫進去的
            1,000（網頁 2,000）  跟網頁不一樣，但程式不會動它（手動／未接管）
            1,000                跟網頁一致，這一輪也沒動過

        舊值取的是這批網頁資料讀進來時 Excel 上原本的數字（self.before）。
        寫入成功之後 item["current"] 就是新數字了，沒有這份快照的話，畫面上
        再也看不到「被蓋掉之前是多少」。

        現金要 compare_web=False：它的 web 是今日淨收付、不是餘額，
        拿去跟 Excel 的餘額比是兩件不同的東西。
        """
        before = self.before.get((name, item["cell"]), item["current"])
        if item["will_write"]:
            return f"{show(before)} → {show(item['proposed'])}", False
        if not values_match(before, item["current"]):
            return f"{show(before)} → {show(item['current'])}", True
        if compare_web and item["web"] is not None and not values_match(item["current"], item["web"]):
            return f"{show(item['current'])}（網頁 {show(item['web'])}）", False
        return show(item["current"]), False

    def _cash_turned(self, item, before):
        """
        餘額這一輪是不是由正翻負。

        紅字是一個數字一個數字上的（負的才紅），所以這裡只剩「由正變負」這件事
        要另外講 —— 那是這一條上最該被看見的一次變化，而兩個各自上色的數字
        只說得出「現在是負的」，說不出「今天才變負的」。
        """
        old = to_num(before, None)
        new = to_num(item["proposed"] if item["will_write"] else item["current"], None)
        return old is not None and old >= 0 and new is not None and new < 0

    def fill_sync_tree(self):
        """左邊的名單與右邊的明細一起重畫。"""
        self._fill_people()
        self._fill_detail()

    def _summary(self, name):
        """一位交易人的濃縮狀態：(要寫幾格, 幾條提醒, 現金顯示值, 現金是不是負的)。"""
        items = self.proposals.get(name, [])
        writes = sum(1 for item in items if item["will_write"])
        warns = len(self.warnings.get(name, []))

        cash = next((item for item in items if item["kind"] == "cash"), None)
        if cash is None:
            return writes, warns, "", False
        value = cash["proposed"] if cash["will_write"] else cash["current"]
        number = to_num(value, None)
        return writes, warns, show(value), number is not None and number < 0

    def _filtered_sheets(self):
        """
        名單上該有誰。

        勾了「只看有差異的」就只留要動的那幾位，但正在看的那一位一定留著 ——
        不然勾下去的瞬間畫面會跳到別人身上，而使用者只是想少看幾個人而已。
        """
        names = list(self.proposals)
        if not self.only_diff.get():
            return names
        return [name for name in names
                if name == self.current_sheet or any(self._summary(name)[:2])]

    def _shown(self):
        """名單上現在真的有誰。上一位／下一位都照這個走，才會跟眼睛看到的一致。"""
        return list(self.people.get_children())

    def _fill_people(self):
        self.people.delete(*self.people.get_children())

        names = self._filtered_sheets()
        # 正在看的那位不見了（換了檔、重讀、改了篩選）就退回第一位。
        # 右邊不能停在一個名單上已經沒有的人身上。
        if self.current_sheet not in names:
            self.current_sheet = names[0] if names else None

        need = 0
        for name in names:
            writes, warns, cash, negative = self._summary(name)
            tags = []
            if writes:
                need += 1
                flag = f"要寫 {writes}"
                tags.append("attention")
            elif warns:
                flag = "⚠"
                tags.append("warned")
            else:
                flag = "✓"
            if negative:
                tags.append("negative")
            self.people.insert(
                "", "end", iid=name,
                text=name + ("（模擬）" if name in self.fake_sheets else ""),
                values=(cash, flag), tags=tuple(tags),
            )

        total = len(self.proposals)
        if not total:
            self.people_count.configure(text="")
        elif need:
            self.people_count.configure(text=f"{need} / {total} 位要處理")
        else:
            self.people_count.configure(text=f"共 {total} 位，都一致")

        if self.current_sheet:
            self.people.selection_set(self.current_sheet)
            self.people.see(self.current_sheet)

    def _on_person_selected(self, _event=None):
        picked = self.people.selection()
        # selection_set 自己也會觸發這個事件。選的還是同一位就別再畫一次，
        # 否則每次重建名單都要多畫一張明細。
        if not picked or picked[0] == self.current_sheet:
            return
        self.current_sheet = picked[0]
        self._sync_scope_to_person()
        self._fill_detail()

    def _step_person(self, delta):
        """換上一位／下一位。名單怎麼排就怎麼走，到頭了就停住（不繞回去）。"""
        names = self._shown()
        if not names:
            return
        index = names.index(self.current_sheet) if self.current_sheet in names else 0
        index = max(0, min(len(names) - 1, index + delta))

        # 自己更新，不靠 <<TreeviewSelect>> —— 那個事件是排進事件佇列才送的，
        # 「按鈕按了、右邊要立刻換」不該押在佇列什麼時候輪到它上面。
        # selection_set 之後那個事件還是會來，但它看到選的是同一位就直接跳過。
        self.current_sheet = names[index]
        self.people.selection_set(self.current_sheet)
        self.people.see(self.current_sheet)
        self._sync_scope_to_person()
        self._fill_detail()

    def _grouped(self, name):
        """
        把一格一列的提案併成畫面上的列：一檔股票一列，現金另外拿出來。

        以 Excel 的列號分組而不是股號 —— 畫面上的一列就是檔案上的同一列，
        對照的時候不必在心裡再翻譯一次。

        一個分頁只有第 4~8 列五個位置，一列一檔、股號不會重複。真的重複了的話
        這裡照樣會畫成兩列，但底下撐不住：紀錄檔是拿股號當 key 的
        （holdings[股號][qty]），兩列共用同一份「程式記得多少」，結果是寫了第一列、
        第二列就被判成人工改動，而網頁那一檔的總股數還會整個寫進其中一列。
        所以那不是一種支援的排法，是一種要避開的排法。
        """
        groups, order, cash = {}, [], None
        for item in self.proposals.get(name, []):
            if item["kind"] == "cash":
                cash = item
                continue
            if item["row"] not in groups:
                groups[item["row"]] = {}
                order.append(item["row"])
            groups[item["row"]][item["which"]] = item
        return [groups[row] for row in order], cash

    def _fill_detail(self):
        """右邊那張表：選中的交易人有哪幾檔、哪幾格要動。"""
        self.tree.delete(*self.tree.get_children())
        name = self.current_sheet
        groups, cash = self._grouped(name)

        # 高度照這位有幾檔算，但至少留八列 —— 每換一個人表格就跳一次高度、
        # 底下的現金跟著上下移動，比空幾列還難看，所以下限抓在「手上大概會有
        # 幾檔」而不是「這一位現在有幾檔」；上限只是防呆（真有人塞了三十檔，
        # 那就讓它捲）。
        self.tree.configure(height=max(8, min(len(groups), 18)))

        for group in groups:
            qty, cost = group.get("qty"), group.get("cost")
            both = [item for item in (qty, cost) if item is not None]
            if not both:
                continue

            texts, written = {}, False
            for which, item in (("qty", qty), ("cost", cost)):
                if item is None:
                    texts[which] = ""
                    continue
                texts[which], done = self._value_text(name, item)
                written = written or done

            tags = []
            if any(item["will_write"] for item in both):
                tags.append("write")
            elif written:
                tags.append("done")
            # 前景色一列只給一個 —— 同時掛兩個管顏色的標籤，最後誰贏是 Tk 的
            # 內部順序決定的，看起來就會時橘時灰。
            if any(item["status"] == ledger_mod.MANUAL for item in both):
                tags.append("manual")
            elif any(item["status"] in (ledger_mod.UNTRACKED, planner.WEB_MISSING)
                     for item in both):
                tags.append("untracked")

            self.tree.insert(
                "", "end",
                values=(
                    " ".join(item["cell"] for item in both),
                    stock_title(both[0]["label"]),
                    texts["qty"], texts["cost"],
                    group_status(qty, cost), group_note(qty, cost),
                ),
                tags=tuple(tags),
            )

        self._fill_head(name)
        self._fill_cash(name, cash)
        self._fill_opening(name, cash)
        self._fill_notes(name)

    def _fill_cash(self, name, item):
        """
        表格底下那一條現金：現金餘額 B8　舊值 → 新值　狀態：…　說明：…

        一段一段插，負的數字自己上紅字。沒資料就整條清掉 ——
        換到還沒讀過的人時留著上一位的餘額，是這畫面上最危險的一種殘影。
        """
        segments = []
        if item is not None:
            before = self.before.get((name, item["cell"]), item["current"])
            after = item["proposed"] if item["will_write"] else item["current"]
            turned = self._cash_turned(item, before)

            segments.append((f"現金餘額 {item['cell']}", "dim"))
            segments.append(("　", None))
            segments.append((show(before), _neg_tag(before)))
            if not values_match(before, after):
                segments.append(("　→　", None))
                segments.append((show(after), _neg_tag(after)))

            segments.append(("　　狀態：", "dim"))
            segments.append((planner.STATUS_NAMES.get(item["status"], item["status"]), None))

            # 「餘額轉負」擺在說明最前面 —— 後面那句「今日淨收付…」每天都在，
            # 由正變負卻是難得一次，排在後面會被當成例行文字滑過去。
            note = item["note"]
            if note or turned:
                segments.append(("　　說明：", "dim"))
                if turned:
                    segments.append(("餘額轉負", "turned"))
                    if note:
                        segments.append(("；", None))
                if note:
                    segments.append((note, None))

        self.cash_line.configure(state="normal")
        self.cash_line.delete("1.0", "end")
        for text, tag in segments:
            self.cash_line.insert("end", text, (tag,) if tag else ())
        self.cash_line.configure(state="disabled")
        # 高度要等版面算完寬度才知道換不換行，所以排到 idle 再量。
        self.cash_line.after_idle(self._fit_cash_line)

    def _fit_cash_line(self):
        """
        讓那一條剛好包住內容。說明偶爾會長到換行，固定一列會被切掉。

        高度沒變就什麼都不做 —— 這個函式也綁在 <Configure> 上，每改一次高度就是
        一次新的 Configure，照改不誤的話兩個值會互相觸發、來回抖個不停。
        """
        try:
            lines = self.cash_line.count("1.0", "end", "displaylines")[0]
        except (tk.TclError, TypeError, IndexError):
            lines = 1
        height = max(1, min(lines, 3))
        if height != int(self.cash_line["height"]):
            self.cash_line.configure(height=height)

    def _fill_opening(self, name, item):
        """
        現金那一條底下那一行：今日初始現金餘額，加一顆改它的按鈕。

        數字直接讀紀錄檔的現金基準（見 ledger.opening_balance），不是提案算出來的
        —— 它講的是「今天從多少錢開始」，跟這一輪要不要寫哪一格無關，就算這位
        今天一格都不必動也要看得到。基準每天由當天第一次登入設成 B8，所以正常
        情況它就是今天早上的那個數字。

        剛按過「修改」還沒落帳的時候寫成「舊 → 新」，跟上面那一條同一個寫法：
        按完卻還顯示舊數字，看起來就像沒按到。

        沒有網頁資料就不給改：新的餘額是「開盤前 + 今日淨收付」，
        沒有淨收付算不出來，也就沒有東西可以寫回 Excel。
        """
        cash = self.ledger.sheet(name)["cash"] if (self.ledger is not None and name) else None
        opening = ledger_mod.opening_balance(cash) if cash is not None else None

        # 一位都還沒選（還沒讀過網頁資料）時寫破折號而不是「還沒設定」——
        # 那時候是「不知道要看誰」，不是「這個人沒有基準」。
        text = "—" if not name else ("(還沒設定)" if opening is None else show(opening))
        if item is not None and item["reset_to"] is not None:
            text = f"{text} → {show(round(item['reset_to'] - item['net'], 2))}"

        # 平常寫的是「這個數字打哪來」。它是整條算式裡唯一可能錯的一項，
        # 而它怎麼來的決定了它什麼時候會錯 —— 講出來，比只擺一個數字有用。
        # 按鈕變灰的時候換成理由：灰掉而不說為什麼，看起來就像壞了。
        reason = ""
        if self.ledger is None or not name:
            reason = ""
        elif item is None:
            reason = "按「讀取網頁資料」之後才改得動"
        elif item["blocked"]:
            # 淨收付本身信不過的時候連基準都不該讓人按（見 planner._cash）：
            # 拿一個已知是錯的淨收付去算，等於把今天的成交永久算進基準裡。
            reason = "今日淨收付對不上，這個數字先不要動"
        else:
            # 後半句是「什麼時候該按那顆按鈕」的完整答案：那個時間點的 B8 已經
            # 含了今天的成交（自己盤中改過），才需要按。平常不必管它。
            reason = "＝ 今天第一次登入時 Excel 上的 B8（那時它已含今天的成交才要按「修改」）"

        number = to_num(opening, None)
        self.opening_value.configure(
            text=text, foreground="#c00000" if number is not None and number < 0 else "")
        self.opening_hint.configure(text=reason)
        self.opening_ready = (self.ledger is not None and item is not None
                              and not item["blocked"])
        self._sync_buttons()

    def edit_opening(self):
        """
        「修改」按下去：問一個新的開盤前現金，套進提案，然後照這一輪的規矩落實。

        套用走的是 planner.apply_cash_reset —— 跟程式自己跳的那個對話框同一段
        程式碼，兩個入口不會算出不一樣的結果。

        落實的方式跟著「程式自動更新」走，不另外開一條路：勾著就直接寫進 Excel
        並落帳（沒有格子要寫也一樣要落帳，理由見 _commit_round）；沒勾就只是把
        建議值畫出來，Excel 由人自己改 —— 人工維護的時候程式一格都不碰，
        這顆按鈕不該是例外。
        """
        name = self.current_sheet
        item = self._cash_item(name)
        if self.ledger is None or item is None or item["blocked"] or self.busy:
            return

        cash = self.ledger.sheet(name)["cash"]
        opening = ask_opening_balance(self.root, self.family, name,
                                      ledger_mod.opening_balance(cash), item)
        if opening is None:
            return

        planner.apply_cash_reset(item, opening)
        # 這顆按鈕動到的只有眼前這一位，寫入與落帳的範圍就跟著縮到他身上
        # —— 名單上別人那些「要寫」的格子是上一輪算的，不該被這一下順手寫出去。
        self.round_scope = {name}
        self.fill_sync_tree()

        if not self.auto_write.get():
            self._say(f"{name} 的今日初始現金餘額改成 {show(opening)}，"
                      f"{item['cell']} 要填 {show(item['proposed'])} —— "
                      f"「程式自動更新」沒有勾，程式不會動 Excel，請自己填進去。")
            return

        writes, total = self._collect_writes()
        if total:
            self._begin_write(writes, total)
            return

        # 算出來剛好等於 Excel 上的數字，一格都不必寫 —— 但基準確實被改掉了，
        # 這一筆不落帳就等於沒按過（見 _commit_round）。
        recorded = self._commit_round()
        self.replan()
        self.refresh_history()
        self._say(f"{name} 的今日初始現金餘額改成 {show(opening)}，"
                  f"{item['cell']} 上的數字剛好一樣，沒有格子要寫。"
                  + (f"紀錄檔更新了 {recorded} 筆（見歷程）。" if recorded else ""))

    def _fill_head(self, name):
        """表格上方那一行：是誰、第幾位、現金多少、這次要寫幾格、資料是幾點讀的。"""
        if not name:
            self.detail_head.configure(text="還沒有資料 —— 按上面的「讀取網頁資料」")
            return

        writes, _warns, cash, _negative = self._summary(name)
        names = self._shown()
        parts = [name + ("（模擬）" if name in self.fake_sheets else "")]
        if name in names:
            parts.append(f"第 {names.index(name) + 1} / {len(names)} 位")
        if cash:
            parts.append(f"現金 {cash}")
        parts.append(f"要寫 {writes} 格" if writes else "跟網頁一致")
        # 一次只更新一位之後，畫面上每個人的資料新舊不一 —— 沒有這個時間，
        # 半小時前讀的數字跟剛剛讀的長得一模一樣。
        read = self.read_at.get(name)
        if read:
            parts.append(f"讀取於 {read:%H:%M}")
        self.detail_head.configure(text="　".join(parts))

    def _fill_notes(self, name):
        """
        提醒只講選中的這一位。

        20 個人的提醒全部堆在同一個框裡等於沒有提醒 —— 別人的事左邊名單上
        已經用 ⚠ 標出來了，要看就換過去看。整組失敗（problems）例外，
        那跟選中誰無關，一定要講。
        """
        text = [f"• {warning}" for warning in self.warnings.get(name, [])]
        for problem in self.problems:
            text.append(f"⚠ {problem}")

        if self.ledger_fresh and self.proposals:
            # 只有這一天要講。紀錄檔存在的日子裡，現金起點是「今天第一次登入時的
            # B8」，而那個時間點今天還沒開盤 —— 沒有猜的成分。紀錄檔是新的（第一次
            # 用這支程式、換了電腦、Excel 改過檔名或搬過資料夾）就不一樣了：
            # 起點只能取「現在的 B8」，它含不含今天的成交，程式看不出來，
            # 唯一看得出來的人是使用者。
            text.append("這份 Excel 還沒有紀錄檔（第一次用、換電腦，或檔名／資料夾"
                        "改過），今天的「今日初始現金餘額」是拿 B8 現在的數字當起點的。"
                        "今天如果已經有成交、而 B8 已經含了它，請按現金底下的「修改」"
                        "改成今天開盤前的金額，否則今天的淨收付會被算兩次。")

        others = sum(1 for other, warns in self.warnings.items() if warns and other != name)
        if others:
            text.append(f"另外有 {others} 位交易人也有提醒，左邊名單上標著 ⚠。")

        if not self.proposals:
            text.append("還沒有資料。按上面的「讀取網頁資料」，這裡就會列出每個交易人要改哪幾格。")
        elif not self._pending_count():
            # 同一句話，主詞跟著上面那幾行 ⚠ 縮小 —— 有人沒完成的時候，
            # 「已經一致」只能替畫面上有的那幾位講。
            text.append("目前沒有任何格子需要寫入 —— "
                        + ("上面 ⚠ 之外的交易人，Excel 上的數字跟網頁已經一致。"
                           if self.problems else "Excel 上的數字跟網頁已經一致。"))
        elif not self.auto_write.get():
            text.append("「程式自動更新」沒有勾，程式不會動 Excel，箭頭右邊那個數字"
                        "只是建議值，要自己填進 Excel；要讓程式代勞就把上面那個"
                        "開關勾起來，再按一次上面那顆讀取／更新。")

        if not text:
            text.append("沒有需要注意的事。")

        # 每天都成立的兩條規矩，擺在最後一行。這個框只有五行高又沒有捲軸，
        # 常駐的字排在前面就會把當天真正的警告推出視線 —— 排最後的話，
        # 沒事的日子看得到（框是空的），有事的日子被擠掉的正好是它。
        if self.proposals:
            text.append("今天操作中不要改 Excel 的檔名或搬資料夾，也不要自己去改 B8 —— "
                        "兩者都會讓今天的起點算錯，今天的淨收付被算兩次；真的改了，"
                        "就用現金底下那顆「修改」把今天開盤前的金額補給程式。")

        self.warn_box.configure(state="normal")
        self.warn_box.delete("1.0", "end")
        self.warn_box.insert("1.0", "\n".join(text))
        self.warn_box.configure(state="disabled")

    def _pending_count(self):
        """所有交易人加起來要寫幾格。寫入是一次寫全部，所以要照全部算，不能只算眼前這位。"""
        return sum(1 for items in self.proposals.values()
                   for item in items if item["will_write"])

    def _sync_buttons(self):
        """上面那兩顆能不能按。畫面上會變灰的按鈕現在只剩它們。"""
        # Excel 沒開著就不給登入，也不給讀取 —— 讀取自己會順便登入，只擋登入的話
        # 這道關卡按另一顆按鈕就繞過去了。擋在最前面的理由是後面每一步都要 Excel：
        # 讀完要拿它的現值算提案，寫入更是直接改它。
        ready = self.excel_open and not self.busy
        self.login_button.configure(state="normal" if ready else "disabled")
        self.fetch_button.configure(state="normal" if ready else "disabled")
        # 「修改」不看 Excel 開著沒 —— 它改的是紀錄檔裡的基準，要寫 Excel 的時候
        # 寫入那邊自己會把檔案開起來。能不能按只看「這一位有沒有網頁資料」，
        # 那是 _fill_opening 判的。
        self.opening_button.configure(
            state="normal" if self.opening_ready and not self.busy else "disabled")

    def _auto_adopt(self):
        """
        自動模式下，把還沒交給程式的格子直接接管，接管完重新算一次提案。

        登入時的初始化（見 _initialize）已經把 Excel 上有的東西都收下來了，
        這裡是補網：登入之後才新增的列、登入當下沒讀到的分頁、以及使用者中途
        手改過的格子。規則跟初始化同一條 —— Excel 上的現值就是新的起點。

        現金的基準一天只設一次。今天已經設過（很可能也已經寫進 B8 了），再設一次
        會讓今天的淨收付被加第二次，所以今天設過的就跳過。
        """
        if self.ledger is None:
            return

        at = datetime.datetime.now().isoformat(timespec="seconds")
        today = self.today.isoformat()
        events = []

        for name, items in self.proposals.items():
            # 接管是「拿 Excel 現值當新起點」，只能對這一輪剛讀過的那幾位做：
            # 別人手上的現值是上一輪讀的，中間有可能已經被人改過。
            if name not in self.round_scope:
                continue
            book = self.ledger.sheet(name)
            for item in items:
                if item["kind"] == "cash" and book["cash"].get("baseline_date") == today:
                    continue
                # 狀態不是「未接管／手動」的話 adopt_one 自己會跳過，這裡不必再判一次。
                _message, event, _needs = planner.adopt_one(
                    item, book, name, self.today, False, at)
                if event:
                    events.append(event)

        if not events:
            return

        self.ledger.save()
        self.ledger.append_history(events)
        self.replan()
        self.refresh_history()
        self._say(f"已接管 {len(events)} 格。")

    # ---------- 歷程分頁 ----------

    def refresh_history(self):
        """重讀歷程檔。篩選是在記憶體裡做的，換選單不會再碰一次硬碟。"""
        path = self.ledger.history_path if self.ledger else None
        self.history_file = path.name if path else ""
        self.history_rows = []

        if path and path.is_file():
            for line in path.read_text(encoding="utf-8").splitlines():
                line = line.strip()
                if not line:
                    continue
                try:
                    self.history_rows.append(json.loads(line))
                except json.JSONDecodeError:
                    continue

        self._refresh_history_choices()
        self._fill_history()

    def _refresh_history_choices(self):
        """
        照現有的歷程重建兩個選單。

        項目只列出「選中的那個交易人真的有的」—— 把 20 個帳號上百格全部攤平在
        同一個下拉選單裡，跟沒有篩選是一樣的。原本選的那一項還在就留著，
        不在了就退回「全部」：選單上的字還在、表格卻是空的，是最難懂的那種畫面。
        """
        who = self.history_who.get() or ALL_CHOICE
        names = sorted({row.get("sheet", "") for row in self.history_rows if row.get("sheet")})
        self.history_who.configure(values=[ALL_CHOICE] + names)
        if who not in names:
            who = ALL_CHOICE
        self.history_who.set(who)

        item = self.history_item.get() or ALL_CHOICE
        labels = sorted({row.get("label", "") for row in self.history_rows
                         if row.get("label") and (who == ALL_CHOICE or row.get("sheet") == who)},
                        key=item_order)
        self.history_item.configure(values=[ALL_CHOICE] + labels)
        if item not in labels:
            item = ALL_CHOICE
        self.history_item.set(item)

    def _on_history_who(self, _event):
        # 換人之後項目清單要跟著換，不然會停在一個那個人根本沒有的項目上。
        self._refresh_history_choices()
        self._fill_history()

    def _fill_history(self):
        """把通過篩選的事件畫進表格。"""
        self.history_tree.delete(*self.history_tree.get_children())
        self._sync_clear_button()

        if not self.history_rows:
            self.history_hint.configure(text="還沒有任何歷程。第一次寫入或交接之後就會有。")
            return

        who, item, when = (self.history_who.get(), self.history_item.get(),
                           self.history_when.get())
        # 今天是現算的，不是開程式那一刻的 self.today —— 這支程式常常開著過夜，
        # 跨過午夜之後「今天」還停在昨天的話，剛跑完的那一批會整批不見。
        today = datetime.date.today()
        shown = [row for row in self.history_rows
                 if (who == ALL_CHOICE or row.get("sheet") == who)
                 and (item == ALL_CHOICE or row.get("label") == item)
                 and within(row.get("at"), when, today)]

        for event in reversed(shown):         # 最新的放最上面
            by = event.get("by", "")
            self.history_tree.insert(
                "", "end",
                values=(
                    (event.get("at") or "").replace("T", " "),
                    event.get("sheet", ""), event.get("cell", ""), event.get("label", ""),
                    describe_change(by, event.get("old"), event.get("new")),
                    SOURCE_NAMES.get(by, by),
                    event.get("note", ""),
                ),
                tags=(by,),
            )

        counted = (f"{len(shown)} 筆" if len(shown) == len(self.history_rows)
                   else f"篩出 {len(shown)} 筆／共 {len(self.history_rows)} 筆")
        self.history_hint.configure(text=f"{counted}，最新的在最上面。檔案：{self.history_file}")

    def _sync_clear_button(self):
        """沒有歷程可清、或正在跑的時候，「清除歷程」不要亮著。"""
        usable = bool(self.history_rows) and not self.busy
        self.history_clear.configure(state="normal" if usable else "disabled")

    def clear_history(self):
        """
        把歷程收進「備份」資料夾，畫面清空。

        清掉的只有「誰在什麼時候改了哪一格」這本日記，不會動到紀錄檔 ——
        每一格歸誰管、現金的基準與流水都在那邊，所以清完再同步，
        算出來的提案跟清之前一模一樣。這句話也要講給使用者聽：
        「清除」兩個字很容易被讀成「把帳歸零」，那是最貴的誤會。
        """
        if not self.ledger or not self.history_rows:
            return

        if not messagebox.askyesno(
                "清除歷程",
                f"要清掉全部 {len(self.history_rows)} 筆歷程嗎？\n\n"
                "舊的歷程檔會改名收進「備份」資料夾，不是真的刪掉。\n"
                "每一格歸誰管、現金的基準都記在另一個紀錄檔裡，不受影響。",
                icon="warning", default="no", parent=self.root):
            return

        try:
            saved = self.ledger.clear_history()
        except OSError as exc:
            messagebox.showerror("清不掉", f"歷程檔可能正被別的程式開著：\n{exc}")
            return

        self.refresh_history()
        # 收去哪裡要寫在狀態列上，不然「備份」兩個字等於白講 —— 人找不到的備份
        # 跟沒有備份是一樣的。
        self._say(f"歷程已清空。舊的收在 {saved}" if saved else "歷程已清空。")

    def _on_tab_changed(self, _event):
        index = self.tabs.index(self.tabs.select())
        if index == 1:
            self.refresh_history()
            # 切過來通常就是想看「剛才那位」的歷程，不必再選一次人。
            if self.current_sheet and self.current_sheet in tuple(self.history_who["values"]):
                self.history_who.set(self.current_sheet)
                self._refresh_history_choices()
                self._fill_history()
        elif index == 2:
            self._refresh_profile_status()
            # 掃描要讀遍 Chrome／Edge 每個 profile 的 Local Storage，不便宜，
            # 所以只在第一次切過去時自動做一次，之後靠手動按「掃描」。
            if not self.cert_tab_scanned:
                self.cert_tab_scanned = True
                self.scan_cert_sources()


def within(stamp, when, today):
    """這一筆的時間在不在選的期間裡。"""
    if when == ALL_CHOICE:
        return True
    try:
        day = datetime.date.fromisoformat((stamp or "")[:10])
    except ValueError:
        # 時間壞掉的那幾筆只有「全部」看得到。當成今天會更糟 ——
        # 那等於憑空生出一筆今天的異動，而歷程是拿來對帳的東西。
        return False
    if when == WHEN_TODAY:
        return day == today
    return (today - day).days < 7      # 含今天往回數七天


def _neg_tag(value):
    """負的數字才上紅字。空的、不是數字的都不算。"""
    number = to_num(value, None)
    return "neg" if number is not None and number < 0 else None


def stock_title(label):
    """從「股數（2059 川湖）」取出「2059 川湖」。取不到就原樣顯示。"""
    inside = label.partition("（")[2].rstrip("）")
    return inside or label


def group_status(qty, cost):
    """
    一列的狀態。兩格同一種狀態就寫一次，不同才拆開講。

    同一檔股票的股數還在自動、成本被手改成手動是可能的，這種時候寫「自動」
    或「手動」都是錯的 —— 而這一列到底哪一格不會被覆蓋，正是使用者要判斷的事。
    """
    names = {which: planner.STATUS_NAMES.get(item["status"], item["status"])
             for which, item in (("股數", qty), ("成本", cost)) if item is not None}
    if len(set(names.values())) <= 1:
        return next(iter(names.values()), "")
    return "／".join(f"{which}{name}" for which, name in names.items())


def group_note(qty, cost):
    """
    一列的說明。兩格講的是同一件事就寫一次，否則各自標明是哪一格。

    只有一格有話說的時候也要標 —— 「手動改過，不會覆蓋」沒說是股數還是成本，
    等於沒說。
    """
    notes = [(which, item["note"])
             for which, item in (("股數", qty), ("成本", cost))
             if item is not None and item["note"]]
    if not notes:
        return ""
    present = len([item for item in (qty, cost) if item is not None])
    if len(notes) == present and len({note for _which, note in notes}) == 1:
        return notes[0][1]
    return "；".join(f"{which}：{note}" for which, note in notes)


def item_order(label):
    """
    項目選單的排序：現金排最前面，其餘照股票分組，同一檔的股數排在成本前面。

    純照字串排會變成「所有股數」一段、「所有成本」另一段，同一檔股票的兩格
    隔了幾十列 —— 但人是先想到哪一檔股票，才想到要看股數還是成本。
    """
    if label.startswith("現金"):
        return (0, "", 0)
    inside = label.partition("（")[2].rstrip("）")
    return (1, inside, 0 if label.startswith("股數") else 1)


def describe_change(by, old, new):
    """
    歷程那一欄要寫什麼。

    「交接」的新舊值講的是**程式的記憶**，不是格子的值：交接就是把
    「我記得這格是多少」改成 Excel 上現在的數字，Excel 本身一格都沒動。
    如果照 program 那樣印成「1 → 0」，看起來會像程式把數字改壞了，
    所以新值那邊寫「改記」，點明變的是記憶。
    """
    if by == "adopt":
        return f"{show(old)} → 改記 {show(new)}"
    if by == "human":
        return f"{show(old)} → {show(new)}（人工）"
    return f"{show(old)} → {show(new)}"


def main():
    root = tk.Tk()
    app = SyncApp(root)

    def on_close():
        if app.busy and not messagebox.askokcancel(
            "還在忙", "背景還在登入或寫入。現在關掉可能會留下寫到一半的狀態，確定要關嗎？"
        ):
            return
        if app.browser_thread is not None and app.browser_thread.is_alive():
            app.browser_cmd_queue.put(("stop", None))
            app.browser_thread.join(timeout=10)
        root.destroy()

    root.protocol("WM_DELETE_WINDOW", on_close)
    root.mainloop()


if __name__ == "__main__":
    main()
