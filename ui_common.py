"""
ui.py 底下幾個分頁共用的字級／視窗尺寸換算、表格欄寬計算、對話框。

獨立成這個模組是為了讓 ui_layout / ui_cert / ui_background / ui_sync / ui_history
可以互相不認識彼此，只認這一份共用底層 —— 不然任兩個分頁模組只要有一個要用到
對方模組裡的東西，就會兜出循環匯入。
"""

import datetime
import tkinter as tk
from tkinter import font as tkfont
from tkinter import messagebox

import ttkbootstrap as ttk

import planner
from util import env_int, show, to_num

# 版號：YYYYMMDD + 字母，格式跟著日期走，同一天出好幾版就往後接字母
# （20260826A -> 20260826B -> ...），換一天就換日期、字母歸零重新從 A 開始。
# 純手動維護：每次要出新版本（改完程式、準備包 exe）就手動改這一行，
# 不做成自動生成——版號代表的是「這次是不是真的發出去的那一版」，
# 不是「這次改了程式」，兩者不一定同步（改完可能還沒包、沒發）。
APP_VERSION = "20260826A"

# 三個歷程篩選選單的「不篩選」那一項。用一段文字當值而不是 None，
# 選單裡本來就會有一列寫著它，兩邊用同一個東西才不會對不起來。
ALL_CHOICE = "全部"

# 期間只給這兩個級距。歷程是拿來回答「剛才那次執行做了什麼」「這幾天有沒有人動過」，
# 再細的區間（挑日期、選月份）沒有人真的會用，卻要多一個日曆元件跟一整套錯誤處理。
WHEN_TODAY, WHEN_WEEK = "今天", "近 7 天"


def stock_title(label):
    """
    從「股數（2059 川湖）」取出「2059 川湖」。取不到就原樣顯示。

    同步分頁訊息框、歷程分頁都要把同一輪股數／成本兩筆事件併成一行，靠的是
    同一個股票名稱，兩邊要用同一支函式取名字才不會對不起來。
    """
    inside = label.partition("（")[2].rstrip("）")
    return inside or label


def within(stamp, when, today):
    """
    這一筆的時間在不在選的期間裡。

    「歷程」分頁跟同步分頁的訊息框（見 ui_sync._fill_notes）都靠這個篩今天
    ／近 7 天／全部，兩邊要是同一套判斷，篩出來的才會是同一份真相。
    """
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

# 外觀設定，全部從 .env 讀，改完重開介面生效。

FONT_ENV_KEY = "UI_FONT_SIZE"
FONT_DEFAULT = 12
FONT_MIN, FONT_MAX = 8, 24

WIDTH_ENV_KEY, HEIGHT_ENV_KEY = "UI_WIDTH", "UI_HEIGHT"
# 視窗再小就不叫小視窗，而是壞掉的版面：左邊名單擠成一條、右邊那幾格全是省略號。
WINDOW_MIN_W, WINDOW_MIN_H = 640, 400

CASH_METHOD_TOGGLE_ENV_KEY = "CASH_METHOD_TOGGLE"


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


def cash_method_toggle_enabled():
    """
    工具列上「現金算法」那個名字點一下能不能換算法。沒設、或設 1 就是開著
    （原本的行為）；設 0 就關掉 —— 名字還在，但不綁點擊、游標也不會變手指，
    看起來就是純顯示。2026/08/20 加點名字切換是為了測試方便，測得差不多之後
    使用者要求能整個關掉，正式使用時避免手滑誤觸；要換算法就回到 .env
    改設定重開程式，或關掉這個開關讓每天讀取前的那個對話框繼續問。
    """
    return env_int(CASH_METHOD_TOGGLE_ENV_KEY, 1) != 0


FONT_SIZE = font_size_from_env()   # 全部文字統一用這個大小，標題／表頭另外加粗做區分，不再另外縮小字級
HINT_SIZE = FONT_SIZE              # 提示/次要文字：路徑、提醒、欄位說明——跟 FONT_SIZE 同大小，只用顏色/粗細區分，不再縮小
WINDOW_W, WINDOW_H = window_size_from_env()


def wide(pixels):
    """
    把「照 10 級字調出來的像素寬」換算成目前字級要的寬度。

    欄寬跟列高不會自己跟著字級長大。放大字卻沒放大欄寬的結果是
    「成本（2059 川湖）」被截成「成本（2059…」—— 使用者只會看到一個莫名其妙
    的省略號，不會知道那是版面問題。
    """
    return int(pixels * FONT_SIZE / 10)


# 表格每一格左右各留的空白，套在 ui_layout 的 Treeview.Cell 樣式上
# （style.configure("Treeview.Cell", padding=(CELL_PAD, 0))）——2026/08/21 使用者
# 反應「數字貼太緊」，預設只留 4 像素，靠右的數字幾乎貼著欄位邊界。
CELL_PAD = 8


# 下單分頁「執行預覽」表格：盤中模式最終委託價還沒算出來時顯示的文字，不是
# 空白也不是猜一個數字（見 orders.plan_intraday_orders）——成交價基準已經
# 是 Excel 讀到的值，只差下單前那一刻查對手方第一檔比價（見 orders.chase_price）。
# 定義在這裡而不是 ui_order.py，是因為 ui_layout.py 算「價格」欄要多寬時也要
# 量到同一句話，兩邊都認 ui_common 不必互相 import（見檔案開頭的說明）。
PRICE_PENDING_TEXT = "依 Excel 成交價追價"


def col_width(family, texts, minimum=0):
    """
    量一批候選字串（Treeview 目前的字型：family + FONT_SIZE）裡最寬的一個，
    算出這一欄該給多寬，取代原本用固定數字硬猜寬度。

    只適合用在「候選字串是有限、可列舉的」欄位（股票名稱清單、帳戶名單、
    orders.py 那幾句固定備註）——不能拿使用者正在打字的即時輸入內容當
    texts，那樣每個按鍵都會讓欄寬跟著跳動一次，反而更難用（2026/08/29
    使用者討論下單分頁「執行預覽」欄寬時確認過這個界線）。

    量出來的已經是 FONT_SIZE 級字型的實際像素，不必再套 wide()——wide()
    只用在沒有實際文字可量、純粹用數字猜的寬度（見 wide() 說明）。
    minimum 是那種猜出來的下限，兩者取大的，量出來的字串再短也不會比
    原本設計的下限還窄。
    """
    pad = CELL_PAD * 2 + 10
    if not texts:
        return minimum
    font = tkfont.Font(family=family, size=FONT_SIZE)
    return max(max(font.measure(t) for t in texts) + pad, minimum)


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


def ask_opening_balance(parent, family, name, current):
    """
    改一個人的「今日初始現金餘額」。回傳新的基準，取消或留空就回 None。

    這裡曾經有另一個對話框：程式自己判斷「今天可能已經開過了」，讀完網頁資料就
    跳出來、一次問完所有分頁。它被這顆按鈕取代了 —— 同一個問題，程式問是「猜
    你需要」，而猜錯的兩個方向都很貴：不該跳的時候跳，20 個分頁按到最後沒人在看
    內容；該跳的時候沒跳，或跳的時候填錯，當天就再也沒有入口。基準現在一直寫在
    畫面上，看到不對再按，不必由程式決定什麼時候該問誰。

    2026/08/24 拿掉了「先心算加上今日淨收付」那一步：這裡填的就是「今日初始
    現金餘額」本身，不必先想今天成交了多少、答案是不是已經含了今天的淨收付。
    net 還是會在套用時自動加上去、寫進 Excel（見 planner.apply_cash_reset），
    只是不必在這個對話框先預覽一次會變成什麼——按下去之後那一格會寫什麼，
    畫面跟訊息框上看得到。
    """
    win = tk.Toplevel(parent)
    win.title("修改今日初始現金餘額")
    win.transient(parent)
    win.resizable(False, False)

    answer = {}

    outer = ttk.Frame(win, padding=12)
    outer.pack(fill="both", expand=True)

    ttk.Label(outer, justify="left", text=(
        f"「{name}」今日初始現金餘額"
    )).grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 10))

    ttk.Label(outer, text="目前", style="Hint.TLabel").grid(row=1, column=0,
                                                          sticky="w", pady=2)
    ttk.Label(outer, text=show(current)).grid(row=1, column=1, sticky="e",
                                              padx=(24, 0), pady=2)

    text = tk.StringVar(value="" if current is None else show(current))

    ttk.Label(outer, text="改成").grid(row=2, column=0, sticky="w", pady=(10, 2))
    entry = ttk.Entry(outer, width=16, font=(family, FONT_SIZE), justify="right",
                      textvariable=text)
    entry.grid(row=2, column=1, sticky="e", padx=(24, 0), pady=(10, 2))

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

    buttons = ttk.Frame(outer)
    buttons.grid(row=3, column=0, columnspan=2, sticky="e", pady=(12, 0))
    ttk.Button(buttons, text="取消", command=win.destroy,
              bootstyle="secondary").pack(side="left", padx=(0, 8))
    ttk.Button(buttons, text="確定", command=confirm,
              bootstyle="primary").pack(side="left")

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


def ask_confirm(parent, title, message, *, confirm_text="是", cancel_text="否", danger=True,
                confirm_style=None):
    """
    是非確認對話框，取代 messagebox.askyesno。

    原生 messagebox 跳出來的是 Windows 系統對話框，字級、配色都不跟著介面走——
    使用者剛看完 tkbootstrap 畫面，下一秒切成灰底系統對話框，容易分心誤判，
    這種要人謹慎看清楚才按下去的場合反而幫倒忙。

    danger=True（預設）對應原本 icon="warning", default="no" 那幾顆：焦點鎖在
    「否」、Enter 也觸發取消，確定鍵預設漆成警示色——這幾顆通常管的是「立刻生效、
    蓋掉既有資料」，手滑按下 Enter 不該真的動到東西。

    danger=False 對應原本沒設 default 的那幾顆（例如複製憑證，目標會先備份，
    風險比較低）：焦點在確定鍵、Enter 直接確認、按鈕用一般強調色，行為跟
    messagebox.askyesno 預設的「Enter＝是」一致。

    confirm_style 只換確定鍵的顏色，不動 danger 管的「焦點鎖在哪、Enter 觸發
    哪一顆」——顏色是「看起來像哪一種按鈕」，danger 管的是「按錯了會不會直接
    生效」，兩件事不一定要綁在一起（例如切換現金算法：換錯了有代價，所以
    danger=True 焦點還是鎖「否」，但按鈕顏色 2026/08/22 使用者要求跟同一天
    新增的「今天的現金餘額怎麼算」那顆藍色「確定」統一，不繼續用警示色）。
    """
    win = tk.Toplevel(parent)
    win.title(title)
    win.transient(parent)
    win.resizable(False, False)

    answer = {"ok": False}

    outer = ttk.Frame(win, padding=16)
    outer.pack(fill="both", expand=True)

    ttk.Label(outer, justify="left", text=message).pack(anchor="w")

    def confirm(*_args):
        answer["ok"] = True
        win.destroy()

    def cancel(*_args):
        win.destroy()

    buttons = ttk.Frame(outer)
    buttons.pack(fill="x", pady=(16, 0))
    cancel_btn = ttk.Button(buttons, text=cancel_text, command=cancel, bootstyle="secondary")
    cancel_btn.pack(side="right")
    confirm_btn = ttk.Button(buttons, text=confirm_text, command=confirm,
                             bootstyle=confirm_style or ("warning" if danger else "primary"))
    confirm_btn.pack(side="right", padx=(0, 8))

    default_btn = cancel_btn if danger else confirm_btn
    win.protocol("WM_DELETE_WINDOW", cancel)
    win.bind("<Escape>", cancel)
    win.bind("<Return>", lambda _e: default_btn.invoke())
    center_on(win, parent)
    win.grab_set()
    default_btn.focus_set()
    parent.wait_window(win)
    return answer["ok"]


def ask_cash_method(parent, family, current):
    """
    今天的現金餘額要用哪一種算法。回傳選定的算法；取消（按「取消」、X 或 Escape）
    回傳 None，呼叫端要當成「這次不讀取了」處理，不能套用任何一種算法。

    2026/08/22 到 2026/08/26 之間這裡是沒有取消的：X 跟 Escape 都關不掉，理由是
    「兩個都得選一個，沒有先不選、之後再說的空間」。2026/08/26 使用者不小心把
    「登入」按成第一次讀取前才會出現的「登入+讀取」（見 ui_background 那顆按鈕的
    labeling），這個視窗跳出來又走不掉，最後只能把整支程式關掉重開。問題不在
    「選錯了會不會被吃掉」——會走到這裡一定是還沒選過，answer 是 None 就什麼都
    不會套用——而是「跳出來的時機不是使用者要的」時，原本設計完全沒給退路。
    所以恢復取消：取消不會偷偷套用任何算法，只是讓呼叫端把這次的「讀取」整個
    收回去（見 ui_background._maybe_ask_cash_method），不影響「選好了會不會被
    誤蓋掉」這件事。

    只問，不算給你看。曾經在這裡列出「哪幾位的兩種算法對不上、差多少」，後來拿掉了：
    算法是全部人共用的一個開關，逐人列出來看起來像可以一位一位挑；而且差額講不出
    是什麼造成的（全額交割？匯撥？有人手改過 B8？），對這個決定幫不上忙。
    真正的依據是「今天有沒有買全額交割股」，那只有人知道，不在畫面上。

    因為不必等資料，這個視窗改成在**讀取之前**問 —— 選好了才去抓，
    才不會發生「選了銀行餘額推算，但這一輪沒抓銀行餘額，要再讀一次」。
    """
    win = tk.Toplevel(parent)
    win.title("今天的現金餘額怎麼算")
    win.transient(parent)
    win.resizable(False, False)

    answer = {}
    choice = tk.StringVar(value=current)

    outer = ttk.Frame(win, padding=12)
    outer.pack(fill="both", expand=True)

    ttk.Label(outer, justify="left", text=(
        "今天的現金餘額要用哪一種算法？"
    )).grid(row=0, column=0, sticky="w", pady=(0, 8))

    # 初始餘額累加排在前面、也是預設選項：它是「想不出今天算哪一種」時比較安全的
    # 那一個 —— 全額交割當天選錯（用了銀行餘額推算）會扣兩次，隔天才看得出來。
    for index, key in enumerate((planner.METHOD_OPENING, planner.METHOD_BANK)):
        ttk.Radiobutton(outer, text=planner.METHOD_NAMES[key], value=key,
                        variable=choice, style="Choice.TRadiobutton").grid(
            row=2 + index, column=0, sticky="w", pady=4)

    # 兩行公式只給人直接對照數字，不重複解釋每一項是什麼（那一段人已經在上面
    # 選過一次了）。「全額交割要選哪一種」單獨一行、用跟主畫面「現金算法」
    # 同一種藍字（Method.TLabel），因為這是唯一一句填錯會被吃掉、隔天才看得出來的話。
    ttk.Label(outer, justify="left", style="Hint.TLabel", text=(
        f"{planner.METHOD_NAMES[planner.METHOD_OPENING]} ＝ 今日初始現金餘額 + 今日淨收付\n"
        f"{planner.METHOD_NAMES[planner.METHOD_BANK]} ＝ 銀行餘額 + 淨收付(T+0) + 淨收付(T+1)"
    )).grid(row=4, column=0, sticky="w", pady=(10, 0))

    ttk.Label(outer, justify="left", style="Method.TLabel", text=(
        f"全額交割當天要選「{planner.METHOD_NAMES[planner.METHOD_OPENING]}」"
    )).grid(row=5, column=0, sticky="w", pady=(10, 0))

    def confirm(*_args):
        answer["method"] = choice.get()
        win.destroy()

    def cancel(*_args):
        win.destroy()

    buttons = ttk.Frame(outer)
    buttons.grid(row=6, column=0, sticky="e", pady=(12, 0))
    ttk.Button(buttons, text="取消", command=cancel,
              bootstyle="secondary").pack(side="left", padx=(0, 8))
    ttk.Button(buttons, text="確定", command=confirm,
              bootstyle="primary").pack(side="left")

    win.protocol("WM_DELETE_WINDOW", cancel)
    win.bind("<Escape>", cancel)
    win.bind("<Return>", confirm)
    center_on(win, parent)
    win.grab_set()
    parent.wait_window(win)
    return answer.get("method")
