"""
ui.py 底下幾個分頁共用的字級／視窗尺寸換算、表格欄寬計算、對話框。

獨立成這個模組是為了讓 ui_layout / ui_cert / ui_background / ui_sync / ui_history
可以互相不認識彼此，只認這一份共用底層 —— 不然任兩個分頁模組只要有一個要用到
對方模組裡的東西，就會兜出循環匯入。
"""

import tkinter as tk
from tkinter import font as tkfont
from tkinter import messagebox

import ttkbootstrap as ttk

import planner
from util import env_int, show, to_num

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
    「說明」，這一格程式會不會覆蓋、為什麼不覆蓋，答案就在那兩欄裡。字級調大或把中間
    那條分隔線往右拖，切掉的就更多。
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
        f"「{name}」今日開盤前的現金餘額"
    )).grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 10))

    lines = (("今日初始現金餘額", show(current)),
             (f"今日淨收付（{item['net_rows']} 筆成交）", show(item["net"])))
    for index, (title, value) in enumerate(lines, start=1):
        ttk.Label(outer, text=title, style="Hint.TLabel").grid(row=index, column=0,
                                                              sticky="w", pady=2)
        ttk.Label(outer, text=value).grid(row=index, column=1, sticky="e",
                                          padx=(24, 0), pady=2)

    text = tk.StringVar(value="" if current is None else show(current))
    row = len(lines) + 1

    ttk.Label(outer, text=f"Excel現金餘額 {item['cell']} 改成").grid(row=row, column=0,
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

    buttons = ttk.Frame(outer)
    buttons.grid(row=row + 3, column=0, columnspan=2, sticky="e", pady=(12, 0))
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


def ask_confirm(parent, title, message, *, confirm_text="是", cancel_text="否", danger=True):
    """
    是非確認對話框，取代 messagebox.askyesno。

    原生 messagebox 跳出來的是 Windows 系統對話框，字級、配色都不跟著介面走——
    使用者剛看完 tkbootstrap 畫面，下一秒切成灰底系統對話框，容易分心誤判，
    這種要人謹慎看清楚才按下去的場合反而幫倒忙。

    danger=True（預設）對應原本 icon="warning", default="no" 那幾顆：焦點鎖在
    「否」、Enter 也觸發取消，確定鍵漆成警示色——這幾顆通常管的是「立刻生效、
    蓋掉既有資料」，手滑按下 Enter 不該真的動到東西。

    danger=False 對應原本沒設 default 的那幾顆（例如複製憑證，目標會先備份，
    風險比較低）：焦點在確定鍵、Enter 直接確認、按鈕用一般強調色，行為跟
    messagebox.askyesno 預設的「Enter＝是」一致。
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
                             bootstyle="warning" if danger else "primary")
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
    今天的現金餘額要用哪一種算法。回傳選定的算法，取消就回 None（沿用原本的）。

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

    for index, key in enumerate((planner.METHOD_BANK, planner.METHOD_OPENING)):
        ttk.Radiobutton(outer, text=planner.METHOD_NAMES[key], value=key,
                        variable=choice, style="Choice.TRadiobutton").grid(
            row=2 + index, column=0, sticky="w", pady=4)

    # 兩行公式只給人直接對照數字，不重複解釋每一項是什麼（那一段人已經在上面
    # 選過一次了）。「全額交割要選哪一種」單獨一行、用跟主畫面「現金算法」
    # 同一種藍字（Method.TLabel），因為這是唯一一句填錯會被吃掉、隔天才看得出來的話。
    ttk.Label(outer, justify="left", style="Hint.TLabel", text=(
        f"{planner.METHOD_NAMES[planner.METHOD_BANK]} ＝ 銀行餘額(網頁) ＋ 當日淨收付(網頁) ＋ 昨日淨收付(網頁)\n"
        f"{planner.METHOD_NAMES[planner.METHOD_OPENING]} ＝ 現金餘額(Excel) ＋ 當日淨收付(網頁)"
    )).grid(row=4, column=0, sticky="w", pady=(10, 0))

    ttk.Label(outer, justify="left", style="Method.TLabel", text=(
        f"全額交割當天要選「{planner.METHOD_NAMES[planner.METHOD_OPENING]}」"
        f"（{planner.METHOD_NAMES[planner.METHOD_BANK]}那天會扣兩次）"
    )).grid(row=5, column=0, sticky="w", pady=(10, 0))

    def confirm(*_args):
        answer["method"] = choice.get()
        win.destroy()

    buttons = ttk.Frame(outer)
    buttons.grid(row=6, column=0, sticky="e", pady=(12, 0))
    ttk.Button(buttons, text="取消", command=win.destroy,
              bootstyle="secondary").pack(side="left", padx=(0, 8))
    ttk.Button(buttons, text="確定", command=confirm,
              bootstyle="primary").pack(side="left")

    win.protocol("WM_DELETE_WINDOW", win.destroy)
    win.bind("<Escape>", lambda _e: win.destroy())
    win.bind("<Return>", confirm)
    center_on(win, parent)
    win.grab_set()
    parent.wait_window(win)
    return answer.get("method")
