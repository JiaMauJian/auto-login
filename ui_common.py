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

# 斑馬紋：每隔一列鋪一層很淡的灰，讓眼睛橫著讀不會串行。
#
# 這是「畫格線」的替代做法（2026/08/21）。格線試過用 image element 畫出來，
# 畫面對但效能很差 —— 每一列每一格都多一個元素要合成，捲動明顯變鈍
# （見 docs/Tkinter ui設計原則.md 第十二節）。斑馬紋只是換底色，
# 一列一個 tag，重畫成本跟沒有它一樣，而它解的是同一個問題。
# 做法與 ttkbootstrap Tableview 的 stripecolor 相同，只是我們用的是 ttk.Treeview，
# 自己掛 tag（Tableview 是另一個元件，帶搜尋列與分頁，換過去等於整組重寫）。
STRIPE_COLOR = "#f2f2f2"


def stripe(index, has_background=False):
    """
    第 index 列要不要加斑馬紋。回傳可以直接接在 tags 後面的 tuple。

    已經有底色的列不加：那些底色在講事情（綠＝這次要寫、藍＝剛寫過、
    名單上的綠＝這位要處理），蓋掉就沒了。一列同時掛兩個管底色的 tag，
    最後誰贏是 Tk 的內部順序決定的，看起來會時有時無。
    """
    return ("stripe",) if index % 2 and not has_background else ()

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


# 表格每一格左右各留的空白，跟 ui_layout 設 Treeview.Cell padding 用的是同一個數字。
# 算欄寬時一定要扣掉它：ttk 的 padding 吃在格子「裡面」，欄寬扣掉 8+8 才是字能用的
# 寬度。兩邊各寫各的下場是欄寬看起來夠、字卻少一位 ——2026/08/21 把留白從 4 加到 8
# 之後，成本欄的「104.6 → 10,400」就被切成「104.6 → 10,40」，而 ttk 切字不補省略號、
# 也沒有橫向捲軸，字就是斷在那裡。
CELL_PAD = 8

# padding 之外，ttk 每一格左右還會自己再吃掉幾個像素（Treeitem.text 元素的邊，
# 樣式調不掉）。抓圖數像素量出來的：padding 8 的時候，「104.6 → 10,400」量出來
# 113px，欄寬要 137px 字才畫得完整 —— 113 + 8 + 8 = 129 還差 8，也就是左右各 4。
# 少算這 8px 的下場跟少算 padding 一樣：欄寬看起來剛好，最後一個字被切掉半個。
CELL_EDGE = 4

_FONTS = {}


def _font(tree, spec):
    """
    拿一個可以量字寬的 Font。spec 是樣式裡查出來的字型（例如 "{Microsoft JhengHei UI} 12"）。

    量的必須是表格真正在用的那一份字型，所以從樣式查、不自己拼 —— 字級是 .env
    調得動的，兩邊分開算就會量出跟畫面不一樣的寬度。同一個字型只建一次：
    每次 tkfont.Font(...) 都是在 Tcl 那邊多一個字型物件，填一次表建兩個就是慢慢漏。
    """
    # 查不到就退回這支程式自己挑的字型（樣式沒設過的表格；正常路徑不會走到）。
    key = spec or "default"
    if key not in _FONTS:
        _FONTS[key] = tkfont.Font(root=tree, font=spec or (pick_font(), FONT_SIZE))
    return _FONTS[key]


def build_columns(tree, spec):
    """
    照 spec 把一張表的欄位設好，並讓寬度跟著表格伸縮。

    spec 是 (欄名, 標題, 比重, 下限, 對齊) 的序列，比重與下限都照 10 級字的像素給。
    spec 存在 tree 上，之後 fit_to_content 只要拿到表格本身就重算得出來，
    填表的地方不必再把欄位定義搬一份過去（ui_sync 也就不必認得 ui_layout）。
    """
    tree.column_spec = spec
    for key, title, _weight, floor, anchor in spec:
        tree.heading(key, text=title)
        tree.column(key, width=wide(floor), minwidth=wide(floor), anchor=anchor, stretch=False)
    tree.bind("<Configure>", lambda _event: fit_columns(tree, spec))


def fit_to_content(tree):
    """
    照表格現在的內容算出每一欄「不切字」要多寬，記在 tree 上，然後重攤一次。

    ttk 的 Treeview 沒有「照內容自動調欄寬」這種東西（width/minwidth/stretch 三個
    選項都跟內容無關，stretch 只管「變寬時要不要分到多的空間」），所以自己量。

    填完資料的地方要叫一次 —— 內容變了寬度才跟著變。刻意不放進 <Configure>：
    量一個字串是一次 Tcl 呼叫（實測 0.2ms），而拖分隔線時 Configure 一秒噴幾十次，
    放進去等於每次重畫都重量整張表。同步分頁這兩張表最多 5 檔 × 3 欄，
    一次約 3ms，而且只在按一下之後發生，感覺不出來；歷程表就不是這個量級
    （「全部」是幾千列 × 6 欄，量下去要好幾秒），所以那張表沒有叫這支，
    維持照固定比重攤。真要給它用的話得先加一層「只量最長的幾個」。

    表頭也一起量：欄位再窄也不該窄到看不出這一欄叫什麼。
    """
    spec = getattr(tree, "column_spec", None)
    if spec is None:
        return

    style = ttk.Style()
    body = _font(tree, style.lookup("Treeview", "font"))
    head = _font(tree, style.lookup("Treeview.Heading", "font"))
    rows = [tree.item(iid, "values") for iid in tree.get_children()]

    tree.content_widths = [
        max([head.measure(title)]
            + [body.measure(str(row[index])) for row in rows if index < len(row)])
        + 2 * (CELL_PAD + CELL_EDGE)
        for index, (_key, title, _weight, _floor, _anchor) in enumerate(spec)
    ]
    fit_columns(tree, spec)


def fit_columns(tree, spec):
    """
    把欄位寬度攤進表格當下的寬度：夠寬就照比重分多的，不夠寬就往下限方向縮。

    每一欄有兩個數字：
        下限  spec 裡寫死的「還讀得出來」寬度，擠到極限時的底線
        理想  裝得下這一欄現在所有內容的寬度，由 fit_to_content 量出來記在 tree 上
    量過的表格（同步分頁那兩張）從理想寬度起跳，沒量過的（歷程表）理想＝下限，
    行為跟以前完全一樣。

    為什麼不交給 ttk 的 stretch：它只有在「一開始就放得下」的時候才會分配。
    欄寬總和一旦超過表格實際拿到的寬度，ttk 會整組凍住 —— 之後把視窗拉多寬、
    把分隔線拖多開都不再重算。兩張表原本都踩在這條線上（明細表的固定寬加起來
    wide(837)，預設字級下就超出 34px；歷程表超出 85px），於是最右邊的欄位被切在
    畫面外，而 Treeview 沒有橫向捲軸，捲也捲不出來。自己算就沒有那個狀態，
    任何寬度都是當場分配。

    下限是「還讀得出來」的下限，不是好看的下限：擠到極限時寧可看到切了尾巴的
    「175,000 → 18」（ttk 是直接切，不會給省略號），也不要一整欄消失 ——
    被切掉的那一欄不會有任何跡象說它存在過，使用者只會看到一張少了一欄的表。
    """
    room = tree.winfo_width()
    if room <= 1:
        return                            # 還沒排版完，等下一次 Configure

    floors = [wide(floor) for _key, _title, _weight, floor, _anchor in spec]
    weights = [weight for _key, _title, weight, _floor, _anchor in spec]
    measured = getattr(tree, "content_widths", None)
    # 量出來的理想寬度不會低於下限：內容短的時候（例如整欄都是「300」）欄位還是
    # 要留住原本的樣子，不然表格會隨著資料一格一格抖動。
    wants = ([max(floor, want) for floor, want in zip(floors, measured)]
             if measured else list(floors))

    spare = room - sum(wants)
    if spare >= 0:
        # 放得下：多出來的照比重分，主要給會變長的那一欄（見 DETAIL_COLUMNS）。
        widths = [want + spare * weight // sum(weights)
                  for want, weight in zip(wants, weights)]
    else:
        # 放不下：從理想寬度往下限方向等比縮，可以讓的多就讓得多。縮到下限就停，
        # 之後才開始切字 —— 先切最寬鬆的那一欄，比每一欄都切一點好讀。
        slack = [want - floor for want, floor in zip(wants, floors)]
        cut = min(-spare, sum(slack))
        widths = ([want - cut * piece // sum(slack) for want, piece in zip(wants, slack)]
                  if sum(slack) else list(floors))

    # 整數除法會掉幾個像素，全部補給最後一欄，才會剛好填滿。差的是個位數像素，
    # 落在哪一欄都看不出來（同步分頁那兩張表的最後一欄是靠右對齊的數字，
    # 多幾個像素只是整欄一起往右挪一點點）。
    # 補到低於下限就不補：全部縮到底還是塞不下的時候，room - sum 是一整段負數，
    # 補下去等於把最後一欄壓扁，那一欄的字會整片不見。
    rest = room - sum(widths)
    if widths[-1] + rest >= floors[-1]:
        widths[-1] += rest

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

    ttk.Label(outer, text="Excel現金餘額改成").grid(row=row, column=0,
                                                sticky="w", pady=(10, 2))
    entry = ttk.Entry(outer, width=16, font=(family, FONT_SIZE), justify="right",
                      textvariable=text)
    entry.grid(row=row, column=1, sticky="e", padx=(24, 0), pady=(10, 2))

    # 結果邊打邊算。要核對的是「按下去會變成什麼」，自己看得到就不必先在心裡
    # 算一次再賭它跟程式算的一樣。
    ttk.Label(outer, text="會變成", style="Hint.TLabel").grid(
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
