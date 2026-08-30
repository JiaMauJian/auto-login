"""建版面：每個分頁的元件全部在這裡蓋出來，不含任何按下去之後要做的事。"""

import tkinter as tk

import ttkbootstrap as ttk

import orders
import profile_tools
from ui_pending import PENDING_ALL
from ui_common import (
    ALL_CHOICE, APP_VERSION, CELL_PAD, FONT_SIZE, HINT_SIZE, PRICE_PENDING_TEXT, WHEN_TODAY,
    WHEN_WEEK, WINDOW_H, WINDOW_W, cash_method_toggle_enabled, col_width, pick_font, wide,
    work_area,
)


class UiLayoutMixin:
    # ---------- 版面 ----------

    def _build(self):
        family = pick_font()
        self.root.title(f"持股同步 {APP_VERSION}"
                        + (f"（模擬模式：另有 {len(self.fake_sheets)} 個假帳號）"
                           if self.fake_sheets else ""))
        # 字放大了版面就要跟著大，但不能大到超出桌面可用區 —— 視窗比可用區還高
        # 的話，最下面那行狀態列會被工作列蓋住，而那是報告寫入結果的地方。
        # .env 填了 UI_WIDTH/UI_HEIGHT 就用他填的，可用區這一關照樣要過。
        left, top_edge, right, bottom = work_area(self.root)
        room_w, room_h = right - left, bottom - top_edge
        width = min(WINDOW_W or wide(960), room_w - 80)
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
        # 最窄 900：右半邊改成常駐狀態列＋訊息框之後（2026/08/22，見
        # docs/同步分頁訊息框改版.md），不再有兩張並排表格的欄位下限撐著，
        # 底線改成「左邊名單＋右上角狀態列（現金算法／今日初始現金餘額／修改
        # 按鈕）擠在一起還看得清楚」的寬度，訊息框本身會自動換行，不怕窄。
        self.root.minsize(min(wide(900), width), min(wide(560), height))

        # ttkbootstrap：套一個主題，元件仍是原本的 ttk.* 名字（ttkbootstrap 是 ttk 的
        # 相容包裝）——theme_use 改的是 Tcl 層的樣式表，不管元件是哪個 Python 類別
        # 建出來的都吃得到。colors 存到 self，讓其他分頁模組（ui_sync 那顆負現金）
        # 也能用主題色，不必自己重複一份色碼。
        style = ttk.Style(theme="cosmo")
        self.colors = style.colors
        self._use_cheap_scrollbars(style)
        style.configure("Treeview", font=(family, FONT_SIZE), rowheight=wide(28))
        style.configure("Treeview.Heading", font=(family, FONT_SIZE, "bold"))
        # 每一格左右各留一點空（2026/08/21 使用者反應「數字貼太緊」）。預設只留
        # 4 像素，靠右的數字幾乎貼著欄位邊界，跟隔壁欄的字擠在一起。
        # 數字寫在 ui_common.CELL_PAD：這段留白是吃在格子裡面的，算欄寬的時候
        # 要扣掉同一個數字，兩邊分開寫就會欄寬看起來夠、字卻被切掉最後一位。
        style.configure("Treeview.Cell", padding=(CELL_PAD, 0))
        # 選取列統一用 cosmo 的 primary 藍 + 白字，不用預設的灰（cosmo 的 selectbg
        # 本來就是灰，不是藍）。配在這裡等於全部表格共用，不必每張表各自設一次。
        style.map("Treeview", background=[("selected", self.colors.primary)],
                  foreground=[("selected", self.colors.selectfg)])
        # 輸入框裡選取文字的顏色跟上面 Treeview 選取列同一個理由：cosmo 預設的選取色
        # 不是這個藍，兩處分開看還好，放在同一個視窗裡對照就會覺得「怎麼選取色不一樣」。
        style.configure("TEntry", selectbackground=self.colors.primary,
                        selectforeground=self.colors.selectfg)
        style.configure("TButton", font=(family, FONT_SIZE))
        style.configure("TLabel", font=(family, FONT_SIZE))
        style.configure("TCheckbutton", font=(family, FONT_SIZE))
        style.configure("TNotebook.Tab", font=(family, FONT_SIZE))
        style.configure("Hint.TLabel", font=(family, HINT_SIZE), foreground="black")
        # 「讀取」用 bootstyle="primary"（見下面 fetch_button），這裡只補粗體 ——
        # ttkbootstrap 產生的實際樣式名就是 primary.TButton，全程式只有這顆按鈕用它。
        style.configure("primary.TButton", font=(family, FONT_SIZE, "bold"))
        # 現金算法只顯示、不給選（每天讀取前那個視窗才是入口），所以要夠醒目
        # ——「今天的錢是用哪一種算法算出來的」不該要人回想。
        style.configure("Method.TLabel", font=(family, FONT_SIZE, "bold"), foreground=self.colors.primary)
        # 每天第一次讀取前那個「今天算法要選哪一種」，是那個視窗裡唯一要人做的決定 ——
        # 沒特別設定的話 ttk.Radiobutton 用的是 Tk 預設字級，比旁邊的說明文字還小。
        # 用粗體標出來就夠了，字級跟其他文字統一，不再另外放大。
        style.configure("Choice.TRadiobutton", font=(family, FONT_SIZE, "bold"))
        style.configure("Auto.TLabel", font=(family, HINT_SIZE), foreground=self.colors.success)
        style.configure("Manual.TLabel", font=(family, HINT_SIZE), foreground=self.colors.warning)
        # 買賣方向的底色，跟網站本身「買紅賣綠」的既有配色一致（見 order/
        # orderConfirmRWD.html 的 text-red/text-green），下單分頁的股票列跟
        # 執行預覽都用同一組顏色，一眼認得出方向不必看文字。ttkbootstrap 的
        # 自訂樣式名稱一定要先 configure() 登記過才會生效，只 map() 會被悄悄
        # 換回預設樣式（見 docs/Tkinter ui設計原則.md 的「ttkbootstrap 自訂樣式
        # 名稱的命名規則」一節），這裡先 configure 一次就是在做這件事。
        style.configure("Buy.TLabel", font=(family, FONT_SIZE), background="#FFDFDF")
        style.configure("Sell.TLabel", font=(family, FONT_SIZE), background="#DCF1EB")
        # Combobox 展開的那個下拉清單其實是另一個 Tk 元件（listbox），不歸 ttk 的
        # style 管，上面幾行設的字級它一個都吃不到，字級改了它還是 Tk 內建的預設值。
        # 要用 option_add 走選項資料庫這條路才碰得到它，而且要對 root 設，晚建的
        # Combobox 才也吃得到。
        self.root.option_add("*TCombobox*Listbox.font", (family, FONT_SIZE))
        self.family = family

        self.tabs = ttk.Notebook(self.root)
        self.tabs.pack(fill="both", expand=True, padx=12, pady=(6, 0))
        self.tabs.bind("<<NotebookTabChanged>>", self._on_tab_changed)

        self._build_sync_tab()
        self._build_order_tab()
        self._build_pending_tab()
        self._build_history_tab()
        self._build_cert_tab()

        # 展開的下拉清單選取列要跟 Treeview 一樣是主題藍——但 ttkbootstrap 每次
        # 套用 bootstyle 都會用 Tcl 呼叫直接蓋一次 selectbackground/foreground
        # （見 update_combobox_popdown_style），優先度比 option_add 高，蓋掉了字級
        # 那條路能用的設法。只能照它的路子，在它蓋完之後再蓋一次我們要的顏色。
        for combo in (self.account_choice, self.history_who, self.history_item, self.history_when,
                      self.order_stock_pick, self.pending_choice):
            self._paint_dropdown_selection(combo)

        # 狀態列跟進度條放同一條、同一列：狀態列說「在幹嘛」，進度條說「還在動」，
        # 是同一件事的兩種講法，本來就該擺在一起看，不是兩個各自獨立的區塊。
        # 進度條原本插在分頁「上面」，一出現整個分頁的內容就被它推低一截、
        # 讀取完再彈回去，畫面跳一下；現在跟狀態列共用這一列，寬度是分出來的，
        # 讀取中不讀取只差右邊那一小塊有沒有東西，不會動到分頁的位置。
        #
        # 用 before=self.tabs 把這一列插到分頁「上面」的位置（其實是 side="bottom"
        # 釘住畫面最下面）：pack 是先來的先分空間，分頁是 expand=True 又排在前面，
        # 視窗一縮小、剩下的空間不夠兩邊分，這一列排在後面就先被擠到 0——
        # 而它是寫入結果的地方，不該憑窗高決定看不看得到。
        bottom = ttk.Frame(self.root)
        bottom.pack(side="bottom", fill="x", before=self.tabs)

        self.status = ttk.Label(bottom, text="", style="Hint.TLabel", anchor="w", padding=(12, 6))
        self.status.pack(side="left", fill="both", expand=True)

        # 「停止」放在這條跨分頁常駐的列上（見 docs/介面規劃.md 10.2）：人按它的
        # 時候很可能正在掛單分頁看委託，而下單分頁裡那顆在別的分頁上按不到。
        # 下單分頁那顆留著當第二個入口（跟「範圍選單／左邊名單是同一個選擇的
        # 兩個入口」同一種做法）。
        # 有東西在跑才亮、沒跑就灰，不搶注意力——狀態由 _sync_stop_all_button
        # 在取件迴圈裡維持（見 ui_background._drain）。
        self.stop_all_button = ttk.Button(bottom, text="■ 停止全部操作",
                                          command=self.stop_all_operations,
                                          bootstyle="danger-outline", state="disabled")
        self.stop_all_button.pack(side="right", padx=(0, 12), pady=6)
        self._stop_all_state = "disabled"

        self.progress = ttk.Progressbar(bottom, mode="indeterminate", length=wide(140))

        # 現金算法這次執行還沒問過，這裡不會顯示出來（見 _refresh_method_label）；
        # 但「修改」按鈕在不在，是照沿用下來的算法決定的，開程式當下就要對，
        # 不能等問完才第一次設定，所以還是要呼叫一次。
        self._refresh_method_label()
        self._refresh_fetch_button()
        # 按鈕的初始亮暗也要照規則來。少了這一次，Excel 沒開著時「登入」會亮到
        # 第一次狀態改變為止 —— 而「一直沒開」正好就是不會有改變的那種情況。
        self._sync_buttons()

        self._center(width, height)

    def _use_cheap_scrollbars(self, style):
        """
        把捲軸換成不用圖片畫的那一套，只為了效能，不為了外觀。

        ttkbootstrap 的捲軸滑塊是一張帶透明邊的 9-slice 圖片（見它的
        style/builders/scrollbar.py），Tk 每次重畫都要把那張圖拉伸到滑塊的實際
        長度、逐像素做 alpha 合成 —— 成本跟滑塊的像素長度成正比，而且**沒東西
        可捲的時候滑塊是滿格，正好最貴**，那恰好是這個程式多數時候的狀態
        （股票列、帳戶名單平常都塞得下）。

        2026/08/29 實測（拉動視窗寬度，一次 resize 的中位數）：

            每多一條 ttkbootstrap 捲軸   +250 ms      換掉之後   +8 ms
            下單分頁（畫面上 4 條）      1268 ms  ->  206 ms
            同步分頁（2 條）              537 ms  ->  113 ms
            歷程分頁（1 條）              251 ms  ->   65 ms
            憑證分頁（0 條，對照組）        73 ms  ->   73 ms

        拖視窗邊框時 Windows 是一路連續送 resize 過來的，所以使用者感覺到的
        不是「慢一秒」而是整個拖曳過程都在頓。cosmo / litera / darkly 都一樣，
        換 ttkbootstrap 的別的主題救不了；Tk 內建主題（vista / clam / default）
        本來就沒這個問題，因為它們的滑塊是直接畫矩形。

        代價是滑塊變成扁平方角，不再是 ttkbootstrap 那顆內縮的圓角膠囊；粗細、
        顏色都照原本的配（見底下），所以並排看幾乎分不出來，版面也不會位移。
        """
        # 一定要用 element_create 從內建主題複製一份真正「不是圖片」的元件過來。
        # 只改 layout、把元件名字寫成 Scrollbar.thumb 是不夠的：Tk 找元件會沿著
        # 樣式名往上找，最後還是找回 ttkbootstrap 註冊的那個圖片元件
        # （它註冊的名字是 Vertical.TScrollbar.thumb），結果是「快了但畫壞」——
        # 圖片不再被拉伸，只剩頭尾兩個端帽、中間空一段。這個坑很陰險，因為
        # widget 的行為完全正確：identify() 量得到的滑塊範圍是對的，錯的只有
        # 畫出來的像素，只能靠截圖才看得出來。
        for src in ("trough", "thumb"):
            style.element_create(f"Flat.Scrollbar.{src}", "from", "clam", src)
        for axis, sticky in (("Vertical", "ns"), ("Horizontal", "we")):
            style.layout(f"{axis}.TScrollbar", [
                ("Flat.Scrollbar.trough", {"sticky": sticky, "children": [
                    ("Flat.Scrollbar.thumb", {"expand": "1", "sticky": "nswe"})]})])
            # arrowsize 是唯一給得動粗細的選項，而且一定要給：換掉 layout 之後
            # 這條捲軸的粗細本來是箭頭撐出來的，沒有箭頭又沒有 arrowsize 就會
            # 縮成 1 像素 —— 一條看不見也點不到的線，不會報錯（width、thickness
            # 這兩個名字看起來比較像，實測完全沒作用）。取 8 是照 ttkbootstrap
            # 原本的粗細（它的 _SCROLLBAR_THICKNESS 就是 8）。
            #
            # gripcount=0 是關掉 clam 畫在滑塊正中間的那三條紋，ttkbootstrap
            # 原本的滑塊是素面的，留著會多一塊原本沒有的裝飾。
            #
            # 滑塊色用 border 壓深 25%：直接用 border 在白色的槽上太淡，看不出
            # 滑塊在哪（ttkbootstrap 自己也是這樣處理，見它的 _scrollbar_thumb_color）。
            style.configure(f"{axis}.TScrollbar",
                            troughcolor=self.colors.bg,
                            background=self.colors.update_hsv(self.colors.border, vd=-0.25),
                            bordercolor=self.colors.bg,
                            darkcolor=self.colors.update_hsv(self.colors.border, vd=-0.25),
                            lightcolor=self.colors.update_hsv(self.colors.border, vd=-0.25),
                            borderwidth=0, arrowsize=wide(8), relief="flat", gripcount=0)

    def _paint_dropdown_selection(self, combobox):
        """把 combobox 展開的那份清單，選取列蓋成主題藍。

        popdown 是 ttk::combobox 內部另開的 Tcl 視窗，第一次呼叫
        PopdownWindow 才會建出來（已建過的話回傳原本那個，不會重建）。
        必須在 ttkbootstrap 蓋完它自己那份顏色之後才蓋，不然會被蓋回去。
        """
        popdown = combobox.tk.eval(f"ttk::combobox::PopdownWindow {combobox}")
        combobox.tk.call(f"{popdown}.f.l", "configure",
                          "-selectbackground", self.colors.primary,
                          "-selectforeground", self.colors.selectfg)

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

        self._build_toolbar(frame)

        # 用 PanedWindow 是因為交易人的名字長短差很多，讓使用者自己拖比我猜寬度準。
        split = ttk.Panedwindow(frame, orient="horizontal")
        split.grid(row=1, column=0, sticky="nsew")

        self._build_people(split)
        self._build_detail(split)

        frame.rowconfigure(1, weight=1)
        frame.columnconfigure(0, weight=1)

    def _build_toolbar(self, parent):
        """
        開啟EXCEL／登入／讀取——原本擺在視窗最上面、三個分頁共用，
        但這幾顆按鈕做的事只跟「同步」有關（讀網頁、寫 Excel），歷程、憑證
        兩頁都用不到，擺在分頁外面等於替另外兩頁背了不相干的操作列。移進來。
        """
        top = ttk.Frame(parent, padding=(0, 0, 0, 8))
        top.grid(row=0, column=0, sticky="ew")

        # 三顆按鈕疊成左邊那一直行，由上而下就是做事的順序：
        # 開啟EXCEL -> 登入 -> 讀取。
        #
        # 原本「讀取」擺在最右邊，動線就變成左上、左下、再橫跨整個視窗
        # ——而它是這裡最常按的一顆（登入一天一次，讀取一天很多次）。垂直對齊
        # 還多帶一個好處：後兩顆在前一步沒完成時是灰的，排成一行才看得出那是
        # 「還沒輪到」而不是「壞了」。sticky="ew" 讓三顆一樣寬，看起來才是一疊
        # 步驟而不是三顆大小不一的按鈕。
        self.excel_button = ttk.Button(top, text="開啟EXCEL", command=self.open_excel,
                                       bootstyle="primary")
        self.excel_button.grid(row=0, column=0, sticky="ew")

        self.path_label = ttk.Label(top, text=self._path_text(), style="Hint.TLabel")
        self.path_label.grid(row=0, column=1, columnspan=3, sticky="w", padx=(16, 0))

        self.login_button = ttk.Button(top, text="登入", command=self.start_login,
                                       bootstyle="primary")
        self.login_button.grid(row=1, column=0, sticky="ew", pady=(6, 0))

        # 「範圍」不只是登入哪幾組 —— 讀取也照它走，而且它跟左邊那份名單是同一個
        # 選擇的兩個入口：名單上點一位，這裡就換成那一位；這裡換一位，名單也跟著跳。
        # 一次只更新一位是常態（一整天下來按最多次的就是它），一次讀全部反而是
        # 開盤前那一次，所以入口做成「預設全部、點了誰就只做誰」。
        ttk.Label(top, text="範圍").grid(row=1, column=1, sticky="w", padx=(16, 0), pady=(6, 0))
        choice_width = 22 if self.fake_sheets else 16
        self.account_choice = ttk.Combobox(top, values=self._account_choices(), state="readonly",
                                           width=choice_width, font=(self.family, FONT_SIZE))
        self.account_choice.current(0)
        self.account_choice.grid(row=1, column=2, sticky="w", padx=(8, 0), pady=(6, 0))
        self.account_choice.bind("<<ComboboxSelected>>", self._on_scope_changed)

        # 按鈕上的字跟著範圍走：全部是「讀取全部帳戶」，選了一位就是「讀取（王小明）帳戶」
        # —— 按下去會動到誰，寫在按鈕上，不必回頭去看那個下拉選單（見 _refresh_fetch_button）。
        self.fetch_button = ttk.Button(top, text="讀取全部帳戶", bootstyle="primary", command=self.start_fetch)
        self.fetch_button.grid(row=2, column=0, sticky="ew", pady=(6, 0))

        # 「全部登出」擺在讀取下面，跟上面三顆同一直行、同樣寬——但它不是流程的
        # 第四步，是反過來結束這一段瀏覽器 session 的動作（清 cookie 再關瀏覽器，
        # 見 ui_background.start_logout_all），使用者離開座位或要換一批帳號跑之前
        # 按一次。不跟 excel_open 掛勾（見 UiSyncMixin._sync_buttons）：跟有沒有選
        # Excel 檔無關，只要瀏覽器開著就能按。
        self.logout_button = ttk.Button(top, text="全部登出", bootstyle="secondary",
                                        command=self.start_logout_all)
        self.logout_button.grid(row=3, column=0, sticky="ew", pady=(6, 0))

        # 右邊留白那一欄負責吃掉多餘寬度，左邊那一直行才不會被拉開。
        top.columnconfigure(3, weight=1)

    def _build_people(self, split):
        """左欄：所有交易人一次看完，誰要處理一眼就知道。"""
        box = ttk.Frame(split, padding=(0, 0, 8, 0))
        split.add(box, weight=0)

        head = ttk.Frame(box)
        head.grid(row=0, column=0, columnspan=2, sticky="ew")
        ttk.Label(head, text="交易人").pack(side="left")
        self.people_count = ttk.Label(head, text="", style="Hint.TLabel")
        self.people_count.pack(side="right")

        # 名單上現在只看得到姓名跟狀態。「現金餘額」那一欄 2026/08/21 使用者要求
        # 隱藏 —— 右邊狀態列與訊息框已經看得到餘額、今日初始餘額、算法，名單再列
        # 一次只是同一個數字的第二個版本，而兩個版本永遠有一個是舊的（一次只讀
        # 一位的時候，別人那一格是上一輪的）。錢是負的還是會讓名字變紅字
        # （negative 標籤），數字則回到右邊看。
        #
        # 是「隱藏」不是「拿掉」：欄位還在、值也照樣填（見 _fill_people），只是
        # displaycolumns 沒把它列出來。要讓它回來就把下面那個 displaycolumns 刪掉，
        # 一行的事，其他什麼都不用改。
        self.people = ttk.Treeview(box, columns=("cash", "flag"), displaycolumns=("flag",),
                                   show="tree headings", selectmode="browse")
        self.people.heading("#0", text="姓名")
        self.people.column("#0", width=wide(120), minwidth=wide(80), stretch=True)
        self.people.heading("cash", text="現金餘額")
        self.people.column("cash", width=wide(100), minwidth=wide(80), anchor="e", stretch=False)
        self.people.heading("flag", text="狀態")
        self.people.column("flag", width=wide(72), minwidth=wide(56), anchor="center", stretch=False)

        bar = ttk.Scrollbar(box, orient="vertical", command=self.people.yview)
        self.people.configure(yscrollcommand=bar.set)
        self.people.grid(row=1, column=0, sticky="nsew")
        bar.grid(row=1, column=1, sticky="ns")

        # 底色管「要不要處理」、前景色管「錢是不是負的」。兩件事用不同屬性，
        # 才能同時成立而不會互相蓋掉。
        self.people.tag_configure("attention", background="#eaf4ea")
        self.people.tag_configure("negative", foreground=self.colors.danger)
        self.people.bind("<<TreeviewSelect>>", self._on_person_selected)

        box.rowconfigure(1, weight=1)
        box.columnconfigure(0, weight=1)

    def _build_detail(self, split):
        """
        右欄：常駐狀態列（現金算法／今日初始現金餘額／修改按鈕）＋訊息框。

        2026/08/22 改版（見 docs/同步分頁訊息框改版.md）：原本現金／股票兩張
        Treeview 整個拿掉。現金算法、今日初始現金餘額是「今天的固定基準／設定」，
        跟這一輪有沒有異動無關，所以搬到這裡常駐顯示；其餘（現金餘額本身、
        股數／成本這一輪的異動）併進訊息框，直接讀歷程檔篩「今天＋這位交易人」
        重新排版（見 ui_sync._fill_notes），不再由介面自己另外維護一份暫存結果。
        """
        box = ttk.Frame(split)
        split.add(box, weight=1)

        status = ttk.Frame(box)
        status.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        status.columnconfigure(4, weight=1)

        # 現金算法只顯示、不給選（每天讀取前那個視窗才是入口）。沿用建版面最上面
        # 那個 Method.TLabel 樣式（跟切算法的確認語氣同一種藍字）。這次執行還沒問
        # 過就留空（見 ui_sync._fill_status）——寫一個名字上去會讓人以為已經選好，
        # 而那其實是沿用下來的預設值。
        self.method_label = ttk.Label(status, style="Method.TLabel", text="",
                                      font=(self.family, FONT_SIZE, "bold"))
        self.method_label.grid(row=0, column=0, sticky="w")
        # 點一下換算法：原本綁在現金表那一列上（見舊版 _on_cash_click），表拿掉
        # 之後這個名字本身就是唯一的入口。.env 的 CASH_METHOD_TOGGLE 設 0 就不綁，
        # 名字照樣看得到、游標也不會變手指，避免正式使用時手滑誤觸
        # （見 ui_common.cash_method_toggle_enabled）。
        if cash_method_toggle_enabled():
            self.method_label.bind("<Button-1>", self._on_method_click)

        # 今日初始現金餘額：只有今天用「初始餘額累加」時才顯示（銀行餘額推算的
        # 日子這個基準用不到），跟後面那顆「修改」按鈕同一組開關
        # （見 ui_sync._show_opening_row）。
        self.opening_label = ttk.Label(status, text="", font=(self.family, FONT_SIZE))
        self.opening_label.grid(row=0, column=1, sticky="w", padx=(16, 0))

        # 現在算出來的現金餘額：不管這一輪訊息框有沒有印出新的一行都固定顯示
        # （2026/08/22 使用者要求）——訊息框只回答「發生了什麼事」，這裡固定
        # 回答「現在是多少」，兩件事分開，不必為了看現在的餘額去訊息框裡找
        # 最後一行 [餘額更新]（見 ui_sync._fill_status）。負的比照名單上現金
        # 負的名字變紅字，同一個規矩。
        self.balance_label = ttk.Label(status, text="", font=(self.family, FONT_SIZE))
        self.balance_label.grid(row=0, column=2, sticky="w", padx=(16, 0))

        # 「這一輪讀到的跟網頁一致，不用更新」的常駐確認（2026/08/22 再訂正，
        # 見 docs/同步分頁訊息框改版.md）：原本這句話進訊息框逐行疊、按幾次讀取
        # 就疊幾行，使用者反映多數時候都是這個結果，訊息框反而被洗版，真正的
        # 異動被淹沒。改成不進訊息框，搬來這裡當一個小確認、原地更新不往下疊，
        # 跟 balance_label 同一組「現在是什麼狀態」（見 ui_sync._fill_status）。
        # success（綠）呼應打勾符號的既有語意，不是「份量較輕」的灰。
        self.quiet_label = ttk.Label(status, text="", foreground=self.colors.success,
                                     font=(self.family, FONT_SIZE))
        self.quiet_label.grid(row=0, column=3, sticky="w", padx=(16, 0))

        # 按鈕字用短版「修改」——旁邊的「今日初始現金餘額」標籤已經在說是改哪一個
        # 數字，按鈕上不必重複一次全名（餘額不是網頁抄來的，是這個基準 + 今日
        # 淨收付算出來的，所以餘額不對的時候要改的正是它，見 ui_sync.edit_opening）。
        # warning（橘）呼應既有的警示色慣例——按下去就是把自動算出來的餘額換成
        # 人填的。
        self.opening_button = ttk.Button(status, text="修改",
                                         command=self.edit_opening, bootstyle="warning-outline")
        self.opening_button.grid(row=0, column=4, sticky="w", padx=(16, 0))

        # 訊息框：今天的歷程（現金、股票這一輪的異動），最後才是提醒
        # （見 ui_sync._fill_notes）。框的標題會換成「{交易人}　訊息」
        # （見 ui_sync._fill_right）——原本表格上方那行「誰的資料」標頭拿掉之後
        # （2026/08/22），換人換得快的時候容易看著一段文字卻認不出是誰的；標題
        # 貼在框邊、不隨內容捲動，換人立刻看得到，比塞進第一行更穩妥。
        # 內容量比以前那個固定高度、沒有捲軸的框大得多——兩張表拿掉之後，原本
        # 表格裡的每一列都變成這裡的一行，裝不下的部分以前會無聲被吃掉，這裡
        # 一定要加捲軸兜底。
        self.msg_frame = ttk.LabelFrame(box, text="訊息", padding=(8, 4))
        self.msg_frame.grid(row=1, column=0, sticky="nsew")
        self.msg_frame.rowconfigure(0, weight=1)
        self.msg_frame.columnconfigure(0, weight=1)

        # 字級跟其他文字（含「現金算法」用的 Method.TLabel）統一用 FONT_SIZE
        # （2026/08/22 使用者要求）：原本用 HINT_SIZE + 2 特意放大一號，跟旁邊
        # 常駐狀態列的字對照著看大小不一致，改成同一個大小。
        self.warn_box = tk.Text(self.msg_frame, wrap="word", font=(self.family, FONT_SIZE),
                                background="#eceff1", relief="flat", state="disabled",
                                padx=8, pady=6, spacing1=1, spacing2=1, spacing3=4)
        self.warn_box.grid(row=0, column=0, sticky="nsew")
        msg_bar = ttk.Scrollbar(self.msg_frame, orient="vertical", command=self.warn_box.yview)
        self.warn_box.configure(yscrollcommand=msg_bar.set)
        msg_bar.grid(row=0, column=1, sticky="ns")

        # 原本兩張表的底色（綠＝這次要寫、藍＝剛寫過、負的整列紅字）沒有格子可以
        # 掛了，改成 tk.Text 的 tag——但這個框裡值得整行標紅的只剩「現金這一輪由
        # 正轉負」這一種（見 ui_sync._cash_line），其餘都是平舖的文字，不必再分
        # write/done/missing 三種狀態色。
        self.warn_box.tag_configure("neg", foreground=self.colors.danger)
        # [警告]／[異常] 兩種提醒要看得出份量比一般歷程重（2026/08/22 使用者
        # 要求）：警告用深黃、異常用紅——沿用既有的 warning／danger 色號，
        # 跟其他地方（`opening_button` 的 warning-outline、負現金的 danger）
        # 同一套顏色語意，不另外挑新色（見 ui_sync._warning_line／_error_line）。
        self.warn_box.tag_configure("warn", foreground=self.colors.warning)
        self.warn_box.tag_configure("err", foreground=self.colors.danger)

        # 換人只靠名單本身。原本這裡還有「上一位／下一位」兩顆按鈕，但名單就在
        # 左邊、點下去更快，那兩顆只是同一件事的第二個入口；「寫入」也拿掉了 ——
        # 程式一定會自己寫，不需要第三條路。
        # 快速鍵綁在視窗上而不是名單上：手放在訊息框的時候焦點不在名單裡，
        # 這時候還是要能換人。
        self.root.bind("<Control-Down>", lambda _event: self._step_person(1))
        self.root.bind("<Control-Up>", lambda _event: self._step_person(-1))

        box.rowconfigure(0, weight=0)
        box.rowconfigure(1, weight=1)
        box.columnconfigure(0, weight=1)

    def _build_pending_tab(self):
        """
        掛單分頁（見 docs/介面規劃.md 10.1）：上面一列查詢、中間一張表、下面
        三顆取消按鈕。

        「列出掛單」跟「取消掛單」放同一頁不是硬湊在一起——取消之前本來就要先
        看到會取消掉什麼。三顆取消按鈕現在一律 disabled：那支 API 還沒偵察過
        （9.7 第 3 步），先畫出來讓版面完整，跟下單分頁那幾個還沒接上的控制項
        同一種做法。

        「範圍」跟同步分頁的範圍選單同一個概念，但這裡預設「全部」——看掛單本來
        就是要一次看完所有帳戶。
        """
        frame = ttk.Frame(self.tabs, padding=8)
        self.tabs.add(frame, text="  掛單  ")

        top = ttk.Frame(frame, padding=(0, 0, 0, 8))
        top.grid(row=0, column=0, sticky="ew")

        ttk.Label(top, text="範圍").pack(side="left")
        self.pending_scope = tk.StringVar(value=PENDING_ALL)
        self.pending_choice = ttk.Combobox(top, textvariable=self.pending_scope, width=18,
                                           state="readonly", values=[PENDING_ALL],
                                           font=(self.family, FONT_SIZE))
        self.pending_choice.pack(side="left", padx=(8, 0))
        self.pending_button = ttk.Button(top, text="查詢掛單", command=self.refresh_pending,
                                         bootstyle="primary-outline")
        self.pending_button.pack(side="left", padx=(16, 0))
        # 時間戳靠右：這一頁的資料是「按下去那一刻的快照」，不會自己更新，
        # 什麼時候查的跟查到什麼一樣重要。
        self.pending_stamp = ttk.Label(top, text="還沒查過", style="Hint.TLabel")
        self.pending_stamp.pack(side="right")

        columns = ("sheet", "stock", "side", "price", "qty", "matched", "left", "status", "ordno")
        titles = {"sheet": "帳戶", "stock": "股票", "side": "買賣", "price": "委託價",
                  "qty": "委託量", "matched": "成交", "left": "未成交",
                  "status": "狀態", "ordno": "委託書號"}
        # 數量欄都是短數字，壓小沒關係；「狀態」可能是網站回的失敗原因（errmsg
        # 是自由文字，量不出上限），給它 stretch 吃掉剩下的寬度。
        narrow = {"side": 50, "price": 80, "qty": 80, "matched": 70, "left": 80, "ordno": 100}
        self.pending_tree = ttk.Treeview(frame, columns=columns, show="headings", height=14)
        for key in columns:
            self.pending_tree.heading(key, text=titles[key])
            self.pending_tree.column(key, width=wide(narrow.get(key, 110)),
                                     anchor="center" if key in ("side", "price", "qty", "matched", "left") else "w",
                                     stretch=(key == "status"))
        self.pending_tree.grid(row=1, column=0, sticky="nsew")
        bar = ttk.Scrollbar(frame, orient="vertical", command=self.pending_tree.yview)
        self.pending_tree.configure(yscrollcommand=bar.set)
        bar.grid(row=1, column=1, sticky="ns")
        hbar = ttk.Scrollbar(frame, orient="horizontal", command=self.pending_tree.xview)
        self.pending_tree.configure(xscrollcommand=hbar.set)
        hbar.grid(row=2, column=0, sticky="ew")
        # 還掛在外面的用買賣底色（跟下單分頁執行預覽同一組，也跟網站本身買紅賣綠
        # 一致）；已經結束的（成交、取消、失敗）淡化，不該跟還能取消的一樣醒目。
        self.pending_tree.tag_configure("buy", background="#FFDFDF")
        self.pending_tree.tag_configure("sell", background="#DCF1EB")
        self.pending_tree.tag_configure("done", foreground=self.colors.secondary)

        cancel_bar = ttk.Frame(frame)
        cancel_bar.grid(row=3, column=0, columnspan=2, sticky="w", pady=(8, 0))
        self.pending_cancel_buttons = []
        for text in ("取消全部買單", "取消全部賣單", "取消全部掛單"):
            button = ttk.Button(cancel_bar, text=text, state="disabled", bootstyle="danger-outline")
            button.pack(side="left", padx=(0, 8))
            self.pending_cancel_buttons.append(button)
        ttk.Label(cancel_bar, text="取消掛單還沒接上（那支 API 還沒偵察過，見 9.7 第 3 步）",
                  style="Hint.TLabel").pack(side="left", padx=(8, 0))

        frame.rowconfigure(1, weight=1)
        frame.columnconfigure(0, weight=1)

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
        self.history_when.set(WHEN_TODAY)   # 開分頁預設看今天，不是一路捲全部歷程
        self.history_when.pack(side="left", padx=(6, 0))
        self.history_when.bind("<<ComboboxSelected>>", lambda _event: self._fill_history())

        # 改用 Text 而不是 Treeview（2026/08/23，見 docs/同步分頁訊息框改版.md
        # 同一個理由）：欄寬固定會切字，切到的字只能靠滑鼠停留的提示框看，
        # 換成一行一筆事件、文字自動換行，就不會有「這格被切掉看不到」這件事，
        # 也不用再維護欄寬／滑鼠提示那一整套。字級跟其他文字統一用 FONT_SIZE。
        # 篩選（交易人／項目／期間）完全不受影響，改的只是 _fill_history 最後
        # 怎麼畫這一步（見 ui_history._fill_history）。
        self.history_box = tk.Text(frame, wrap="word", font=(self.family, FONT_SIZE),
                                   background="#eceff1", relief="flat", state="disabled",
                                   padx=8, pady=6, spacing1=1, spacing2=1, spacing3=4)
        self.history_box.grid(row=1, column=0, sticky="nsew")
        bar = ttk.Scrollbar(frame, orient="vertical", command=self.history_box.yview)
        self.history_box.configure(yscrollcommand=bar.set)
        bar.grid(row=1, column=1, sticky="ns")

        # 現金餘額變負的那一行標紅，跟同步頁訊息框（warn_box 的 "neg"）、
        # 左邊名單負現金同一套顏色語意（2026/08/23 使用者要求）。
        self.history_box.tag_configure("neg", foreground=self.colors.danger)

        row = ttk.Frame(frame)
        row.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(8, 0))
        self.history_hint = ttk.Label(row, text="", style="Hint.TLabel")
        self.history_hint.pack(side="left")
        ttk.Button(row, text="重新整理", command=self.refresh_history,
                  bootstyle="primary-outline").pack(side="right")
        # 清除擺在重新整理左邊，中間留一段空白：兩顆按鈕一顆是「再讀一次」、
        # 一顆是「全部收走」，手滑按錯的代價差太多，不能並排貼在一起。
        self.history_clear = ttk.Button(row, text="清除歷程", command=self.clear_history,
                                        bootstyle="danger-outline")
        self.history_clear.pack(side="right", padx=(0, 16))

        frame.rowconfigure(1, weight=1)
        frame.columnconfigure(0, weight=1)

    def _build_cert_tab(self):
        """
        把 setup-profile.ps1（建立/重建使用者資料夾）與 migrate-cert.ps1（掃描、複製憑證）
        整合進來 —— 這兩件事本來都要開 PowerShell 手動跑，現在收進同一個分頁，
        按鈕按下去就是了。
        """
        frame = ttk.Frame(self.tabs, padding=8)
        self.tabs.add(frame, text="  憑證  ")

        self._build_profile_section(frame)
        self._build_migrate_section(frame)

        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(1, weight=1)

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

        self.profile_button = ttk.Button(box, text="建立 Profile", command=self.create_profile,
                                         bootstyle="primary")
        self.profile_button.grid(row=1, column=2, sticky="w", pady=(6, 0))

        ttk.Label(box, text="會開一個 Chrome 初始化資料夾後，自動關閉。",
                  style="Hint.TLabel", wraplength=wide(760)).grid(
            row=2, column=0, columnspan=4, sticky="w", pady=(4, 0))

        box.columnconfigure(3, weight=1)
        self._refresh_profile_status()

    def _build_migrate_section(self, parent):
        """從平常在用的 Chrome/Edge 找出已經申請過的憑證，複製到自動登入用的 Profile（對應 migrate-cert.ps1）。"""
        box = ttk.LabelFrame(parent, text="遷移憑證", padding=8)
        box.grid(row=1, column=0, sticky="nsew")

        head = ttk.Frame(box)
        head.grid(row=0, column=0, sticky="ew")
        ttk.Button(head, text="掃描", command=self.scan_cert_sources,
                  bootstyle="primary-outline").pack(side="left")
        self.migrate_copy_button = ttk.Button(head, text="複製到自動登入用的 Profile",
                                              command=self.copy_selected_cert, state="disabled",
                                              bootstyle="primary-outline")
        self.migrate_copy_button.pack(side="left", padx=(8, 0))
        self.migrate_status = ttk.Label(head, text="", style="Hint.TLabel")
        self.migrate_status.pack(side="left", padx=(16, 0))

        columns = ("browser", "name", "found", "path")
        titles = {"browser": "瀏覽器", "name": "Profile", "found": "憑證", "path": "路徑"}
        widths = {"browser": 100, "name": 90, "found": 60, "path": 380}
        self.migrate_tree = ttk.Treeview(box, columns=columns, show="headings",
                                         height=6, selectmode="browse")
        for key in columns:
            self.migrate_tree.heading(key, text=titles[key])
            anchor = "center" if key == "found" else "w"
            self.migrate_tree.column(key, width=wide(widths[key]), minwidth=wide(widths[key] // 2),
                                     anchor=anchor, stretch=(key == "path"))
        self.migrate_tree.grid(row=1, column=0, sticky="nsew", pady=(6, 0))
        self.migrate_tree.bind("<<TreeviewSelect>>", self._on_migrate_select)

        box.rowconfigure(1, weight=1)
        box.columnconfigure(0, weight=1)

        box.columnconfigure(0, weight=1)

    def _build_order_tab(self):
        """
        版面分四列（見 docs/介面規劃.md 9.2）：

            1. 持股與報酬率　│　作業 ○買賣股票 ●出清股票 ○全持股交易
            ────────────────────────────────────────────────
            2. 那個作業自己的設定 —— 整列隨作業換掉（三個預建 Frame 互換）
            3. 左邊指定股票 │ 右邊 勾帳戶＋執行預覽＋執行按鈕列

        重點是**右半邊三個作業共用同一份 widget，不複製**：三個作業的差別只有
        「左邊那格要填什麼」跟「張數與價格從哪來」，帳戶勾選、執行預覽、依序
        執行、多輪、自動送出逐字相同，拆成三個分頁等於把這一整套複製三份，
        而且「現在在跑什麼」會變成三個地方要看。

        目前只有「出清股票」的行為真的接上了（見 ui_order._order_job_ready）。
        """
        frame = ttk.Frame(self.tabs, padding=8)
        self.tabs.add(frame, text="  下單  ")

        top = ttk.Frame(frame, padding=(0, 0, 0, 6))
        top.grid(row=0, column=0, sticky="ew")

        # 「持股與報酬率」排在最前面：讀資料是整個分頁的起點，先讀完才知道
        # 有哪些持股可選、B17 報酬率排序長什麼樣（2026/08/28 使用者要求調整
        # 順序——跟選哪個作業比起來，讀不讀資料才是決定接下來能不能動作的
        # 第一步）。
        self.order_refresh_button = ttk.Button(top, text="持股與報酬率",
                                               command=self.refresh_order_data,
                                               bootstyle="primary-outline")
        self.order_refresh_button.pack(side="left")

        # 「作業」三選一：買賣股票／出清股票／全持股交易（見 docs/介面規劃.md
        # 9.2）。這一列是「要做哪一種事」，下面那一列是「這種事自己的設定」，
        # 兩列的層級不一樣——中間用一條橫線斷開，不要讓兩列看起來像同一組
        # 選項（跟原本盤前/盤中與賣/買之間那條直線同一個理由）。
        ttk.Separator(top, orient="vertical").pack(side="left", fill="y", padx=(16, 0), pady=2)
        ttk.Label(top, text="作業").pack(side="left", padx=(16, 0))
        for value in (orders.JOB_TRADE, orders.JOB_CLEAR, orders.JOB_FULL):
            ttk.Radiobutton(top, text=orders.JOB_NAMES[value], variable=self.order_job,
                            value=value, style="Choice.TRadiobutton",
                            command=self._on_order_job_changed).pack(side="left", padx=(8, 0))

        self.order_status = ttk.Label(top, text="", style="Hint.TLabel")
        self.order_status.pack(side="left", padx=(16, 0))

        ttk.Separator(frame, orient="horizontal").grid(row=1, column=0, sticky="ew")

        # 第二列：三個作業各自的設定，整列互換（見 ui_order._on_order_job_changed）。
        # 三個 Frame 都先建好、疊在同一格，切作業時 grid_remove() 舊的、grid()
        # 新的，不是每次重建——grid_remove 記得住格子位置，放回來會回到原位。
        job_bar = ttk.Frame(frame, padding=(0, 6, 0, 8))
        job_bar.grid(row=2, column=0, sticky="ew")
        self.order_job_frames = {
            orders.JOB_TRADE: self._build_order_job_trade(job_bar),
            orders.JOB_CLEAR: self._build_order_job_clear(job_bar),
            orders.JOB_FULL: self._build_order_job_full(job_bar),
        }
        for key, box in self.order_job_frames.items():
            box.grid(row=0, column=0, sticky="w")
            if key != self.order_job.get():
                box.grid_remove()

        body = ttk.Panedwindow(frame, orient="horizontal")
        body.grid(row=3, column=0, sticky="nsew")
        frame.rowconfigure(3, weight=1)
        frame.columnconfigure(0, weight=1)

        self._build_order_stocks(body)
        self._build_order_right(body)
        # 執行按鈕上的字是算出來的（作業／單位／時機，見 ui_order._order_exec_label），
        # 不能只靠使用者切了什麼才更新——開機那一次也要算一次，不然按鈕會停在
        # 下面寫死的那個預設字串上。
        self._update_order_exec_ui()

    def _build_order_unit(self, parent):
        """
        「單位」整張／零股，買賣股票與出清股票兩張第二列各畫一組。兩組綁的是
        同一個變數（self.order_unit），Tk 會自己讓它們保持一致，不必手動同步。

        零股還沒實作（9.7 第 4 步），先 disabled——看得到、選不到，跟原本
        「買」那顆單選鈕同一種做法：留著讓人知道有這個東西，不整個藏起來。
        """
        ttk.Label(parent, text="單位").pack(side="left")
        for value in (orders.UNIT_LOT, orders.UNIT_ODD):
            ttk.Radiobutton(parent, text=orders.UNIT_NAMES[value], variable=self.order_unit,
                            value=value, style="Choice.TRadiobutton",
                            state="normal" if value in orders.UNITS_READY else "disabled",
                            command=self._on_order_unit_changed).pack(side="left", padx=(8, 0))

    def _build_order_job_trade(self, parent):
        """
        買賣股票的第二列：張數與價格都來自 Excel 的下單試算 M14:N18，人不必填
        任何數字，所以這一列只有「單位」跟一顆觸發 `初始化下單` 巨集的按鈕
        （M14:M18 ← O4:O8、N14:N18 ← I4:I8，見 9.1 那張表）。

        「初始化下單」會把 M14:N18 蓋掉，所以按下去會先跳確認、列出會動到哪幾個
        分頁（見 ui_order.run_init_order）。零股那半段還沒接（下單表單的交易盤別
        要選哪個值還沒人看過），所以「單位」目前只有整張選得到。
        """
        box = ttk.Frame(parent)
        self._build_order_unit(box)
        self.order_init_button = ttk.Button(box, text="初始化下單",
                                            command=self.run_init_order,
                                            bootstyle="secondary-outline")
        self.order_init_button.pack(side="left", padx=(16, 0))
        return box

    def _build_order_job_clear(self, parent):
        """
        出清股票的第二列＝2026/08 已經在用的那一排，只是從最上面那一列降下來
        （9.3 第 4 點：「盤前／盤中」對另外兩個作業沒有意義，不該是整個分頁的
        頂層開關，是這個作業自己的設定）。
        """
        box = ttk.Frame(parent)
        self._build_order_unit(box)

        ttk.Separator(box, orient="vertical").pack(side="left", fill="y", padx=(16, 0), pady=2)
        ttk.Label(box, text="時機").pack(side="left", padx=(16, 0))
        ttk.Radiobutton(box, text="盤前", variable=self.order_mode, value="pre",
                        style="Choice.TRadiobutton",
                        command=self._on_order_mode_changed).pack(side="left", padx=(8, 0))
        ttk.Radiobutton(box, text="盤中", variable=self.order_mode, value="intraday",
                        style="Choice.TRadiobutton",
                        command=self._on_order_mode_changed).pack(side="left", padx=(8, 0))

        # 追價檔數：整批共用一個值（不是像比重那樣每檔股票各自設定，使用者
        # 2026/08/28 確認過），只有盤中用得到，盤前維持看得到但打不動——理由
        # 見 _build_order_unit 那段（留著、用 disabled 講清楚「這格現在沒作用」）。
        ttk.Separator(box, orient="vertical").pack(side="left", fill="y", padx=(16, 0), pady=2)
        self.order_ticks_label = ttk.Label(box, text="追價檔數")
        self.order_ticks_label.pack(side="left", padx=(16, 0))
        self.order_ticks_entry = ttk.Entry(box, textvariable=self.order_ticks, width=4,
                                           font=(self.family, FONT_SIZE), state="disabled")
        self.order_ticks_entry.pack(side="left", padx=(4, 0))
        ttk.Label(box, text="檔", style="Hint.TLabel").pack(side="left")
        return box

    def _build_order_job_full(self, parent):
        """
        全持股交易的第二列：不必指定股票，整張與零股各自一個追價檔數，加上一顆
        觸發 `自動計算` 巨集的按鈕——那支巨集算出來的 M14:N18 就是「這一輪打算
        做什麼」的完整計畫，人看過執行預覽才按下去（9.5：兩步，不是一顆按鈕
        直通下單）。

        全部先 disabled：這個作業的行為要到 9.7 第 4 步才接上，而且 `自動計算`
        還會暫時改寫 E/F/B8（9.6 第 3 點），不是可以隨手按著玩的東西。
        """
        box = ttk.Frame(parent)
        for index, (text, var) in enumerate(
                (("整張追價", self.order_full_lot_ticks), ("零股追價", self.order_full_odd_ticks))):
            ttk.Label(box, text=text).pack(side="left", padx=(16, 0) if index else (0, 0))
            ttk.Entry(box, textvariable=var, width=4, font=(self.family, FONT_SIZE),
                      state="disabled").pack(side="left", padx=(4, 0))
            ttk.Label(box, text="檔", style="Hint.TLabel").pack(side="left")
        self.order_calc_button = ttk.Button(box, text="跑自動計算", state="disabled",
                                            bootstyle="secondary-outline")
        self.order_calc_button.pack(side="left", padx=(16, 0))
        return box

    def _build_order_stocks(self, paned):
        """
        左邊：指定股票，一檔一列。盤前模式比重／價格各自設定（見 CLAUDE.md
        的討論決定）；盤中模式只有比重，價格改成整批共用的「追價檔數」
        （上面 top 那一列）。
        """
        box = ttk.LabelFrame(paned, text="指定股票", padding=8)
        # ttk Panedwindow 沒有 pane 的 width 選項，初始寬度預設是靠子元件的
        # reqwidth 撐出來的——只調 weight 不夠，子元件（下面這個下拉選單、
        # 說明文字）本身還是會把這一格撐寬。這裡改成直接鎖死寬度
        # （grid_propagate(False)），子元件多寬都不會反過來撐大這一格，
        # 讓出來的空間才會確實跑到右邊的執行預覽 Treeview
        # （2026/08/28 使用者反映「備註」欄還是被切到看不全，調過 weight
        # 一次沒解決，這次換這個做法）。
        # 300 這個數字不是隨便抓的下限：盤前模式一列要同時塞「比重」「價格」
        # 兩組 label+entry（見 _build_order_stock_row），實測窄於這個寬度
        # 那一列最後的「元」會被切到看不見——這個 Canvas 只有垂直捲軸，横向
        # 沒有任何補救辦法，鎖寬度時不能只顧右邊、把左邊自己的欄位擠壞。
        box.configure(width=wide(300))
        box.grid_propagate(False)
        paned.add(box, weight=1)

        pick = ttk.Frame(box)
        pick.grid(row=0, column=0, sticky="ew")
        self.order_stock_pick = ttk.Combobox(pick, width=15, font=(self.family, FONT_SIZE))
        self.order_stock_pick.pack(side="left")
        # 存成屬性是因為它會跟「持股與報酬率」一起變灰：盤中模式按「新增」會
        # 附帶跑一次「更新股價」巨集（見 ui_order._refresh_added_stock_price），
        # 那是一條會動 COM 的路，不能在別人正在動同一份活頁簿的時候按下去。
        self.order_add_button = ttk.Button(pick, text="新增", command=self.add_order_stock,
                                           bootstyle="primary-outline")
        self.order_add_button.pack(side="left", padx=(8, 0))

        # 加進來的股票一樣用 Canvas＋Scrollbar 包起來（跟帳戶勾選區同一個
        # 理由）：這裡原本用 sticky="new" 的 Frame，加多了會把 LabelFrame
        # 撐得比視窗還高，超出視窗底下的部分沒有任何辦法捲到——「按新增
        # 沒反應」很可能就是新那一列其實加進去了，只是被撐到看不見的地方。
        canvas = tk.Canvas(box, highlightthickness=0)
        canvas.grid(row=1, column=0, sticky="nsew")
        stock_bar = ttk.Scrollbar(box, orient="vertical", command=canvas.yview)
        stock_bar.grid(row=1, column=1, sticky="ns")
        canvas.configure(yscrollcommand=stock_bar.set)

        self.order_stock_frame = ttk.Frame(canvas)
        window_id = canvas.create_window((0, 0), window=self.order_stock_frame, anchor="nw")
        self.order_stock_frame.bind(
            "<Configure>", lambda _e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind("<Configure>", lambda e: canvas.itemconfigure(window_id, width=e.width))

        def _on_wheel(event):
            canvas.yview_scroll(int(-event.delta / 120), "units")

        canvas.bind("<Enter>", lambda _e: canvas.bind_all("<MouseWheel>", _on_wheel))
        canvas.bind("<Leave>", lambda _e: canvas.unbind_all("<MouseWheel>"))

        box.rowconfigure(1, weight=1)
        box.columnconfigure(0, weight=1)

    def _build_order_right(self, paned):
        """右邊：選帳戶（報酬率排序）＋執行預覽。"""
        box = ttk.Frame(paned)
        paned.add(box, weight=3)

        accounts = ttk.LabelFrame(box, text="選擇執行的帳戶", padding=8)
        accounts.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        ttk.Label(accounts, text="依今年報酬率由低到高排序",
                 style="Hint.TLabel", wraplength=wide(480)).grid(
            row=0, column=0, columnspan=2, sticky="w")

        # 真的 checkbox，不是「選取列＝有勾」那種要猜互動方式的畫法（Listbox
        # 的多選模式看起來就是個藍底選取，不像勾選框）。帳戶數不固定（最多
        # 20 組），checkbox 疊起來可能比面板高，所以外面包一層 Canvas＋
        # Scrollbar 做捲動；每個帳戶的 Checkbutton 本身在 ui_order.py 裡動態建。
        canvas = tk.Canvas(accounts, height=wide(150), highlightthickness=0)
        canvas.grid(row=1, column=0, sticky="nsew", pady=(4, 0))
        acc_bar = ttk.Scrollbar(accounts, orient="vertical", command=canvas.yview)
        acc_bar.grid(row=1, column=1, sticky="ns", pady=(4, 0))
        canvas.configure(yscrollcommand=acc_bar.set)

        self.order_account_inner = ttk.Frame(canvas)
        window_id = canvas.create_window((0, 0), window=self.order_account_inner, anchor="nw")
        self.order_account_inner.bind(
            "<Configure>", lambda _e: canvas.configure(scrollregion=canvas.bbox("all")))
        # 內層 Frame 寬度跟著 Canvas 走，checkbox 才會撐滿整個面板寬度，
        # 不會因為文字比較短就縮成一小塊、右邊留一大片空白看起來像沒放滿。
        canvas.bind("<Configure>", lambda e: canvas.itemconfigure(window_id, width=e.width))

        def _on_wheel(event):
            canvas.yview_scroll(int(-event.delta / 120), "units")

        # 滑鼠移進這塊才接手滾輪，離開就放掉——不然這裡的滾輪會蓋掉整個
        # 視窗其他地方原本的滾動。
        canvas.bind("<Enter>", lambda _e: canvas.bind_all("<MouseWheel>", _on_wheel))
        canvas.bind("<Leave>", lambda _e: canvas.unbind_all("<MouseWheel>"))

        accounts.columnconfigure(0, weight=1)

        preview = ttk.LabelFrame(box, text="執行預覽（依序）", padding=8)
        preview.grid(row=1, column=0, sticky="nsew")
        columns = ("order", "sheet", "side", "stock", "held", "lots", "price", "note")
        titles = {"order": "順序", "sheet": "帳戶", "side": "買賣", "stock": "股票", "held": "持股",
                 "lots": "張數", "price": "價格", "note": "備註"}
        # 窄欄位（順序／買賣／持股／張數）都是短數字或單一個字，內容範圍本來
        # 就小，壓小一點沒關係。held 加了千分位逗號，可能到「1,234,567」這種
        # 長度，壓太窄反而看不到完整數字（Treeview 超出欄寬不會換行也不會
        # 刪節號，就是直接切掉）——這幾欄就算真的去量，量出來也跟這裡寫死的
        # 數字差不多，不值得為它們另外量寬度（2026/08/29 使用者討論過的界線）。
        narrow = {"order": 40, "side": 40, "held": 95, "lots": 60}
        # 「備註」「價格」的候選內容是固定、可列舉的（orders.py 那幾句
        # REASON_* 訊息、盤中查不到價格時的 PRICE_PENDING_TEXT），量一次寫死，
        # 不隨每次重畫變動——不是使用者亂打的自由文字，量的是「這欄可能出現
        # 的所有內容」，不會漏算也不會因為某一輪比較短就跟著縮窄。
        # ticks="99" 是追價檔數的安全上限猜值，不是真的限制（.env 那個
        # ORDER_MULTI_ROUND_CAP 同一種「抓個明顯夠用的數字」做法）。
        note_candidates = [
            f"{orders.REASON_NO_HOLDING}；{orders.REASON_NO_PRICE}",
            f"{orders.REASON_UNDER_ONE_LOT}；{orders.REASON_NO_PRICE}",
            orders.REASON_CHASE_TEMPLATE.format(ticks="99", opposite="委買一"),
            orders.REASON_CHASE_TEMPLATE.format(ticks="99", opposite="委賣一"),
            orders.REASON_CHASE_FROZEN_TEMPLATE.format(opposite="委買一", value="9,999.99"),
            orders.REASON_CHASE_FROZEN_TEMPLATE.format(opposite="委賣一", value="9,999.99"),
        ]
        price_candidates = [PRICE_PENDING_TEXT, "9,999.99"]
        wide_cols = {
            "note": col_width(self.family, note_candidates, minimum=wide(160)),
            "price": col_width(self.family, price_candidates, minimum=wide(70)),
        }
        # 「股票」「帳戶」欄的候選內容要等重新整理讀到股票清單／帳戶名單才
        # 知道，這裡先用預設寬度墊著，讀到資料後由 ui_order.py
        # _resize_order_stock_column／_resize_order_sheet_column 重新量寬。
        self.order_preview = ttk.Treeview(preview, columns=columns, show="headings", height=10)
        for key in columns:
            self.order_preview.heading(key, text=titles[key])
            centered = key in ("order", "side", "held", "lots", "price")
            self.order_preview.column(key, width=wide_cols.get(key, wide(narrow.get(key, 90))),
                                      anchor="center" if centered else "w",
                                      stretch=(key == "note"))
        self.order_preview.grid(row=0, column=0, sticky="nsew")
        prev_bar = ttk.Scrollbar(preview, orient="vertical", command=self.order_preview.yview)
        self.order_preview.configure(yscrollcommand=prev_bar.set)
        prev_bar.grid(row=0, column=1, sticky="ns")
        # 視窗真的太窄、備註還是被切掉的話，至少捲得到——不能讓內容存在
        # 卻沒有任何辦法看到全部（跟帳戶/股票那兩處要能捲動同一個道理）。
        prev_hbar = ttk.Scrollbar(preview, orient="horizontal", command=self.order_preview.xview)
        self.order_preview.configure(xscrollcommand=prev_hbar.set)
        prev_hbar.grid(row=1, column=0, sticky="ew")
        # 跳過的列（沒有這檔／比重算出來不到 1 張）淡化顯示，不是默默消失。
        # 買賣底色是 Treeview 列的 tag，不是 ttk 樣式名稱，不受上面「先
        # configure() 才生效」那條規則管，直接 tag_configure 就會生效。
        self.order_preview.tag_configure("skip", foreground=self.colors.secondary)
        self.order_preview.tag_configure("buy", background="#FFDFDF")
        self.order_preview.tag_configure("sell", background="#DCF1EB")
        self.order_preview_hint = ttk.Label(preview, text="", style="Hint.TLabel")
        self.order_preview_hint.grid(row=2, column=0, columnspan=2, sticky="w", pady=(4, 0))

        # 依序執行：一顆按鈕身兼「開始下單」與「下一筆」（見 ui_order_exec.py
        # start_order_execution），旁邊「停止」放棄這一輪；下面那行狀態文字
        # 講現在卡在第幾筆、在等什麼。
        exec_bar = ttk.Frame(preview)
        exec_bar.grid(row=3, column=0, columnspan=2, sticky="ew", pady=(8, 0))
        self.order_exec_button = ttk.Button(exec_bar, text="開始下單（依序執行）",
                                            command=self.start_order_execution,
                                            bootstyle="danger")
        self.order_exec_button.pack(side="left")
        # 「查詢委買賣」：盤中模式限定，先把清單裡股票的即時委買賣一整批查
        # 回來，讓執行預覽直接顯示算好的價格（見 ui_order.fetch_order_quotes）
        # ——2026/08/29 使用者要求，出清股票時想在按下「開始下單」之前就看到
        # 實際會用的價位，不是等依序跑到那一筆才臨時查。跟 order_ticks_entry
        # 同一個道理，盤前模式底下維持看得到但 disabled，不整個藏起來。
        self.order_quotes_button = ttk.Button(exec_bar, text="查詢委買賣",
                                              command=self.fetch_order_quotes,
                                              bootstyle="info-outline", state="disabled")
        self.order_quotes_button.pack(side="left", padx=(8, 0))
        self.order_exec_stop_button = ttk.Button(exec_bar, text="停止", command=self.stop_order_execution,
                                                 bootstyle="secondary-outline", state="disabled")
        self.order_exec_stop_button.pack(side="left", padx=(8, 0))
        # 自動送出：關（預設）＝半自動，停在確認視窗給人看；開＝程式自己按
        # 「確認」真的送出委託。用 danger 的 checkbutton 樣式，跟旁邊那顆
        # 一按下去會操作瀏覽器的紅色主按鈕給同一種「這裡要小心」的視覺提示，
        # 不要跟普通勾選框長一樣不起眼。
        ttk.Checkbutton(exec_bar, text="自動送出委託單",
                       variable=self.order_auto_confirm, command=self._on_order_auto_changed,
                       bootstyle="danger").pack(side="left", padx=(16, 0))

        # 多輪直到出清／自動更新股價：出清股票(整張) 規劃文件裡「是否要跑多輪
        # 到全部完成」這一條只列在盤中設定底下，盤前沒有（2026/08/28 使用者
        # 更正）——這兩個勾選框只在盤中模式下打得開，初始模式是盤前，所以
        # 一開始就是 disabled，切到盤中才解鎖（見 ui_order._on_order_mode_changed）。
        # 自動更新股價還多一層：就算在盤中，沒勾多輪它也沒有意義，一樣要
        # 先勾多輪才打開（見 ui_order._on_order_multi_round_changed）。
        round_bar = ttk.Frame(preview)
        round_bar.grid(row=4, column=0, columnspan=2, sticky="w", pady=(4, 0))
        self.order_multi_round_check = ttk.Checkbutton(
            round_bar, text="多輪直到出清（僅盤中）",
            variable=self.order_multi_round, state="disabled",
            command=self._on_order_multi_round_changed)
        self.order_multi_round_check.pack(side="left")
        self.order_auto_price_check = ttk.Checkbutton(
            round_bar, text="自動更新股價（每輪先觸發 Excel 的「更新股價」）",
            variable=self.order_auto_price, state="disabled")
        self.order_auto_price_check.pack(side="left", padx=(16, 0))

        self.order_exec_status = ttk.Label(preview, text="", style="Hint.TLabel", wraplength=wide(480))
        self.order_exec_status.grid(row=5, column=0, columnspan=2, sticky="w", pady=(4, 0))

        preview.rowconfigure(0, weight=1)
        preview.columnconfigure(0, weight=1)

        box.rowconfigure(1, weight=1)
        box.columnconfigure(0, weight=1)
