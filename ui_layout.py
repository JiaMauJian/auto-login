"""建版面：三個分頁的元件全部在這裡蓋出來，不含任何按下去之後要做的事。"""

import tkinter as tk

import ttkbootstrap as ttk

import profile_tools
from ui_common import (
    ALL_CHOICE, CELL_PAD, FONT_SIZE, HINT_SIZE, WHEN_TODAY, WHEN_WEEK, WINDOW_H, WINDOW_W,
    STRIPE_COLOR, build_columns, cash_method_toggle_enabled, pick_font, wide, work_area,
)

# 明細表的三欄：(欄名, 標題, 比重, 下限, 對齊)。比重與下限都照 10 級字的像素給，
# 實際寬度由 wide() 乘上字級。
#
# 這兩張表的欄寬 2026/08/21 起以「裝得下現在的內容」為準（見 ui_common.fit_to_content
# 與 ui_sync 填完表之後那兩行）：下限只在內容縮不下去的時候才用得到，比重只決定
# 「多出來的寬度給誰」。ttk 本身沒有照內容調欄寬的功能，是自己量字寬算出來的。
#
# 2026/08/21 現金那張表搬到旁邊並排之後，這張表只分到右半邊的一半寬度，下限跟著
# 調低（原本 90/110/95/60）—— 三欄的下限加起來要放得進視窗縮到最小時它拿得到的
# 寬度，不然最右邊那欄會被切在畫面外，而 Treeview 沒有橫向捲軸，捲也捲不出來。
# 一般尺寸下幾乎沒有差別：那時候寬度是按比重攤的，下限只是起跳點。
#
# 同一天再拿掉第四欄「說明」（原本 150/40）：兩張表並排之後每一欄都只剩一半寬度，
# 說明放的是整句話，一定被切尾巴 —— 而切了尾巴的句子比沒有那一欄更糟，看得到開頭
# 就會以為讀得完。這一欄的內容整批搬到底下的訊息框（見 ui_sync._fill_notes）：
# 那裡是整段寬度、會自動換行，長句子本來就該長在那裡。
# 股票這邊唯一寫得出來的說明是「網頁庫存已無此檔」，而那件事訊息框本來就有一條
# 更完整的提醒（planner.plan 裡「在網頁庫存中已不存在」那句），所以是真的沒少東西。
#
# 空出來的寬度全部給「股票」（比重 300 對 90/90）：另外兩欄是靠右對齊的數字，
# 長度封頂在「553,161 → 524,854」，給多了只是在數字左邊空出一片；股票那一欄
# 放的是「006208 富邦台50」這種會變長的名字。
DETAIL_COLUMNS = (
    ("stock", "股票", 300, 90, "w"),
    ("qty", "股數", 90, 75, "e"),
    ("cost", "成本", 90, 75, "e"),
)

# 現金表的兩欄，格式同上。說明欄跟明細表一起拿掉了（原本 250/70），理由見上面
# 那段 —— 現金這一句（「銀行餘額 + 淨收付(T+0) + 淨收付(T+1) = 893 - 238 + 0 = 655」）
# 每天都在、又是整個畫面上最長的一句，本來就是被切得最兇的那一欄。它現在是訊息框
# 的第二行（見 ui_sync._fill_notes）。
#
# 現金算法那一列沒有金額（它不是一個數目），名字改寫在「金額」欄 —— 那個標題對
# 這一列來說是不準的，但比起為了它單獨留一欄，寧可讓這一列的值跟另外兩列對齊在
# 同一個位置。點得動的是整列不是某一欄（見 ui_sync._on_cash_click），所以搬欄
# 不影響切換算法。
#
# 金額欄的下限 75 → 95：以前它只放數字，現在還要放得下「銀行餘額推算」六個字。
CASH_COLUMNS = (
    ("item", "現金", 100, 105, "w"),
    ("value", "金額", 60, 95, "e"),
)

# 歷程表的七欄，格式同上。
# 「說明」的比重跟下限都特別加重：這一欄放的是自由文字（例如「今日初始現金餘額
# （今天第一次登入）」），比其他固定格式的欄位長得多，分不到足夠空間就整段被切掉，
# 而切掉的偏偏是解釋「為什麼變成這樣」的那句話。
HISTORY_COLUMNS = (
    ("at", "時間", 120, 100, "w"),
    ("sheet", "交易人", 80, 60, "w"),
    ("label", "項目", 190, 120, "w"),
    ("change", "變化", 190, 130, "w"),
    ("by", "來源", 60, 40, "center"),
    ("note", "說明", 620, 140, "w"),
)


class UiLayoutMixin:
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
        # 最窄 1000（原本 900）：右半邊現在要並排放兩張表（現金｜股票），兩張表的
        # 欄位下限加起來就是那個數字 —— 再窄下去只能靠切欄位換，而切掉的會是
        # 「今日初始現金餘額」這種一眼要認出是哪一列的字。
        self.root.minsize(min(wide(1000), width), min(wide(560), height))

        # ttkbootstrap：套一個主題，元件仍是原本的 ttk.* 名字（ttkbootstrap 是 ttk 的
        # 相容包裝）——theme_use 改的是 Tcl 層的樣式表，不管元件是哪個 Python 類別
        # 建出來的都吃得到。colors 存到 self，讓其他分頁模組（ui_sync 那顆負現金）
        # 也能用主題色，不必自己重複一份色碼。
        style = ttk.Style(theme="cosmo")
        self.colors = style.colors
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
        self._build_history_tab()
        self._build_cert_tab()

        # 展開的下拉清單選取列要跟 Treeview 一樣是主題藍——但 ttkbootstrap 每次
        # 套用 bootstyle 都會用 Tcl 呼叫直接蓋一次 selectbackground/foreground
        # （見 update_combobox_popdown_style），優先度比 option_add 高，蓋掉了字級
        # 那條路能用的設法。只能照它的路子，在它蓋完之後再蓋一次我們要的顏色。
        for combo in (self.account_choice, self.history_who, self.history_item, self.history_when):
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
        # 隱藏 —— 現金那張表已經把餘額、今日初始餘額、算法一起攤在右邊，名單再列
        # 一次只是同一個數字的第二個版本，而兩個版本永遠有一個是舊的（一次只讀
        # 一位的時候，別人那一格是上一輪的）。錢是負的還是會讓名字變紅字
        # （negative 標籤），數字則回到右邊那張表上看。
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
        self.people.tag_configure("stripe", background=STRIPE_COLOR)
        self.people.tag_configure("attention", background="#eaf4ea")
        self.people.tag_configure("warned", foreground=self.colors.warning)
        self.people.tag_configure("negative", foreground=self.colors.danger)
        self.people.bind("<<TreeviewSelect>>", self._on_person_selected)

        box.rowconfigure(1, weight=1)
        box.columnconfigure(0, weight=1)

    def _build_detail(self, split):
        """右欄：選中那位的每一格，外加提醒與按鈕。"""
        box = ttk.Frame(split)
        split.add(box, weight=1)

        # 這裡原本有一行粗體標頭（「簡嘉懋　讀取於 19:54:23」）。2026/08/21 使用者
        # 要求把它也收進底下的訊息框當第一行（見 ui_sync._fill_notes）：右半邊
        # 由上到下本來有標頭、兩張表、訊息框四塊，而標頭跟訊息框講的是同一種話
        # ——「這份資料是誰的、什麼時候來的、有什麼要注意」。併成一塊之後表格
        # 直接從最上面開始，看資料的眼睛不必先跳過一行字。row 0 就空著不佔高度。

        # 版面照 Excel 那張「持股資料」表：一檔股票一列，股數與成本並排。
        #
        # 原本是一格一列（E4 一列、F4 另一列），5 檔股票就攤成 10 列，眼睛得自己
        # 把上下兩列配對回同一檔 —— 但在 Excel 上它們本來就是同一列的兩欄，
        # 對照的時候還要在心裡做一次轉換。計算層仍然是一格一格算，併成一列
        # 純粹是顯示層的事。
        #
        # 值平常只寫一個數字，有話要說的時候才寫成兩個（「舊 → 新」，或不會寫的
        # 「現值（網頁 …）」）。20 檔裡通常只有一兩格要動，這樣掃一眼就找得到。
        # 三欄全部可伸縮，寬度靠 wide()/build_columns 按比重攤，不用手動寫死，
        # 多出來的寬度主要給「股票」那一欄（見 DETAIL_COLUMNS）。
        #
        # minwidth 是「還讀得出來」的下限，不是好看的下限 —— 擠到極限時寧可看到
        # 「175,000 → 18…」，也不要一整欄消失。三欄的下限加起來比視窗縮到最小時
        # 右半邊拿得到的寬度還小，所以任何尺寸下三欄都留在畫面上。
        self.tree = ttk.Treeview(box, columns=[c[0] for c in DETAIL_COLUMNS],
                                 show="headings", selectmode="browse")
        # 寬度一變就重新攤一次（拖分隔線、改視窗大小都算）。
        build_columns(self.tree, DETAIL_COLUMNS)

        bar = ttk.Scrollbar(box, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=bar.set)
        # 右半邊（現金在左、股票在右，見下面那張表）。
        self.tree.grid(row=1, column=1, sticky="nsew")
        bar.grid(row=1, column=2, sticky="ns")

        # 顏色只用來輔助，不是唯一線索 —— 同一件事訊息框裡有一句話寫著
        # （見 ui_sync._fill_notes），值那一欄也看得出「舊 → 新」。
        self.tree.tag_configure("stripe", background=STRIPE_COLOR)
        self.tree.tag_configure("write", background="#eaf4ea")
        self.tree.tag_configure("done", background="#eaf0fb")
        self.tree.tag_configure("missing", foreground=self.colors.secondary)

        # 現金自己一張表，不塞進股票那張。Excel 裡它本來就在「持股資料」那張表
        # 外面（B8），而且它的網頁值是「今日淨收付」不是餘額 —— 跟股數成本不同義，
        # 擺進同一欄底下會讓人以為那是網頁上的餘額。
        #
        # 2026/08/21 由「一條 Text ＋一行標籤＋一顆按鈕」改成一張表。現金這邊要看的
        # 已經不只餘額一個數字：還有今日初始現金餘額、今天用哪一種算法，本來是同一組
        # 東西卻長成三種不同的形狀，各自對齊各自的左邊。整成表格之後兩邊同一個排法
        # （項目、值）。
        #
        # 同一天再改成左右並排（原本現金那張貼在股票那張底下）：股票那張表高度是
        # 照「手上大概會有幾檔」給的，實際上常常只有兩三檔，右邊那半塊有大半是空的
        # ——與其讓現金排在那片空白下面，不如搬到旁邊，兩張表一起收在同一段高度裡。
        # 代價是兩張表各自變窄，長句子放不下 —— 所以同一天把兩張表的「說明」欄
        # 整個拿掉，那些話改在底下的訊息框講（見 DETAIL_COLUMNS、_fill_notes）。
        #
        # 另一個代價寫在這裡：原本用 Text 是為了一段一段上色 ——「553,161 → -13,877,373」
        # 這種一正一負的字串只讓負的那個紅。Treeview 的前景色是一整列的，所以改成
        # 整列照「換完之後是不是負的」上色；被多染紅的只有箭頭左邊那個舊值，
        # 而那一列真正要看的本來就是箭頭右邊那個數字。
        # 表格跟底下那顆按鈕包成一塊，才會黏在一起 —— 直接排在 box 上的話，
        # 按鈕落在「表格那一列」的下面，而那一列的高度是由右邊比較高的股票表撐出來的，
        # 中間會空一大段。sticky 不給 "s"：現金列數比右邊少，跟著拉長只會多出
        # 一片空的表格底。
        cash_box = ttk.Frame(box)
        cash_box.grid(row=1, column=0, sticky="new", padx=(0, 12))
        cash_box.columnconfigure(0, weight=1)

        self.cash_tree = ttk.Treeview(cash_box, columns=[c[0] for c in CASH_COLUMNS],
                                      show="headings", selectmode="none", height=3)
        build_columns(self.cash_tree, CASH_COLUMNS)
        self.cash_tree.grid(row=0, column=0, sticky="ew")
        self.cash_tree.tag_configure("stripe", background=STRIPE_COLOR)
        self.cash_tree.tag_configure("neg", foreground=self.colors.danger)
        # 現金算法那一列用主題藍加粗，跟原本那個 Method.TLabel 一樣 ——
        # 「今天的錢是用哪一種算法算出來的」不該要人回想。
        self.cash_tree.tag_configure("method", foreground=self.colors.primary,
                                     font=(self.family, FONT_SIZE, "bold"))

        # 「點一下換算法」原本綁在那個名字標籤上，名字進了表格就改綁在那一列上
        # （見 _on_cash_click）。.env 的 CASH_METHOD_TOGGLE 設 0 就整個不綁，
        # 名字照樣看得到、游標也不變手指，避免正式使用時手滑誤觸
        # （見 ui_common.cash_method_toggle_enabled）。
        if cash_method_toggle_enabled():
            self.cash_tree.bind("<Button-1>", self._on_cash_click)
            self.cash_tree.bind("<Motion>", self._on_cash_motion)

        # 表格底下單獨一排：改「今日初始現金餘額」那一列的按鈕。
        #
        # 餘額不是網頁抄來的，是「今日初始現金餘額 + 今日淨收付」算出來的，所以
        # 餘額不對的時候要改的是那個數字。它一直只存在紀錄檔裡，唯一改得到它的入口
        # 曾經是程式自己判斷該不該跳的那個對話框 —— 沒跳、或跳的時候填錯，就得等明天。
        #
        # 按鈕沒辦法放進 Treeview 的格子裡，所以留在表外，字也從「修改」改成寫全名：
        # 離開了那一行之後，只寫「修改」看不出改的是哪一個數字。整排只在用
        # 「初始餘額累加」時顯示（見 _show_opening_row）——銀行餘額推算的日子
        # 每次讀取直接算好寫回 B8，這顆今天用不到。
        # warning（橘）呼應專案既有的警示色慣例（見 Manual.TLabel、名單上的 warned
        # 標籤）——按下去就是把自動算出來的餘額換成人填的。
        self.opening_row = ttk.Frame(cash_box)
        self.opening_row.grid(row=1, column=0, sticky="w", pady=(8, 0))
        self.opening_button = ttk.Button(self.opening_row, text="修改今日初始現金餘額",
                                         command=self.edit_opening, bootstyle="warning-outline")
        self.opening_button.pack(side="left")

        # spacing3：每一行（用 \n 分開的一條提醒）之間留白，不然行跟行黏在一起很擠。
        # height 從 5 加到 6 是為了補回這些留白吃掉的高度，避免原本裝得下的
        # 五條提醒現在反而被擠出框外看不到（見 _fill_notes 為什麼順序有講究）。
        # 2026/08/21 再加到 8：兩張表的「說明」欄整批搬進這個框，每天固定多一到兩行
        # （現金那句話長到會折行），不加高的話等於把原本裝得下的提醒擠掉兩條。
        # 這個框沒有捲軸，裝不下的內容不會有任何跡象 —— 高度是唯一的保險。
        # 底色跟面板本來的白幾乎沒有差（#fbfbfb），這個框在畫面上等於隱形。
        # 換成灰一點的 #eceff1，跟其他色塊（綠 #eaf4ea＝write、藍 #eaf0fb＝done）
        # 區隔開來 —— 這裡是提醒文字，不是狀態色，用中性色才不會被誤讀成別的意思。
        # 這塊在畫面上沒有任何邊界，內容又是動態換的提醒文字，容易被當成
        # 背景空白掃過去。包一層 LabelFrame 給它一個名字跟邊框，跟憑證頁
        # 那三個區塊（Profile／遷移憑證／憑證到期日）同一種手法。
        msg_frame = ttk.LabelFrame(box, text="訊息", padding=(8, 4))
        msg_frame.grid(row=2, column=0, columnspan=3, sticky="ew", pady=(12, 0))
        msg_frame.columnconfigure(0, weight=1)

        self.warn_box = tk.Text(msg_frame, height=8, wrap="word", font=(self.family, HINT_SIZE),
                                background="#eceff1", relief="flat", state="disabled",
                                padx=8, pady=6, spacing1=1, spacing2=1, spacing3=4)
        self.warn_box.grid(row=0, column=0, sticky="ew")

        # 換人只靠名單本身。原本這裡還有「上一位／下一位」兩顆按鈕，但名單就在
        # 左邊、點下去更快，那兩顆只是同一件事的第二個入口；「寫入」也拿掉了 ——
        # 程式一定會自己寫，不需要第三條路。
        # 快速鍵綁在視窗上而不是名單上：手放在右邊那張表的時候焦點不在名單裡，
        # 這時候還是要能換人。
        self.root.bind("<Control-Down>", lambda _event: self._step_person(1))
        self.root.bind("<Control-Up>", lambda _event: self._step_person(-1))

        # 持股撐死五六檔，表格卻吃掉整片高度的話，底下的東西會被推到視窗最底下。
        # 所以表格照列數決定高度（見 _fill_detail），多出來的空白由最下面那一列吸收。
        box.rowconfigure(1, weight=0)
        box.rowconfigure(3, weight=1)

        # 兩張表的寬度 2:3 分。uniform 讓這兩欄照 weight 成比例分（單給 weight 只會分
        # 「多出來」的那些寬度，兩張表的寬度就變成各自欄位下限決定的，比例不受控）。
        # 原本是對半分，那是「現金的說明欄放的是每天都在的長句子」換來的；說明欄
        # 拿掉之後現金只剩一個名字加一個數字，再佔半個畫面就是空著，寬度還給右邊
        # 那張表（它的「股票」欄放得下更長的名字）。捲軸那一欄不進這個群組，
        # 它要多寬就多寬。
        box.columnconfigure(0, weight=2, uniform="detail")
        box.columnconfigure(1, weight=3, uniform="detail")

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

        # 文字統一用預設的近黑色（cosmo 的 inputfg），不再照來源（人工／交接）上色 ——
        # 「來源」欄本身就寫著字，顏色只是多一層、還跟 cosmo 的 info／warning
        # 撞成紫橘色，反而比純黑難讀。選取列的藍底白字是共用的 Treeview 樣式
        # （見 _build 開頭 style.map），這裡不用再單獨設一次。
        self.history_tree = ttk.Treeview(frame, columns=[c[0] for c in HISTORY_COLUMNS],
                                         show="headings")
        build_columns(self.history_tree, HISTORY_COLUMNS)
        self.history_tree.tag_configure("stripe", background=STRIPE_COLOR)

        bar = ttk.Scrollbar(frame, orient="vertical", command=self.history_tree.yview)
        self.history_tree.configure(yscrollcommand=bar.set)
        self.history_tree.grid(row=1, column=0, sticky="nsew")
        bar.grid(row=1, column=1, sticky="ns")

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
        widths = {"browser": 70, "name": 90, "found": 60, "path": 380}
        self.migrate_tree = ttk.Treeview(box, columns=columns, show="headings",
                                         height=6, selectmode="browse")
        for key in columns:
            self.migrate_tree.heading(key, text=titles[key])
            anchor = "center" if key == "found" else "w"
            self.migrate_tree.column(key, width=wide(widths[key]), minwidth=wide(widths[key] // 2),
                                     anchor=anchor, stretch=(key == "path"))
        self.migrate_tree.grid(row=1, column=0, sticky="nsew", pady=(6, 0))
        self.migrate_tree.tag_configure("stripe", background=STRIPE_COLOR)
        self.migrate_tree.bind("<<TreeviewSelect>>", self._on_migrate_select)

        box.rowconfigure(1, weight=1)
        box.columnconfigure(0, weight=1)
