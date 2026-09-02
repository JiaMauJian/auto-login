"""
持股同步的桌面介面。

    python ui.py

五個分頁：

    更新      讀網頁的持股與現金，寫回 Excel（這一支最早、最核心的功能）
    下單      讀 Excel 的下單試算，依序把委託填進網頁
    掛單      今天送出去的委託整批查回來，自動送出的驗證面
    歷程      誰在什麼時候改了哪一格（程式寫入、現金基準重設）
    憑證      建立／掃描／複製自動登入用的憑證

更新分頁是左右兩半：左邊一份交易人名單（誰要處理、現金剩多少），右邊只畫
選中的那一位。20 個帳號攤平成一張表是兩百多列，捲到一半連標頭都看不到是誰；
而實際的工作方式是依序輪詢、一次只處理一個人，所以「換下一位」被做成一個
動作（按鈕或 Ctrl+↑ / Ctrl+↓），右邊那張表永遠一頁看得完。

右邊分兩塊：右上角是常駐狀態列（現金算法、今日初始現金餘額、「修改」按鈕），
不管這一輪有沒有異動都在；底下是訊息框，讀歷程檔篩「今天＋這位交易人」重排
成一行一句話（見 ui_sync._fill_notes），只列這一輪真的有寫入的現金與股數／
成本，沒有異動的檔完全不提。2026/08/22 之前這裡是現金／股票兩張並排的
Treeview（見 docs/更新分頁訊息框改版.md），拿掉的理由是那兩張表大半列平常
都是「跟網頁一致，沒事」，真正要看的一兩格反而被淹沒在裡面。

一次讀全部，還是只更新一位
--------------------------
更新分頁的「範圍」跟左邊那份名單是同一個選擇的兩個入口：名單上點誰，範圍就換成誰，
「更新全部帳戶」那顆也跟著改名成「更新（王小明）帳戶」，按下去只查他一個、只寫他那一頁。
要重讀全部就把範圍切回「全部」。

範圍只管讀取。「登入」在分頁列上面那條跨分頁常駐的列上（跟「開啟EXCEL」「全部登出」
同一條），一律登入全部——2026/08/30 使用者確認不會單獨只登入一組，而要補登入某一組時，
「更新（他）帳戶」本來就會順便把他登進去。

一天下來按最多次的是「只更新一位」（盯著某一位的部位在動），一次讀全部反而只有
開盤前那一次。預設仍然是全部，因為交易人的名字要登入之後才從網站拿得到
（.env 裡只有帳密），沒讀過一輪之前左邊名單是空的，也就沒有人可以點。

只更新一位要守住：寫入、落帳只能碰這一輪讀到的那幾位（round_scope）。名單上
別人也可能有「要寫」的格子，那是用上一輪的網頁資料算出來的，順手寫出去就是
拿舊資料改 Excel。

（畫面上的資料新舊不一——別人那幾列是上一輪讀的——這件事原本靠每一位的讀取
時間標出來，訊息框第一行寫著「讀取於 10:32:07」；2026/08/22 使用者要求拿掉
那一行，目前沒有別的畫面信號補這個位置，只是不影響上面那條「寫入範圍」的
硬規則，資料看起來新舊不分不會讓程式寫錯地方，只是人自己不容易一眼看出來。）

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
手改會在下次寫入時被程式直接蓋掉。所以基準本身就寫在畫面上：右上角常駐狀態列一項
「今日初始現金餘額」，旁邊一顆「修改今日初始現金餘額」（見 ui_sync._fill_status、
edit_opening）。那顆按鈕只在今天用「初始餘額累加」的時候才出現 —— 它改的是那一種
算法的基準。

    現金餘額 = 今日初始現金餘額 + 今日淨收付

右邊那項是網頁抄來的、不會錯，所以餘額不對的時候要改的一定是左邊那項。

現金餘額有兩種算法
------------------
上面那個是其中一種（「初始餘額累加」）。另一種是「銀行餘額推算」。兩種並存
不是備援，是各有各正確的日子，公式與原因只放一份，見 docs/現金餘額兩種算法.md，
這裡不重複，只講 UI 這層特有的行為。

它是一個總開關，管全部人（大家買賣的標的一樣，要中一起中）。這次程式開起來、
第一次按「更新全部帳戶」會先跳一個視窗問今天用哪一種 —— 問在抓資料之前，因為
算法決定要不要去查銀行餘額那兩支（20 個帳號就是 40 次往返，用初始餘額累加的
日子一次都用不到）。同一次執行裡讀第二次以後不會再問；關掉程式重開，下次讀取
會再問一次（見 `cash_method_asked`，故意只放記憶體、不記進紀錄檔）。

預設一律是「初始餘額累加」：程式開起來、換一份 Excel、對話框第一次打開時
選中的都是它，選過的答案不跨執行沿用（2026/08/21 使用者要求，見
`_default_method`）。錯的方式不對稱 —— 全額交割當天留在銀行餘額推算會扣兩次、
寫進 Excel 隔天才看得出來，所以「忘了換」要往安全的那一邊倒。

問過、名字顯示出來之後，狀態列上「現金算法」那個名字本身也是一個開關：點一下先跳
確認視窗、再換成另一種（見 `ui_sync._on_method_click`、
`ui_background._toggle_cash_method`），不必等重開程式或還原 Excel 備份。2026/08/20 加的，原因是測試時每次都要還原備份才能換答案太重；
原本每次重開程式都要問一次的規矩沒有變，這是在那之上多開的第二個入口。

測得差不多之後這個入口本身也能關：`.env` 的 `CASH_METHOD_TOGGLE` 設 0，名字
還在畫面上但點不動了（見 `ui_common.cash_method_toggle_enabled`、`ui_layout.py`
建版面那段），避免正式使用時手滑誤觸。跳視窗問答那條路完全不受影響。

程式只算選中的那一種。曾經兩種都算、在現金那一列後面寫著另一種算出來多少當對帳
訊號，後來整個拿掉：差額講不出是什麼造成的（全額交割？匯撥？有人手改過 B8？），
對「今天要用哪一種」這個決定幫不上忙，而那個決定的依據本來就只有人知道。
數字一直在畫面上、按鈕一直按得動，就不必由程式決定什麼時候該問誰。

Excel 一定由程式自動更新，沒有開關
----------------------------------
「讀取」讀完就寫進 Excel，一顆按鈕從頭做到尾。原本這裡有
一個「程式自動更新」開關，可以取消勾選改成人工維護；後來拿掉了 —— 會開這支
程式，就是要讓程式自動更新，留著那個開關只是多一個不會有人關的選項。

原本寫入前會先備份一份到「備份」資料夾，2026/08/24 使用者要求拿掉——使用者
自己那邊已經有備份機制，程式這層備份只是每次寫入多一趟複製檔案的時間，
留著沒有實益。

原本每一格前面都有勾選框、外加全選／全不選、還有一顆「交還給程式」，
等於同一件事有三層開關。20 個帳號、上百格的規模下沒有人按得完，
所以先簡化成一個總開關，後來連那個總開關也拿掉了。

再往下一層：股數與成本原本還有「自動／手動」的每格狀態機 —— 程式記得自己
上次寫了什麼，Excel 現值跟記憶對不起來就自動轉成「手動」、不再覆蓋，要交還
給程式管理得靠 CLI 的 --adopt 或介面自動跑的「接管」流程。這套偵測/保護也
拿掉了：修改 Excel 的風險交給操作的人自己管控，程式一律用算出來的值覆蓋。
現金因為有「今日初始現金餘額」這個基準要顧，機制不變（見上面）。

登入完成的當下，Excel 上的現金餘額會被收成今天的基準（見 _initialize）。
那個時間點今天要買賣什麼都還沒發生，Excel 就是唯一真相，所以不必問任何
問題 —— 連現金也不必問「含不含今天的淨收付」，登入時它一定還沒含，
今天成交了什麼等「讀取」時再往上加。

畫面只做顯示與操作，所有判斷都來自 planner.py —— 介面與命令列走同一段程式碼，
才不會出現「介面算出來的結果跟命令列不一樣」這種最難查的問題。

執行緒
------
登入、抓資料、開 Excel 都很慢，全部丟到背景執行緒，否則視窗會整個凍住。
Tk 的元件只能在主執行緒碰，所以背景做完只把純資料丟進 queue，
由主執行緒定時取出來畫 —— 不在背景執行緒動任何 widget。

背景執行緒要用 COM（Excel）之前一定要先 CoInitialize，這是 Windows 的規定，
少了它 win32com 會直接丟例外。

模組拆分
--------
這個檔案本來是 SyncApp 一個類別、三千行、混雜五種職責。現在照畫面上原本就有的
五個區塊拆成五個 mixin，各自一個檔案：

    ui_common.py      字級／視窗尺寸、表格欄寬、對話框 —— 給下面五個模組共用的底層
    ui_layout.py       版面：三個分頁的元件怎麼蓋出來
    ui_cert.py          憑證分頁：建立/遷移 Profile、憑證到期提醒
    ui_background.py    背景工作：瀏覽器執行緒、登入/讀取/寫入、Excel 開啟輪詢
    ui_sync.py          更新分頁：左邊名單、右邊明細、現金那張表
    ui_history.py       歷程分頁：篩選、顯示、清除

SyncApp 把五個 mixin 疊起來，自己只留 __init__（狀態怎麼初始化）跟
_on_tab_changed（切分頁時兩個分頁都要碰的膠水邏輯）。方法之間互相呼叫全靠
共用的 self，跟拆分前完全一樣，只是方法定義搬了家。
"""

import datetime
import queue
import tkinter as tk
import traceback

import excel_io
import ledger as ledger_mod
from login import load_accounts, log_crash
from ui_background import UiBackgroundMixin
from ui_cert import UiCertMixin
from ui_common import ask_confirm, show_error
from ui_history import UiHistoryMixin
from ui_layout import UiLayoutMixin
from ui_order import UiOrderMixin
from ui_order_exec import UiOrderExecMixin
from ui_pending import UiPendingMixin
from ui_sync import UiSyncMixin


class SyncApp(UiLayoutMixin, UiCertMixin, UiBackgroundMixin, UiSyncMixin, UiHistoryMixin,
              UiOrderMixin, UiOrderExecMixin, UiPendingMixin):
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
        # 分頁名 -> {"text", "at"}：這一位的 B8 是空的，今日初始現金餘額設不成
        # （見 planner.initialize 的 blocked、ui_background._initialize）。跟
        # problem_of 一樣留著不隨每輪重讀清空，直到這一位真的讀到非空的 B8 為止。
        self.cash_baseline_errors = {}
        # 歷程檔整份讀進記憶體的結果，refresh_history() 每次 commit 完都會重填一次。
        # 這裡先給空清單：_build() 會在 refresh_history() 第一次被呼叫之前就先畫一次
        # 右邊的訊息框（見 ui_sync._fill_notes），沒有這行會在那一刻找不到這個屬性。
        self.history_rows = []
        self.current_sheet = None  # 右邊正在看哪一位交易人

        # 「第幾組帳號」與「哪一位交易人」的對照。帳號設定裡只有帳密沒有名字，
        # 名字要登入之後才從網站的 sessionStorage 拿得到，所以這份對照是一邊做
        # 一邊長出來的。模擬帳號例外 —— 它的名字本來就寫在 .env 裡，
        # 一開機就填得進去，逐一交易人更新在模擬模式下不必先讀一輪。
        self.trader_of = {i: a["name"] for i, a in enumerate(self.accounts, start=1)
                          if a.get("fake")}
        # 失敗原因改成用「第幾組」當 key，不再是一整串重來一次的清單：只更新一位的
        # 時候，別人上一輪的失敗還沒被解決，不能因為這一輪沒讀到他就當作沒事了。
        self.problem_of = {}     # 第幾組 -> {"text": 失敗原因, "at": 記錄時間}
        # 這一輪動到哪幾位。寫入、落帳都只能在這個範圍裡做 —— 別人手上那份
        # 是上一輪的舊資料，拿舊資料去寫 Excel 是這個改動最大的風險。
        self.round_scope = set()
        # 這一輪按下去的時候要做誰（None = 全部）。報告用的，不是判斷用的。
        self.round_target = None
        # 分頁名 -> 這一位最後一次讀到資料（或修改今日初始現金餘額）的時間
        # （ISO 字串）。右邊常駐狀態列那個「✓ 與網頁一致」小提示要掛的時間戳
        # 就是它（見 ui_sync._fill_status）——2026/08/22 這句話原本進訊息框、
        # 逐輪疊成一行行歷程，使用者發現多數時候都是這個結果，訊息框反而被
        # 洗版、真正的異動被淹沒，改成不進訊息框，搬到這裡當一個原地更新、
        # 不往下疊的小提示（見 docs/更新分頁訊息框改版.md）。
        self.round_at = {}
        self.busy = False
        # 「修改今日初始現金餘額」現在能不能按。_fill_status 判定，_sync_buttons 套用
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

        self._migrate_candidates = {}   # 遷移憑證那張表的列 id -> profile_tools.scan_cert_sources() 的一筆
        self.profile_busy = False       # 「建立 Profile」進行中，避免重複點

        self._pending_init_state()      # 掛單分頁的狀態，見 ui_pending.py
        self._order_init_state()        # 下單分頁的狀態，見 ui_order.py（它自己會再叫
                                        # ui_order_exec 那半邊的 _order_exec_init_state）

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
        # 錨點檢查（見 ui_background.check_excel_layout）：對不上的話這裡放那句
        # 說明，excel_open 就一直是 False，等於整支程式的 Excel 功能都停住。
        # 使用者下次按「開啟EXCEL」時清掉、重驗一次。
        self.excel_layout_problem = None
        self.excel_layout_busy = False
        if self.path is not None:
            try:
                self.ledger = ledger_mod.Ledger(self.path)
                self.ledger_fresh = not self.ledger.existed
            except RuntimeError as exc:
                self.ledger_error = str(exc)

        # 現金餘額用哪一種算法。刻意是一個總開關（多人一起切），不是一人一個 ——
        # 大家買賣的標的一樣，全額交割那天是全部一起中。
        self.cash_method = tk.StringVar(value=self._default_method())
        # 這次執行問過算法的檔案路徑。刻意只放記憶體，不寫進紀錄檔 —— 要問的是
        # 「這次程式開起來問過了沒」，不是「今天問過了沒」：同一天關掉重開，
        # 使用者可能真的換過主意（例如發現今天有全額交割），要能再問一次。
        self.cash_method_asked = set()

        self._build()
        self._drain()

        self._poll_excel()

        if self.path is None:
            self._say("還沒選檔案。按左上角「開啟EXCEL」挑一份持股管理表，選過就會記住。")
        elif not self.path.is_file():
            self._say(f"找不到 Excel：{self.path}")
        elif self.ledger_error:
            show_error(self.root, "紀錄檔有問題", self.ledger_error)
        elif not self.accounts:
            self._say("找不到帳號設定，請先在 .env 填入 TBB_ID_1 / TBB_PASSWORD_1")
        elif not self.excel_open:
            # 開機這一次只在狀態列講一句，不彈視窗。每次開程式都彈一個，看兩天就
            # 變成閉著眼睛按掉的東西，真的有事要說的時候也會被一起按掉。
            # 講的是「更新全部帳戶」跟下單的「讀取持股」，不是「登入」——
            # 登入從 2026/08/30 起就不看 Excel 了（見 ui_sync._sync_buttons）。
            self._say("這份 Excel 還沒開著 —— 按左上角「開啟EXCEL」把它打開，"
                      "「更新全部帳戶」跟下單的「讀取持股」才會亮起來。")
        else:
            self._say("按「登入」開瀏覽器並自動登入，之後要更新資料時再按「更新全部帳戶」。")

        # 開機也要畫一次。右半邊沒資料的樣子（空狀態列、空訊息框、「修改」按鈕
        # 該不該亮）是在 _fill_right 那一串裡定下來的，而那串只有 fill_sync_tree
        # 會叫到——不在這裡叫一次的話，那些狀態要等到第一次讀取（或換檔、切開關）
        # 才第一次被設對。
        self.fill_sync_tree()
        self.refresh_history()

    def _on_tab_changed(self, _event):
        """
        切到某個分頁時要先做的事。

        用分頁**名字**判斷，不用索引：索引會因為中間插一個新分頁就整排位移，而
        那種錯不會報錯，只會變成「切到憑證分頁卻跑去重讀歷程」——2026/08/30 加
        掛單分頁時這裡本來就會位移一格。
        """
        name = self.tabs.tab(self.tabs.select(), "text").strip()
        if name == "掛單":
            # 範圍選單的選項跟著「已經登入過、知道名字」的帳戶走，切過來刷一次。
            self._refresh_pending_scope()
        elif name == "下單":
            # 同上，「執行帳戶」清單也跟著已知名字走。它還會順便把還不知道的
            # 報酬率補讀回來（只讀 B22 一格，不跑巨集），因為清單是照報酬率
            # 由低到高排的——見 ui_order.refresh_order_accounts。
            self.refresh_order_accounts()
        elif name == "歷程":
            self.refresh_history()
            # 切過來通常就是想看「剛才那位」的歷程，不必再選一次人。
            if self.current_sheet and self.current_sheet in tuple(self.history_who["values"]):
                self.history_who.set(self.current_sheet)
                self._refresh_history_choices()
                self._fill_history()
        elif name == "憑證":
            self._refresh_profile_status()
            # 掃描要讀遍 Chrome／Edge 每個 profile 的 Local Storage，不便宜，
            # 不自動做，一律等使用者自己按「掃描」（2026/08/23 使用者要求：
            # 原本第一次切過去會自動掃一次，一開分頁就跑一段看不到理由的等待）。


def main():
    root = tk.Tk()

    def on_callback_error(exc, val, tb):
        """
        Tkinter 預設把按鈕/事件callback 裡沒接住的例外 print 到 stderr —— exe 是
        --windowed 打包沒有主控台，等於整個吃掉，畫面上只看得到「按下去沒反應」。
        跟 login.log_crash 共用同一份 crash.log，事後請使用者把檔案內容貼出來就好。
        """
        detail = "".join(traceback.format_exception(exc, val, tb))
        log_crash(detail)
        show_error(root, "發生未預期的錯誤",
                   f"{val}\n\n詳細內容已經寫進 crash.log。")

    root.report_callback_exception = on_callback_error

    app = SyncApp(root)

    def on_close():
        # danger=False：使用者是自己按 X 才走到這裡，本來就打算關，Enter 應該
        # 順著他的意思關掉（跟原本 messagebox.askokcancel 的預設一致）。
        if app.busy and not ask_confirm(
            root, "還在忙",
            "背景還在登入或寫入。現在關掉可能會留下寫到一半的狀態，確定要關嗎？",
            danger=False,
        ):
            return
        if app.browser_thread is not None and app.browser_thread.is_alive():
            # 關掉之前先做一次「全部登出」，跟那顆按鈕走同一條路（2026/08/31
            # 使用者要求）。差別在「登出」跟「關瀏覽器」不是同一件事：
            # clear_cookies 才是這個網站認定的登出（JSESSIONID 這類 cookie 是
            # 伺服器唯一認人的依據，見 _browser_worker 的 logout 分支），而
            # USER_DATA_DIR 指到的是持久化 profile——只送 stop 的話，收尾那段
            # 只會 close 掉瀏覽器，cookie 還躺在磁碟上，下次開起來網站可能直接
            # 把人當成還登入著。
            #
            # 兩則指令一起排隊：logout 那則做完（清 cookie、關 context 與
            # browser）才輪到 stop 讓迴圈收工。它回的那則 "logged_out" 沒有人
            # 會收——畫面正在關，不必為了它多等一輪。
            app._say("登出中，瀏覽器會自己關掉…")
            # update_idletasks 只重畫，不處理使用者事件——用 update() 的話，
            # 這個空檔裡再按一次 X 會把 on_close 整支重進一遍。
            root.update_idletasks()
            app.browser_cmd_queue.put(("logout", None))
            app.browser_cmd_queue.put(("stop", None))
            # 比原本的 10 秒寬一點：多了「清 cookie、關 context」這一段，而且
            # 前面可能還有一則指令正在跑（使用者是在忙的時候按 X 進來的）。
            app.browser_thread.join(timeout=20)
        excel_io.clear_all_markers(app.path)
        root.destroy()

    root.protocol("WM_DELETE_WINDOW", on_close)
    root.mainloop()


if __name__ == "__main__":
    main()
