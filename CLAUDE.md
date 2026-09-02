# auto-login / 持股同步

自動登入某券商網站（tbbstock）+ Playwright 抓即時資料 + COM 寫回 Excel「持股管理」檔。
Tkinter GUI 是唯一的操作介面（雙擊 `tbb-login.exe`），命令列模式已整支拿掉。

Excel 版面完全不動（公式、巨集原地保留），程式只認得 B8（現金）、D/E/F 欄（股數／成本）
這三處，其餘資訊靠旁邊兩個檔案自己記帳。**這個「Excel 是唯一畫面、程式只覆蓋固定幾格」
的定位是刻意的、且已經定案**——曾經規劃過反過來把 Excel 的日常操作（更新股價／填
比重／自動計算／下單試算）整個搬進程式、Excel 降成單純存檔格式的方向，使用者已確認
不做，別為了方便又把這條路翻出來。

## 資料放哪裡

| 檔案 | 內容 | 誰維護 |
| --- | --- | --- |
| `持股管理-*.xls` | B8 現金餘額、E/F 欄股數與成本，其餘全是原本的公式 | 人＋程式 |
| `*-同步紀錄.json` | 現金基準（今日初始現金餘額）、上次程式寫入的值、`settings` | 程式（`ledger.py`） |
| `*-同步歷程.jsonl` | 一行一筆異動，只增不改，「歷程」分頁與更新分頁的訊息框都讀這份 | 程式（`ledger.py`） |
| `.env` | 帳密、模擬帳號數、UI 尺寸、`CASH_METHOD_TOGGLE` 等開關 | 人 |

兩個新檔案的檔名跟著 Excel 檔名走（不是固定名稱），放在 Excel 旁邊同一個資料夾。
紀錄檔掉了不會壞事——程式會判定「所有格子都不認得」，安全地整格不寫。

## 模組地圖

| 檔案 | 職責 |
| --- | --- |
| `login.py` | 登入、開瀏覽器 |
| `fetch.py` | 抓網頁資料（`collect()`）、cookie store（換人＝換 cookie，不是重登），登入完立刻抓完那一組才換下一組 |
| `recon.py` | AJAX 重放，只讀不寫，偵察新查詢用（`python recon.py 1`） |
| `order_query.py` | 掛單查詢正式版（`queryOrder`），對 `order_recon.py` 就像 `fetch.py` 對 `recon.py` |
| `planner.py` | 網頁資料 × Excel 現值 × 紀錄檔 → 一張「變更提案」清單，純計算 |
| `ledger.py` | 紀錄檔讀寫、現金基準、歷程追加 |
| `excel_io.py` | COM 開檔、讀寫 B8/E/F，只認得這三處（2026/08/24 起不再自動備份） |
| `util.py` | 數字與寬度對齊等小工具 |
| `ui.py` / `ui_layout.py` / `ui_sync.py` / `ui_background.py` / `ui_common.py` / `ui_history.py` / `ui_cert.py` | Tkinter GUI，唯一有畫面的一批檔案；背景執行緒跑 Playwright／COM，主執行緒才碰 widget |
| `ui_order.py` / `ui_order_exec.py` | 下單分頁：前者收設定、讀 Excel、算執行預覽，後者是按下「開始下單」之後的依序執行引擎（吃凍結好的 queue，不管那份 queue 是哪個作業產生的） |
| `ui_pending.py` | 掛單分頁：把今天送出去的委託整批查回來攤成一張表，是自動送出的驗證面 |
| `dev_tools/simulate.py` | 假帳號、假網頁（`window.__SIM__`），讓 `fetch.py` 走假資料但形狀跟真 API 一樣 |
| `dev_tools/sim_excel.py` | 在 Excel 裡加/移除模擬分頁 |
| `dev_tools/check_two_accounts.py` | 多帳號「各拿各的資料」回歸測試 |

## 改動前必看：已定案、別擅自改的規則

現金餘額算法、B8 讀寫策略是這個專案裡踩過最多坑的部分，細節看 [[現金餘額兩種算法]]：

- 現金餘額兩種算法並存（方法一：初始餘額累加；方法二：銀行餘額推算），一個總開關
  多人共用，不是一人一格，兩種都「對」，只是對的日子不同。
- 現金每天重新起算，B8 無條件覆蓋（不像股數/成本要比對 `last_written`）。
- **只認 `持股管理-台美股-10家.xls` 那一版**（一頁 10 檔，2026/09/02 使用者定案）。
  位置全部由 `excel_io.STOCK_LIMIT`（＝巨集的 `Def_stock_limit`）推出來：持股
  D4:D13、下單試算 M19:N28、今年報酬率 A22/B22——**要換格式只改那一個數字**，
  不要回頭把哪一個範圍寫死。5 檔那一版（今年報酬率在 A17）已經不支援：開檔時會
  檢查 A22 是不是「今年報酬率」，對不上就整個擋住（`excel_io.layout_problem` ＋
  `ui_background.check_excel_layout`），不是警告一下還能用——版面對不上代表要寫的
  E/F、要讀的 M/N 全部落在別人的格子上，而且不會報錯。曾經為了相容三種格式改成
  掃 A 欄找標籤（同一天早上），定案之後拿掉了，別再加回來。
- 股數／成本／現金全部一律覆蓋 Excel，沒有自動/手動偵測、沒有「接管」這個概念。
- `branchId` 查詢要加 `'1'` 前綴（`query610`、`queryBankBalance`），`transDateQuery`
  不加；`queryBankBalance` 的 `Amount` 單位是分要除以 100。這兩條錯了都不會報錯，只
  是靜靜算錯，改到這幾支查詢務必對照 `docs/現金餘額兩種算法.md` 重新核對。
- 一次讀取只動「這一輪的範圍」（`round_scope`），不能因為讀了新資料就順手把別人
  上一輪的舊提案也寫進 Excel。
- Excel 已有的巨集（`Module1.更新股價`、`Module1.自動計算`）用的是無限定的
  `Range()`，**只對 ActiveSheet 動作**——一定要一個分頁 Activate 一次、跑一次，不能
  整批只呼叫一次。這條錯了不會報錯、不會少一欄，只是除了使用者剛好停留的那一頁
  以外全部讀到舊數字（2026/08/29 實測確認過，見 `excel_io.run_update_price_macro`）。
- `自動計算` 巨集會**暫時改寫 E/F/B8 再還原**，跟更新分頁「E/F/B8 一律覆蓋」
  正面衝突：它執行期間更新分頁的讀取／寫入要鎖住，反之亦然。
- 承上，**所有會用 COM 開活頁簿的路一律走 `excel_io.opened()`**（它順便鎖住「一次
  只有一條執行緒在動這份檔」），畫面那一層則問 `_excel_in_use()` 決定按鈕能不能按。
  接上的是使用者眼前那個 Excel 實例，不是各開各的——兩條執行緒交錯 `Activate` 會
  讓巨集跑在別人剛切過去的那一頁上，一樣不報錯、只是靜靜讀到舊 I4:I13。
- 要跳訊息一律用 `ui_common` 自己畫的那幾顆：通知走 `show_error` / `show_warning` /
  `show_info`（parent 放第一個參數），要人回答是非走 `ask_confirm`。**不要用
  `tkinter.messagebox`**——原生的是 Windows 系統對話框，字級不跟著 `UI_FONT_SIZE` 走，
  介面字調大了它還是系統預設的小字。也**不要換成 `ttkbootstrap.dialogs.Messagebox`**：
  它斷行用 `textwrap.wrap(width=50)`，那個 50 數的是字元個數，中文會從第 50 個字中間
  硬切、檔案路徑照切、視窗被撐到 1133 像素寬，按鈕還是英文 OK（2026/08/30 實測後定案）。
- `.bat` 檔案內容只能純 ASCII（連中文註解都不行），呼叫同層 exe 用 `%~dp0` 完整路徑。
- exe 是 `--windowed` 打包，啟動失敗看 exe 旁邊的 `crash.log`。
