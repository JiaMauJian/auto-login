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
| `*-同步歷程.jsonl` | 一行一筆異動，只增不改，「歷程」分頁與同步分頁的訊息框都讀這份 | 程式（`ledger.py`） |
| `.env` | 帳密、模擬帳號數、UI 尺寸、`CASH_METHOD_TOGGLE` 等開關 | 人 |

兩個新檔案的檔名跟著 Excel 檔名走（不是固定名稱），放在 Excel 旁邊同一個資料夾。
紀錄檔掉了不會壞事——程式會判定「所有格子都不認得」，安全地整格不寫。

## 模組地圖

| 檔案 | 職責 |
| --- | --- |
| `login.py` | 登入、開瀏覽器、cookie store（換人＝換 cookie，不是重登） |
| `fetch.py` | 抓網頁資料（`collect()`），登入完立刻抓完那一組才換下一組 |
| `recon.py` | AJAX 重放，只讀不寫，偵察新查詢用（`python recon.py 1`） |
| `planner.py` | 網頁資料 × Excel 現值 × 紀錄檔 → 一張「變更提案」清單，純計算 |
| `ledger.py` | 紀錄檔讀寫、現金基準、歷程追加 |
| `excel_io.py` | COM 開檔、讀寫 B8/E/F，只認得這三處（2026/08/24 起不再自動備份） |
| `util.py` | 數字與寬度對齊等小工具 |
| `ui.py` / `ui_layout.py` / `ui_sync.py` / `ui_background.py` / `ui_common.py` / `ui_history.py` / `ui_cert.py` | Tkinter GUI，唯一有畫面的一批檔案；背景執行緒跑 Playwright／COM，主執行緒才碰 widget |
| `dev_tools/simulate.py` | 假帳號、假網頁（`window.__SIM__`），讓 `fetch.py` 走假資料但形狀跟真 API 一樣 |
| `dev_tools/sim_excel.py` | 在 Excel 裡加/移除模擬分頁 |
| `dev_tools/check_two_accounts.py` | 多帳號「各拿各的資料」回歸測試 |

## 改動前必看：已定案、別擅自改的規則

現金餘額算法、B8 讀寫策略是這個專案裡踩過最多坑的部分，細節看 [[現金餘額兩種算法]]：

- 現金餘額兩種算法並存（方法一：初始餘額累加；方法二：銀行餘額推算），一個總開關
  多人共用，不是一人一格，兩種都「對」，只是對的日子不同。
- 現金每天重新起算，B8 無條件覆蓋（不像股數/成本要比對 `last_written`）。
- 股數／成本／現金全部一律覆蓋 Excel，沒有自動/手動偵測、沒有「接管」這個概念。
- `branchId` 查詢要加 `'1'` 前綴（`query610`、`queryBankBalance`），`transDateQuery`
  不加；`queryBankBalance` 的 `Amount` 單位是分要除以 100。這兩條錯了都不會報錯，只
  是靜靜算錯，改到這幾支查詢務必對照 `docs/現金餘額兩種算法.md` 重新核對。
- 一次讀取只動「這一輪的範圍」（`round_scope`），不能因為讀了新資料就順手把別人
  上一輪的舊提案也寫進 Excel。
- `.bat` 檔案內容只能純 ASCII（連中文註解都不行），呼叫同層 exe 用 `%~dp0` 完整路徑。
- exe 是 `--windowed` 打包，啟動失敗看 exe 旁邊的 `crash.log`。
