# 打包成 Windows 執行檔（dist\tbb-login.exe）
#
# 一個 exe：
#   tbb-login.exe              持股同步介面（不帶參數就是它，雙擊直接開 GUI）
#
# 用法：
#   .\build.ps1                  # 一般打包，exe 約 60MB，第一次在新電腦執行時會自動下載 Chromium
#   .\build.ps1 -WithBrowsers    # 連 Chromium 一起帶著走，dist 整包複製到別台電腦就能用（多約 200MB）
#   .\build.ps1 -OneDir          # 產生資料夾版（啟動比較快，但不是單一檔案）
#
# 打包完把 .env 放到 exe 旁邊即可（.env 不會被打包進去，帳密不會被封在執行檔裡）。

param(
    [switch]$WithBrowsers,
    [switch]$OneDir
)

$ErrorActionPreference = "Stop"
Set-Location $PSScriptRoot

python -m pip install --quiet --upgrade pyinstaller
if ($LASTEXITCODE -ne 0) { throw "安裝 pyinstaller 失敗" }

python -m pip install --quiet -r requirements.txt
if ($LASTEXITCODE -ne 0) { throw "安裝相依套件失敗" }

$pkgMode = if ($OneDir) { "--onedir" } else { "--onefile" }

# --collect-all playwright：把 playwright 的 node driver（driver\node.exe 與 package\）一起打包，
# 少了它 exe 會在啟動 Playwright 時失敗。
# --collect-all ttkbootstrap：主題定義、圖示等資料檔是套件內的非 .py 檔案，
# PyInstaller 靜態分析抓不到，得整包收進去，否則執行期主題會跑掉或報錯。
# --windowed：exe 現在只剩 GUI 這個預設行為（自動登入、--update 那個命令列版
# 已經整支拿掉了），而 GUI 的錯誤全部走 messagebox（見 ui_background.py 等），
# 不靠印在主控台上，所以不需要黑視窗。代價：如果 exe 連 Python 都還沒跑起來就
# 整個炸掉（例如缺 DLL），使用者會完全看不到任何錯誤訊息——這個風險是刻意
# 接受的，靠 login.py 的 log_crash() 寫進 crash.log 補救。
$pyiArgs = @(
    "--noconfirm",
    "--clean",
    $pkgMode,
    "--windowed",
    "--name", "tbb-login",
    "--collect-all", "playwright",
    "--collect-all", "ttkbootstrap"
)

# python-dotenv 有一個選用的 IPython 整合，PyInstaller 會順著它把 IPython、matplotlib、
# numpy 一整串都打包進來（exe 會膨脹好幾十 MB）。這裡明確排除掉。
#
# tkinter 不能排除：持股同步介面（ui.py）就是用它做的，排掉之後 exe 照樣打包成功，
# 但一執行就當場炸掉 —— 而且只有在目標電腦上才會發現。
#
# PIL 也不能排除：ttkbootstrap 本身依賴 Pillow（圖示、部分元件渲染要用），
# 不像以前只是被 IPython 意外拖進來的無用依賴。
foreach ($m in @("IPython", "matplotlib", "numpy", "pytest", "zmq")) {
    $pyiArgs += @("--exclude-module", $m)
}

# pywin32 用來操作 Excel。win32com.client / pythoncom 都是在函式裡才 import 的，
# 而 win32timezone 是 pywin32 自己在執行期才載入的，靜態分析看不到，得明講。
foreach ($m in @("win32com.client", "pythoncom", "pywintypes", "win32timezone")) {
    $pyiArgs += @("--hidden-import", $m)
}

if ($WithBrowsers) {
    $browsers = Join-Path $env:LOCALAPPDATA "ms-playwright"
    if (-not (Test-Path $browsers)) {
        throw "找不到 $browsers，請先執行：python -m playwright install chromium"
    }
    # 只帶 chromium 相關的，firefox/webkit 用不到就不要塞進去。
    $tmp = Join-Path $env:TEMP "tbb-login-browsers\ms-playwright"
    if (Test-Path $tmp) { Remove-Item -Recurse -Force $tmp }
    New-Item -ItemType Directory -Force $tmp | Out-Null
    Get-ChildItem $browsers -Directory |
        Where-Object { $_.Name -like "chromium*" -or $_.Name -like "winldd*" -or $_.Name -like "ffmpeg*" } |
        ForEach-Object { Copy-Item -Recurse -Force $_.FullName (Join-Path $tmp $_.Name) }
    $pyiArgs += @("--add-data", "$tmp;ms-playwright")
}

python -m PyInstaller @pyiArgs login.py
if ($LASTEXITCODE -ne 0) { throw "打包失敗" }

# 把 .env.example 一併放到 dist，方便直接改成 .env 使用。
Copy-Item -Force ".env.example" "dist\.env.example"

# 不需要另外的啟動器了：GUI 是預設行為（不帶參數就是它，見 login.py 的
# route()），exe 也打包成 --windowed，直接雙擊 tbb-login.exe 就會開介面、
# 不跳黑視窗。之前用過 .bat／.vbs 包一層去帶參數，現在不需要了。

# setup-profile.ps1 / migrate-cert.ps1 已經沒人在看，不再打包進 dist——
# 這兩件事現在都在 GUI 的「憑證」分頁按鈕就能做（見 ui.py 的 _build_cert_tab）。
# 兩支腳本還留著當 GUI 壞了時的備用手動工具，搬去 dev_tools\ 了。

# 簡易/詳細說明.md 也已經沒人在看，不再打包進 dist（archived at docs\）。
# 如果之後又需要出貨用的說明文件，用 dev_tools\build_docs.py 手動轉成 HTML。

Write-Host ""
Write-Host "完成。執行檔在 dist 資料夾。"
Write-Host "把 .env 複製到 exe 旁邊（同一層資料夾）再執行。"
Write-Host ""
Write-Host "  tbb-login.exe               持股同步介面（雙擊即可，不跳黑視窗）"
Write-Host ""
Write-Host "持股同步還需要 dist 裡有那個持股管理的 .xls，"
Write-Host "位置也可以在 .env 用 EXCEL_PATH 指定。"
