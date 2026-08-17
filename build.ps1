# 打包成 Windows 執行檔（dist\tbb-login.exe）
#
# 一個 exe，三種用法：
#   tbb-login.exe            自動登入
#   tbb-login.exe --sync     持股同步介面（dist\持股同步.bat 雙擊就是跑這個）
#   tbb-login.exe --update   持股同步的命令列版
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
$pyiArgs = @(
    "--noconfirm",
    "--clean",
    $pkgMode,
    "--console",
    "--name", "tbb-login",
    "--collect-all", "playwright"
)

# python-dotenv 有一個選用的 IPython 整合，PyInstaller 會順著它把 IPython、matplotlib、
# numpy 一整串都打包進來（exe 會膨脹好幾十 MB）。這裡明確排除掉。
#
# tkinter 不能排除：持股同步介面（ui.py）就是用它做的，排掉之後 exe 照樣打包成功，
# 但一執行 --sync 就當場炸掉 —— 而且只有在目標電腦上才會發現。
foreach ($m in @("IPython", "matplotlib", "numpy", "pytest", "PIL", "zmq")) {
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

# dist 裡的 .bat 跟原始碼那份不一樣：那邊跑 python ui.py，這邊要跑 exe。
# 目標電腦不一定裝了 Python，整包資料夾複製過去就只有這支 exe 可用。
#
# 內容刻意只用 ASCII：cmd.exe 解析 .bat 檔靠的是系統的 OEM 編碼，不保證是
# UTF-8。曾經實測到 rem 註解裡放中文，在某些編碼環境下會讓 cmd.exe
# 把註解的一部分誤判成指令去執行 —— 這種問題只有在目標機器上才會現形。
#
# 呼叫 exe 時用 "%~dp0tbb-login.exe" 而不是裸指令名稱：有些機器（尤其是
# 企業或做過安全強化的）會設定 NoDefaultCurrentDirectoryInExePath，
# 這會讓 cmd.exe 不去目前目錄找不帶路徑的指令 —— 帶路徑就不受這個影響，
# 而且不管從哪裡雙擊、目前目錄是什麼都一定找得到同一層的 exe。
$syncBat = @"
@echo off
rem Opens the stock-sync UI. Uses %~dp0 for both cd and the exe path
rem so it works regardless of the current directory or exe-lookup
rem security policy (some machines disable current-directory lookup
rem for bare command names via NoDefaultCurrentDirectoryInExePath).
cd /d "%~dp0"
"%~dp0tbb-login.exe" --sync
if errorlevel 1 pause
"@
Set-Content -Path "dist\持股同步.bat" -Value $syncBat -Encoding ASCII

# setup-profile.ps1 也要跟著走：新電腦上必須用它把自動登入專用的使用者資料夾建出來，
# migrate-cert.ps1 才有地方把憑證複製進去（該資料夾不能從別台電腦複製，見 .env 註解）。
Copy-Item -Force "setup-profile.ps1" "dist\setup-profile.ps1"

# migrate-cert.ps1 同理：新電腦上要把 tbbstock 的網頁版憑證從日常 Chrome profile
# 複製到自動登入用的 profile，否則登入後會被「瀏覽器查無有效數位憑證」擋下來。
Copy-Item -Force "migrate-cert.ps1" "dist\migrate-cert.ps1"

# 說明文件也一起帶走，dist 整包複製到別台電腦時就有完整的設定步驟可以照做。
Copy-Item -Force "詳細說明.md" "dist\詳細說明.md"
Copy-Item -Force "簡易說明.md" "dist\簡易說明.md"

# 另外產一份 HTML：那台電腦不一定有 Markdown 編輯器，.md 被記事本開啟時表格會糊掉。
# HTML 樣式內嵌、不連外部資源，雙擊用瀏覽器就能讀。
python -m pip install --quiet markdown
if ($LASTEXITCODE -ne 0) { throw "安裝 markdown 失敗" }
python build_docs.py
if ($LASTEXITCODE -ne 0) { throw "產生 HTML 說明文件失敗" }

Write-Host ""
Write-Host "完成。執行檔在 dist 資料夾。"
Write-Host "把 .env 複製到 exe 旁邊（同一層資料夾）再執行。"
Write-Host ""
Write-Host "  tbb-login.exe        自動登入"
Write-Host "  持股同步.bat         持股同步介面（等同 tbb-login.exe --sync）"
Write-Host ""
Write-Host "持股同步還需要 dist 裡有那個持股管理的 .xls，"
Write-Host "位置也可以在 .env 用 EXCEL_PATH 指定。"
