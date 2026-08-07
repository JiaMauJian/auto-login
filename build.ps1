# 打包成 Windows 執行檔（dist\tbb-login.exe）
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
# numpy、tkinter 一整串都打包進來（exe 會從 20MB 變成 90MB）。這裡明確排除掉。
foreach ($m in @("IPython", "matplotlib", "numpy", "tkinter", "pytest", "PIL", "zmq")) {
    $pyiArgs += @("--exclude-module", $m)
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

Write-Host ""
Write-Host "完成。執行檔在 dist 資料夾。"
Write-Host "把 .env 複製到 exe 旁邊（同一層資料夾）再執行。"
