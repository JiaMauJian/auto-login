# 用「一般模式」的 Chrome 開啟自動登入專用的使用者資料夾（tbbstock 的數位憑證就存在裡面）。
#
# 用法（在專案根目錄執行）：
#   .\dev_tools\setup-profile.ps1          # 開起來手動操作（第一次執行順便把資料夾建出來）
#   .\dev_tools\setup-profile.ps1 -Reset   # 先把整個資料夾砍掉重建，再開起來
#                                          # （資料夾裡已經有 tbbstock 憑證時會擋下來，要再加 -Force）
#
# 現在 GUI（tbb-login.exe --sync 的「憑證」分頁）已經把這支腳本的功能做成按鈕，
# 這支留著當 GUI 壞了時的備用手動工具。
#
# 兩個用途：
#   1. 首次設定時把資料夾初始化出來，好讓 migrate-cert.ps1 有地方把憑證複製進去。
#      這種情況下視窗開起來直接關掉就好，不用登入任何帳號。
#   2. 平常想手動下單、或要在這個 profile 裡申請憑證時，用它開視窗操作。
#
# 為什麼要獨立一個資料夾（詳見 .env 註解）：
#   - Chrome 136 之後禁止自動化工具連上「預設使用者資料夾」，所以必須另外開一個資料夾。
#   - 同一個資料夾不能同時被兩個 Chrome 開著，共用的話自動化在跑時你就不能用自己的瀏覽器。
#   - Chrome 127 之後的 App-Bound Encryption 把加密金鑰綁定資料夾路徑與這台電腦，
#     所以資料夾不能從別台電腦複製過來，每台電腦都要各自做一次首次設定。
#   - 要在這裡面手動操作（例如申請憑證）必須用「一般模式」的 Chrome，也就是本腳本；
#     自動化控制的瀏覽器會被某些網站擋下來。

param(
    [switch]$Reset,
    [switch]$Force
)

$ErrorActionPreference = "Stop"

# 這支腳本住在 dev_tools\ 底下，但 .env、chrome-profile 都在專案根目錄。
$ProjectRoot = Split-Path $PSScriptRoot -Parent

# --- 1. 從旁邊的 .env 讀出 USER_DATA_DIR，沒有就用預設值 -------------------------

$envFile = Join-Path $ProjectRoot ".env"
$rawPath = "chrome-profile"

if (Test-Path $envFile) {
    # 跟 python-dotenv 一樣：忽略註解，同一個 key 出現多次時以最後一次為準。
    foreach ($line in (Get-Content $envFile -Encoding utf8)) {
        $t = $line.Trim()
        if ($t.StartsWith("#") -or -not $t.Contains("=")) { continue }
        $k = $t.Substring(0, $t.IndexOf("=")).Trim()
        if ($k -eq "USER_DATA_DIR") {
            $rawPath = $t.Substring($t.IndexOf("=") + 1).Trim().Trim('"')
        }
    }
} else {
    Write-Host "找不到 $envFile，改用預設資料夾名稱 chrome-profile。" -ForegroundColor Yellow
}

if (-not $rawPath) {
    Write-Host "USER_DATA_DIR 是空的，代表程式每次都會用全新的空白 profile，不需要執行本腳本。" -ForegroundColor Yellow
    Write-Host "但這樣憑證也存不住，登入後會被「瀏覽器查無有效數位憑證」擋下來。" -ForegroundColor Yellow
    Write-Host "請在 .env 設定 USER_DATA_DIR=chrome-profile 後再執行一次。"
    exit 1
}

# 支援 %LOCALAPPDATA% 這類寫法；相對路徑以 .env 所在資料夾為基準（跟 login.py 的邏輯一致）。
$profilePath = [Environment]::ExpandEnvironmentVariables($rawPath)
if (-not [IO.Path]::IsPathRooted($profilePath)) {
    $profilePath = Join-Path $ProjectRoot $profilePath
}
$profilePath = [IO.Path]::GetFullPath($profilePath)

$defaultUserData = [IO.Path]::GetFullPath((Join-Path $env:LOCALAPPDATA "Google\Chrome\User Data"))
if ($profilePath -eq $defaultUserData) {
    Write-Host "USER_DATA_DIR 不能指向 Chrome 的預設使用者資料夾：" -ForegroundColor Red
    Write-Host "  $profilePath"
    Write-Host "Chrome 136 之後禁止自動化工具連上這個資料夾，程式會一直連不上而逾時。"
    Write-Host "請改成別的資料夾名稱（例如 chrome-profile）再執行一次。"
    exit 1
}

Write-Host "使用者資料夾: $profilePath"

# --- 2. 找出 chrome.exe ---------------------------------------------------------

$chrome = $null
$candidates = @(
    (Join-Path $env:ProgramFiles "Google\Chrome\Application\chrome.exe"),
    (Join-Path ${env:ProgramFiles(x86)} "Google\Chrome\Application\chrome.exe"),
    (Join-Path $env:LOCALAPPDATA "Google\Chrome\Application\chrome.exe")
)
foreach ($c in $candidates) {
    if ($c -and (Test-Path $c)) { $chrome = $c; break }
}

if (-not $chrome) {
    Write-Host "這台電腦找不到 Google Chrome。" -ForegroundColor Red
    Write-Host "請先安裝 Chrome。憑證是存在 Chrome profile 裡的，沒有 Chrome 就沒辦法做這一步。"
    exit 1
}
Write-Host "Chrome: $chrome"

# --- 3. 判斷資料夾裡有沒有 tbbstock 憑證 ----------------------------------------

# tbbstock 的數位憑證存在 profile 的 localStorage（leveldb）裡，詳見 migrate-cert.ps1 檔頭。
# -Reset 前要用它擋下誤刪，收尾時也用它告訴使用者接下來還要不要跑 migrate-cert.ps1。
function Test-ProfileHasCert {
    param([string]$Path)

    # 逐一列舉子 profile（Default、Profile 1...）底下的 leveldb。
    # 不要用 "*\Local Storage\leveldb\*" 配 -Include：路徑中間有萬用字元時 -Include 會失效，
    # 一個檔案都掃不到卻不報錯，等於防護形同虛設。
    $enc = [Text.Encoding]::GetEncoding(28591)
    foreach ($sub in @(Get-ChildItem -Path $Path -Directory -ErrorAction SilentlyContinue)) {
        $leveldb = Join-Path $sub.FullName "Local Storage\leveldb"
        if (-not (Test-Path $leveldb)) { continue }
        foreach ($f in @(Get-ChildItem -Path (Join-Path $leveldb "*") -Include *.log, *.ldb -File)) {
            try {
                if ($enc.GetString([IO.File]::ReadAllBytes($f.FullName)).Contains("TWCACertIdxRef")) {
                    return $true
                }
            } catch { }
        }
    }
    return $false
}

# --- 4. 需要的話先重建資料夾 ----------------------------------------------------

if ($Reset -and (Test-Path $profilePath)) {
    # 資料夾被開著的 Chrome 鎖住時會刪不掉，先擋下來講清楚，不然會刪一半留下壞掉的 profile。
    # 比對要忽略大小寫：.NET 的 String.Contains 區分大小寫，而命令列裡的磁碟機代號大小寫
    # 不一定跟這裡算出來的一致（例如 .env 寫 c:\ 而 Chrome 命令列是 C:\），
    # 用 Contains 會漏判成「沒人在用」，接著就把使用中的 profile 刪掉。
    # -like 預設不分大小寫；路徑裡的 [ ] 等萬用字元要先跳脫。
    $pattern = "*" + [Management.Automation.WildcardPattern]::Escape($profilePath) + "*"
    $inUse = Get-CimInstance Win32_Process -Filter "Name='chrome.exe'" |
        Where-Object { $_.CommandLine -and $_.CommandLine -like $pattern }
    if ($inUse) {
        Write-Host "這個資料夾正被 Chrome 使用中，請先關掉那個視窗再執行一次。" -ForegroundColor Red
        exit 1
    }

    # 憑證砍掉就沒了，而重新申請會讓你日常瀏覽器那張憑證失效，所以有憑證時預設擋下來。
    $hasCert = Test-ProfileHasCert $profilePath

    if ($hasCert -and -not $Force) {
        Write-Host ""
        Write-Host "擋下來了：這個資料夾裡有 tbbstock 的數位憑證。" -ForegroundColor Red
        Write-Host "  $profilePath"
        Write-Host ""
        Write-Host "-Reset 會把憑證一起刪掉，而重新申請一張會讓你日常瀏覽器那張失效。"
        Write-Host "只是要開起來手動操作的話，不用加 -Reset，直接執行 .\setup-profile.ps1 就好。"
        Write-Host "真的要重建（例如 profile 已經損毀），請先用 migrate-cert.ps1 確認來源還有憑證可以再複製一次，"
        Write-Host "然後加上 -Force：.\setup-profile.ps1 -Reset -Force"
        exit 1
    }

    Write-Host "-Reset：刪除既有資料夾..."
    if ($hasCert) {
        Write-Host "  （-Force：連同裡面的數位憑證一起刪除）" -ForegroundColor Yellow
    }
    Remove-Item -Recurse -Force $profilePath
}

if (-not (Test-Path $profilePath)) {
    New-Item -ItemType Directory -Force $profilePath | Out-Null
}

# --- 5. 開起來 ------------------------------------------------------------------

Write-Host ""
Write-Host "接下來會開一個 Chrome 視窗（一般模式，不是自動化控制的），用的就是上面那個資料夾。" -ForegroundColor Cyan
Write-Host ""
Write-Host "  首次設定：直接把整個視窗關掉就好，不用登入任何帳號（憑證跟 Google 帳號無關）。" -ForegroundColor Cyan
Write-Host "  手動操作：要下單或申請憑證就在這個視窗裡做完，再把整個視窗關掉。" -ForegroundColor Cyan
Write-Host ""
Write-Host "視窗關掉後本腳本會自動繼續。" -ForegroundColor Cyan
Write-Host ""

$chromeArgs = @(
    "--user-data-dir=`"$profilePath`"",
    "--no-first-run",
    "--no-default-browser-check",
    "https://www.tbbstock.com.tw/tbb/index/home.jsp"
)
Start-Process -FilePath $chrome -ArgumentList $chromeArgs -Wait

# --- 6. 確認資料夾建好了，並回報憑證狀態 ----------------------------------------

# Chrome 正常啟動過就會寫出 Local State 與 Default\ 這兩樣；沒有代表資料夾沒初始化成功。
$initialized = (Test-Path (Join-Path $profilePath "Local State")) -and
               (Test-Path (Join-Path $profilePath "Default"))

Write-Host ""
if (-not $initialized) {
    Write-Host "資料夾看起來沒有初始化成功。" -ForegroundColor Yellow
    Write-Host "  $profilePath"
    Write-Host "Chrome 可能沒真的開起來，或是視窗一閃就被關掉了。請再執行一次本腳本，"
    Write-Host "等 Chrome 視窗完整顯示出來之後再關掉它。"
    exit 1
}

Write-Host "資料夾已就緒: $profilePath" -ForegroundColor Green

if (Test-ProfileHasCert $profilePath) {
    Write-Host "裡面已經有 tbbstock 的數位憑證，直接執行 tbb-login.exe 即可。" -ForegroundColor Green
} else {
    Write-Host ""
    Write-Host "裡面還沒有 tbbstock 的數位憑證 —— 首次設定的話這是正常的。" -ForegroundColor Yellow
    Write-Host "接下來執行 .\migrate-cert.ps1 把日常 Chrome 裡的憑證複製進來（用法見說明文件）。"
    Write-Host "少了這一步，登入後會被橘色的「瀏覽器查無有效數位憑證」擋下來。"
}
