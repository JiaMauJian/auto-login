# 建立 / 更新自動登入用的 Chrome 使用者資料夾（Google 登入狀態就存在裡面）。
#
# 用法：
#   .\setup-profile.ps1          # 開啟資料夾對應的 Chrome，讓你手動登入 Google
#   .\setup-profile.ps1 -Reset   # 先把整個資料夾砍掉重建，再開起來登入
#                                # （資料夾裡已經有 tbbstock 憑證時會擋下來，要再加 -Force）
#
# 為什麼需要這個腳本（詳見 .env 註解）：
#   - Chrome 136 之後禁止自動化工具連上「預設使用者資料夾」，所以必須另外開一個資料夾。
#   - Chrome 127 之後的 App-Bound Encryption 把 Cookie 金鑰綁定資料夾路徑與這台電腦，
#     所以資料夾不能從別台電腦複製過來，每台電腦都要自己登入一次。
#   - 登入這一步必須用「一般模式」的 Chrome（本腳本就是），
#     用自動化控制的瀏覽器登入會被 Google 擋下來，顯示「這個瀏覽器或應用程式可能不安全」。

param(
    [switch]$Reset,
    [switch]$Force
)

$ErrorActionPreference = "Stop"

# --- 1. 從旁邊的 .env 讀出 USER_DATA_DIR，沒有就用預設值 -------------------------

$envFile = Join-Path $PSScriptRoot ".env"
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
    Write-Host "若要保留 Google 登入狀態，請在 .env 設定 USER_DATA_DIR=chrome-profile 後再執行一次。"
    exit 1
}

# 支援 %LOCALAPPDATA% 這類寫法；相對路徑以 .env 所在資料夾為基準（跟 login.py 的邏輯一致）。
$profilePath = [Environment]::ExpandEnvironmentVariables($rawPath)
if (-not [IO.Path]::IsPathRooted($profilePath)) {
    $profilePath = Join-Path $PSScriptRoot $profilePath
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
    Write-Host "請先安裝 Chrome；或者放棄 Google 登入狀態，把 .env 的 USER_DATA_DIR 留空，"
    Write-Host "程式就會改用 Playwright 內建的 Chromium。"
    exit 1
}
Write-Host "Chrome: $chrome"

# --- 3. 需要的話先重建資料夾 ----------------------------------------------------

if ($Reset -and (Test-Path $profilePath)) {
    # 資料夾被開著的 Chrome 鎖住時會刪不掉，先擋下來講清楚，不然會刪一半留下壞掉的 profile。
    $inUse = Get-CimInstance Win32_Process -Filter "Name='chrome.exe'" |
        Where-Object { $_.CommandLine -and $_.CommandLine.Contains($profilePath) }
    if ($inUse) {
        Write-Host "這個資料夾正被 Chrome 使用中，請先關掉那個視窗再執行一次。" -ForegroundColor Red
        exit 1
    }

    # tbbstock 的數位憑證存在這個資料夾的 localStorage 裡（詳見 migrate-cert.ps1 檔頭）。
    # 砍掉就沒了，而重新申請會讓你日常瀏覽器那張憑證失效，所以有憑證時預設擋下來。
    # 逐一列舉子 profile（Default、Profile 1...）底下的 leveldb。
    # 不要用 "*\Local Storage\leveldb\*" 配 -Include：路徑中間有萬用字元時 -Include 會失效，
    # 一個檔案都掃不到卻不報錯，等於防護形同虛設。
    $hasCert = $false
    $enc = [Text.Encoding]::GetEncoding(28591)
    foreach ($sub in @(Get-ChildItem -Path $profilePath -Directory -ErrorAction SilentlyContinue)) {
        $leveldb = Join-Path $sub.FullName "Local Storage\leveldb"
        if (-not (Test-Path $leveldb)) { continue }
        foreach ($f in @(Get-ChildItem -Path (Join-Path $leveldb "*") -Include *.log, *.ldb -File)) {
            try {
                if ($enc.GetString([IO.File]::ReadAllBytes($f.FullName)).Contains("TWCACertIdxRef")) {
                    $hasCert = $true
                    break
                }
            } catch { }
        }
        if ($hasCert) { break }
    }

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

# --- 4. 開起來讓使用者登入 ------------------------------------------------------

Write-Host ""
Write-Host "接下來會開一個 Chrome 視窗（一般模式，不是自動化控制的）。" -ForegroundColor Cyan
Write-Host "請在裡面登入你的 Google 帳號，登入完成後把「整個視窗關掉」，本腳本會自動繼續。" -ForegroundColor Cyan
Write-Host ""

$chromeArgs = @(
    "--user-data-dir=`"$profilePath`"",
    "--no-first-run",
    "--no-default-browser-check",
    "https://accounts.google.com/"
)
Start-Process -FilePath $chrome -ArgumentList $chromeArgs -Wait

# --- 5. 驗證登入狀態有沒有寫進去 ------------------------------------------------

Write-Host ""
$prefs = Join-Path $profilePath "Default\Preferences"
$emails = @()
if (Test-Path $prefs) {
    try {
        $json = Get-Content $prefs -Raw -Encoding utf8 | ConvertFrom-Json
        if ($json.account_info) { $emails = @($json.account_info | ForEach-Object { $_.email }) }
    } catch {
        Write-Host "讀取 Preferences 失敗: $($_.Exception.Message)" -ForegroundColor Yellow
    }
}

$cookies = Join-Path $profilePath "Default\Network\Cookies"
$hasCookies = (Test-Path $cookies) -and ((Get-Item $cookies).Length -gt 0)

if ($emails.Count -gt 0 -and $hasCookies) {
    Write-Host "完成，已登入: $($emails -join ', ')" -ForegroundColor Green
    Write-Host "以後執行程式就會帶著這個登入狀態。"
} else {
    Write-Host "看起來還沒登入成功。" -ForegroundColor Yellow
    if ($emails.Count -eq 0) { Write-Host "  - Preferences 裡沒有任何 Google 帳號記錄" }
    if (-not $hasCookies) { Write-Host "  - 找不到 Cookie 資料庫" }
    Write-Host "請再執行一次本腳本，在開啟的視窗中完成 Google 登入後再關掉視窗。"
}
