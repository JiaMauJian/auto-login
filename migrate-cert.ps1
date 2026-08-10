# 把 tbbstock 的網頁版數位憑證從你平常在用的瀏覽器 profile 複製到自動登入用的 profile。
#
# 用法：
#   .\migrate-cert.ps1                     # 只掃描，列出哪些 profile 裡有憑證（不會動到任何檔案）
#   .\migrate-cert.ps1 -From "<profile路徑>"  # 實際複製（路徑是含 Local Storage 的那層，例如 ...\User Data\Default）
#   .\migrate-cert.ps1 -From "..." -Force  # 瀏覽器有背景程序殘留、確認過沒在用時強制執行
#
# 憑證在另一台電腦時：
#   1. 在有憑證那台跑 .\migrate-cert.ps1（純掃描），看出是哪一個 profile；
#   2. 把那個 profile 的「Local Storage」整個資料夾壓縮帶到這台（裡面有私鑰，當成密碼看待）；
#   3. 解壓成 <任意路徑>\Local Storage，然後 -From "<任意路徑>"（給父層，不是 Local Storage 本身）。
#   跨機複製沒有同機那麼可靠（台網有 TWCAAccountUqiDeviceIDByOU 裝置綁定），失敗就只能重新申請。
#
# 為什麼是複製 Local Storage：
#   - tbbstock 用台網 WebCA（tbb/js/TWCA_Util.js、twcaCryptoLib.js），憑證存在
#     https://www.tbbstock.com.tw 這個網域的 localStorage（key: TWCACertIdxRef 等），
#     不在 Windows 憑證存放區、也不在 IndexedDB，所以整包 Local Storage 資料夾就是憑證的載體。
#   - 複製檔案不會經過 RA（www.tbbra.com.tw），不是「憑證申請(ApplyCert)」，
#     所以來源那張憑證不會被作廢，複製後兩邊都能用。
#   - 不能直接把 USER_DATA_DIR 指向 Chrome 預設資料夾：Chrome 136 之後禁止自動化連上它，
#     而且同一個資料夾不能同時被兩個 Chrome 開著（詳見 setup-profile.ps1）。
#
# 注意：Local Storage 是整個 profile 共用一個 leveldb，複製會把來源 profile
#       所有網站的 localStorage 一起帶過來。功能上無害，但如果只想搬 tbbstock 那幾個 key，
#       要改用 DevTools 手動 copy/paste，沒辦法用複製檔案做到。

param(
    [string]$From,
    [switch]$Force
)

$ErrorActionPreference = "Stop"

# leveldb 檔案裡出現這些字串，代表這個 profile 存過 tbbstock 的憑證資料。
$markers = @("TWCACertIdxRef", "tbbstock")

# --- 1. 從 .env 讀出 USER_DATA_DIR，算出目標 profile ---------------------------

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
        if ($k -eq "BROWSER_PROFILE_DIR") {
            $envProfileDir = $t.Substring($t.IndexOf("=") + 1).Trim().Trim('"')
        }
    }
}

if (-not $rawPath) {
    Write-Host "USER_DATA_DIR 是空的，程式每次都會開全新的空白 profile，憑證放不住。" -ForegroundColor Red
    Write-Host "請先在 .env 設定 USER_DATA_DIR=chrome-profile，並執行 setup-profile.ps1。"
    exit 1
}

$targetUserData = [Environment]::ExpandEnvironmentVariables($rawPath)
if (-not [IO.Path]::IsPathRooted($targetUserData)) {
    $targetUserData = Join-Path $PSScriptRoot $targetUserData
}
$targetUserData = [IO.Path]::GetFullPath($targetUserData)

$profileDirName = "Default"
if ($envProfileDir) { $profileDirName = $envProfileDir }
$targetProfile = Join-Path $targetUserData $profileDirName
$targetLS = Join-Path $targetProfile "Local Storage"

# --- 2. 掃描函式：這個 profile 的 Local Storage 裡有沒有憑證痕跡 ---------------

function Test-CertMarker {
    param([string]$ProfileDir)

    $ls = Join-Path $ProfileDir "Local Storage\leveldb"
    $result = [ordered]@{ Found = $false; Hits = @(); Files = 0; Bytes = 0 }
    if (-not (Test-Path $ls)) { return [pscustomobject]$result }

    # leveldb 的 .ldb 是 Snappy 壓縮的，字串搜尋可能漏；.log（最近寫入）是未壓縮的。
    # 所以「找到」很可靠，「找不到」則要當成「很可能沒有」而不是絕對沒有。
    $files = @(Get-ChildItem -Path (Join-Path $ls "*") -Include *.log, *.ldb -File)
    $result.Files = $files.Count
    $enc = [Text.Encoding]::GetEncoding(28591)   # ISO-8859-1：位元組原樣轉字元，不會亂拋例外
    $hits = @()

    foreach ($f in $files) {
        $result.Bytes += $f.Length
        try {
            $text = $enc.GetString([IO.File]::ReadAllBytes($f.FullName))
        } catch {
            continue   # 被鎖住的檔案跳過，不影響其他檔
        }
        foreach ($m in $markers) {
            if ($text.Contains($m)) { $hits += $m }
        }
    }

    $result.Hits = @($hits | Sort-Object -Unique)
    $result.Found = ($result.Hits.Count -gt 0)
    return [pscustomobject]$result
}

# --- 3. 列出所有候選 profile 與掃描結果 ---------------------------------------

$browsers = @(
    @{ Name = "Chrome"; Exe = "chrome.exe"; UserData = (Join-Path $env:LOCALAPPDATA "Google\Chrome\User Data") },
    @{ Name = "Edge";   Exe = "msedge.exe"; UserData = (Join-Path $env:LOCALAPPDATA "Microsoft\Edge\User Data") }
)

$candidates = @()
foreach ($b in $browsers) {
    if (-not (Test-Path $b.UserData)) { continue }
    $dirs = @(Get-ChildItem -Path $b.UserData -Directory |
        Where-Object { $_.Name -eq "Default" -or $_.Name -like "Profile *" })
    foreach ($d in $dirs) {
        if ([IO.Path]::GetFullPath($d.FullName) -eq $targetProfile) { continue }   # 目標自己不算來源
        $candidates += [pscustomobject]@{
            Browser = $b.Name
            Exe     = $b.Exe
            Name    = $d.Name
            Path    = $d.FullName
            Scan    = (Test-CertMarker $d.FullName)
        }
    }
}

Write-Host ""
Write-Host "目標 profile: $targetProfile" -ForegroundColor Cyan
$targetScan = Test-CertMarker $targetProfile
if ($targetScan.Found) {
    Write-Host "  目標已經有憑證痕跡（$($targetScan.Hits -join ', ')），不需要搬。" -ForegroundColor Green
} else {
    Write-Host "  目標目前沒有憑證痕跡（Local Storage: $($targetScan.Files) 個檔、$($targetScan.Bytes) bytes）"
}

Write-Host ""
Write-Host "掃描來源 profile：" -ForegroundColor Cyan
$withCert = @()
foreach ($c in $candidates) {
    if ($c.Scan.Found) {
        $withCert += $c
        Write-Host ("  [有憑證] {0} / {1}  ({2})" -f $c.Browser, $c.Name, ($c.Scan.Hits -join ', ')) -ForegroundColor Green
        Write-Host ("            $($c.Path)")
    } else {
        Write-Host ("  [  --  ] {0} / {1}" -f $c.Browser, $c.Name) -ForegroundColor DarkGray
    }
}

if ($withCert.Count -eq 0) {
    Write-Host ""
    Write-Host "沒有任何 profile 找到憑證痕跡。" -ForegroundColor Yellow
    Write-Host "  - 最可能的情況：你還沒申請過 tbbstock 的網頁版憑證。"
    Write-Host "    那就直接跑 setup-profile.ps1，在開起來的視窗裡登入 tbbstock、點「憑證申請」完成一次即可。"
    Write-Host "  - 次要可能：憑證藏在壓縮過的 .ldb 裡沒被搜到。"
    Write-Host "    請在平常用的瀏覽器開 tbbstock -> F12 -> Application -> Local Storage"
    Write-Host "    -> https://www.tbbstock.com.tw，看有沒有 TWCACertIdxRef 這個 key。"
}

# --- 4. 沒指定 -From 就停在這裡（純掃描模式） ---------------------------------

if (-not $From) {
    Write-Host ""
    if ($withCert.Count -gt 0) {
        Write-Host "要複製的話，先把來源瀏覽器完全關掉，然後執行：" -ForegroundColor Cyan
        foreach ($c in $withCert) {
            Write-Host ("  .\migrate-cert.ps1 -From `"{0}`"" -f $c.Path)
        }
    }
    Write-Host "（本次只掃描，沒有動到任何檔案。）"
    exit 0
}

# --- 5. 檢查來源與目標 --------------------------------------------------------

$sourceProfile = [IO.Path]::GetFullPath([Environment]::ExpandEnvironmentVariables($From.Trim('"')))
$sourceLS = Join-Path $sourceProfile "Local Storage"

if (-not (Test-Path $sourceLS)) {
    Write-Host ""
    Write-Host "來源找不到 Local Storage 資料夾：" -ForegroundColor Red
    Write-Host "  $sourceLS"
    Write-Host "-From 要給「含 Local Storage 的那一層」，例如 ...\User Data\Default"
    exit 1
}

if ($sourceProfile -eq $targetProfile) {
    Write-Host "來源和目標是同一個資料夾，不用複製。" -ForegroundColor Yellow
    exit 1
}

if (-not (Test-Path $targetProfile)) {
    Write-Host ""
    Write-Host "目標 profile 還不存在：$targetProfile" -ForegroundColor Red
    Write-Host "請先執行 .\setup-profile.ps1 建立資料夾（順便完成 Google 登入）再跑本腳本。"
    exit 1
}

$sourceScan = Test-CertMarker $sourceProfile
if (-not $sourceScan.Found) {
    Write-Host ""
    Write-Host "來源 profile 掃不到憑證痕跡，複製過去大概不會有用：" -ForegroundColor Yellow
    Write-Host "  $sourceProfile"
    if (-not $Force) {
        Write-Host "確定要照做請加上 -Force。"
        exit 1
    }
    Write-Host "-Force：仍然繼續。" -ForegroundColor Yellow
}

# 來源瀏覽器還開著的話，leveldb 會有沒 flush 的資料，複製到的是舊狀態。
# 但如果來源是從別台電腦複製/解壓出來的資料夾，本機的瀏覽器開著沒關係，不用檢查。
$chromeUD = Join-Path $env:LOCALAPPDATA "Google\Chrome\User Data"
$edgeUD = Join-Path $env:LOCALAPPDATA "Microsoft\Edge\User Data"
$sourceExe = $null
if ($sourceProfile.StartsWith($chromeUD, [StringComparison]::OrdinalIgnoreCase)) { $sourceExe = "chrome.exe" }
elseif ($sourceProfile.StartsWith($edgeUD, [StringComparison]::OrdinalIgnoreCase)) { $sourceExe = "msedge.exe" }

if ($sourceExe) {
    # Chrome 常留背景程序，所以留 -Force 逃生口。
    $running = @(Get-CimInstance Win32_Process -Filter "Name='$sourceExe'")
    if ($running.Count -gt 0 -and -not $Force) {
        Write-Host ""
        Write-Host "$sourceExe 還在執行（$($running.Count) 個程序）。" -ForegroundColor Red
        Write-Host "leveldb 有還沒寫入磁碟的資料，現在複製可能拿到不完整的憑證。"
        Write-Host "請把該瀏覽器所有視窗關掉（含背景程序，工作管理員確認）後再執行一次。"
        Write-Host "確認只是背景常駐、沒在使用的話，可以加 -Force 跳過這個檢查。"
        exit 1
    }
} else {
    Write-Host ""
    Write-Host "來源不是本機瀏覽器正在使用的資料夾（看起來是複製/解壓出來的），" -ForegroundColor DarkGray
    Write-Host "跳過「瀏覽器是否執行中」檢查。" -ForegroundColor DarkGray
}

# 目標資料夾被某個 Chrome 開著的話，寫進去會被它蓋回來。
$targetInUse = @(Get-CimInstance Win32_Process -Filter "Name='chrome.exe'" |
    Where-Object { $_.CommandLine -and $_.CommandLine.Contains($targetUserData) })
if ($targetInUse.Count -gt 0) {
    Write-Host ""
    Write-Host "目標資料夾正被 Chrome 使用中，請先關掉那個視窗：" -ForegroundColor Red
    Write-Host "  $targetUserData"
    exit 1
}

# --- 6. 備份目標現有的 Local Storage，然後複製 --------------------------------

Write-Host ""
Write-Host "來源: $sourceLS"
Write-Host "目標: $targetLS"

if (Test-Path $targetLS) {
    $backup = "$targetLS.bak-" + (Get-Date -Format "yyyyMMdd-HHmmss")
    Write-Host "備份目標現有的 Local Storage -> $(Split-Path $backup -Leaf)"
    Move-Item -Path $targetLS -Destination $backup
}

Write-Host "複製中..."
Copy-Item -Path $sourceLS -Destination $targetLS -Recurse -Force

# 上次執行留下的 LOCK 會讓 Chrome 以為資料夾還被佔用，複製過來就順手清掉。
$lock = Join-Path $targetLS "leveldb\LOCK"
if (Test-Path $lock) { Remove-Item $lock -Force -ErrorAction SilentlyContinue }

# --- 7. 驗證 ------------------------------------------------------------------

Write-Host ""
$after = Test-CertMarker $targetProfile
if ($after.Found) {
    Write-Host "完成：目標 profile 現在有憑證痕跡（$($after.Hits -join ', ')）。" -ForegroundColor Green
    Write-Host "  Local Storage: $($after.Files) 個檔、$($after.Bytes) bytes"
    Write-Host ""
    Write-Host "接下來實際驗證：執行 python login.py（或 tbb-login.exe）。" -ForegroundColor Cyan
    Write-Host "如果不再跳出「瀏覽器查無有效數位憑證」，就成功了。"
    Write-Host "來源那張憑證不受影響（這只是複製檔案，沒有經過 RA 重新簽發）。"
} else {
    Write-Host "複製完成，但目標仍掃不到憑證痕跡。" -ForegroundColor Yellow
    Write-Host "可能是來源憑證藏在壓縮的 .ldb 裡（掃不到不代表沒有），還是先跑一次 login.py 看實際結果。"
}
