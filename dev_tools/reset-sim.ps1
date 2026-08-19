# 把模擬帳號（交易人A~S）的 19 個分頁重設回 simulate.py 的 FIXED_ACCOUNTS 原始數字。
#
# 用法（在專案根目錄或 dev_tools\ 底下執行都可以）：
#   .\dev_tools\reset-sim.ps1
#
# 依序做三件事，每一步都是真的寫入（--write）：
#   1. sim_excel.py --remove --write   刪掉 19 個模擬分頁，連同紀錄檔/歷程檔裡的資料
#   2. sim_excel.py --write            重新複製出 19 個分頁，內容是 FIXED_ACCOUNTS 的固定數字
#   3. update_excel.py --adopt --today=pending --write
#                                      把 20 個分頁（含真帳號）一次接管進紀錄檔
#
# 用途：模擬測試跑一輪之後，想把 dist\持股管理-1真人19模擬.xls 的模擬帳號部分
# 重設回乾淨的起始狀態，重來一輪測試。真帳號分頁不受影響。
#
# 執行前請先把這份 Excel 關閉——sim_excel.py 會擋開著的檔案（動的是分頁結構，
# 跟開著的 Excel 打架）。每一步都會自動備份 Excel 與歷程檔，但清掉的紀錄檔/
# 歷程還是回不去，先確認目前的模擬測試進度不需要保留再執行。

$ErrorActionPreference = "Stop"

# 這支腳本住在 dev_tools\ 底下，但要在專案根目錄執行三支 python 腳本。
$ProjectRoot = Split-Path $PSScriptRoot -Parent

Write-Host "即將把模擬帳號（交易人A~S）重設回 FIXED_ACCOUNTS 的原始數字。" -ForegroundColor Cyan
Write-Host "這 19 個帳號目前的同步紀錄與歷程會被清掉（會先自動備份，但清掉的部分回不去）。" -ForegroundColor Yellow
$answer = Read-Host "確定要繼續嗎？(y/N)"
if ($answer -ne "y" -and $answer -ne "Y") {
    Write-Host "已取消，沒有動到任何檔案。"
    exit 0
}

Push-Location $ProjectRoot
try {
    Write-Host ""
    Write-Host "== 1/3 移除現有模擬分頁 ==" -ForegroundColor Cyan
    python dev_tools\sim_excel.py --remove --write
    if ($LASTEXITCODE -ne 0) { throw "sim_excel.py --remove --write 失敗（結束碼 $LASTEXITCODE）" }

    Write-Host ""
    Write-Host "== 2/3 重新產生模擬分頁（FIXED_ACCOUNTS 原始數字）==" -ForegroundColor Cyan
    python dev_tools\sim_excel.py --write
    if ($LASTEXITCODE -ne 0) { throw "sim_excel.py --write 失敗（結束碼 $LASTEXITCODE）" }

    Write-Host ""
    Write-Host "== 3/3 接管全部分頁（含真帳號）==" -ForegroundColor Cyan
    python update_excel.py --adopt --today=pending --write
    if ($LASTEXITCODE -ne 0) { throw "update_excel.py --adopt --write 失敗（結束碼 $LASTEXITCODE）" }
}
finally {
    Pop-Location
}

Write-Host ""
Write-Host "完成：模擬帳號已重設回 FIXED_ACCOUNTS 的原始數字。" -ForegroundColor Green
