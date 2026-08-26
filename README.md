# auto-login / 持股同步

自動登入券商網站 + Playwright 抓即時資料 + COM 寫回 Excel「持股管理」檔。
Tkinter GUI 是唯一的操作介面。詳細架構與規則見 [CLAUDE.md](CLAUDE.md)。

## 安裝

沒有虛擬環境設定，直接在專案資料夾用 PowerShell（或一般終端機）執行：

```powershell
python -m pip install -r requirements.txt
python -m pip install pywin32
python -m playwright install chromium
```

- `requirements.txt` 裝的是 `playwright`（抓網頁資料）、`python-dotenv`（讀 `.env`）、
  `ttkbootstrap`（GUI）。
- `pywin32` 是 COM 寫 Excel 用的，`requirements.txt` 沒列但 [excel_io.py](excel_io.py)
  會用到，要另外裝。
- `python -m playwright install chromium` 是額外下載瀏覽器執行檔，第一次一定要跑，否則
  Playwright 會找不到瀏覽器。

一律用 `python -m` 呼叫（而不是直接打 `pip install` / `playwright install`），原因有兩個：
- 確保裝到的套件跟等一下執行 `python ui.py` 用的是同一個 Python（機器上如果同時裝了
  多個 Python 版本，兩者可能不一致，導致「明明裝了卻 ModuleNotFoundError」）。
- `pip`、`playwright` 這些指令稿在 PATH 沒設好時會抓不到（`CommandNotFoundException`），
  但 `python` 通常找得到，`python -m` 繞過 PATH 問題直接用該 Python 執行對應模組。

## 設定

複製 `.env.example` 為 `.env`，填入帳密、Excel 路徑等設定：

```powershell
copy .env.example .env
```

## 執行

```powershell
python ui.py
```
