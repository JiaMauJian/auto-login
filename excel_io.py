"""
Excel 讀寫。只認得持股管理表的這幾格，其餘全是公式，一律不碰：

    D4:D8  股票名稱(代號)   只讀，用來對出每一列是哪一檔
    E4:E8  股數            <- 未實現損益的「成交股數」
    F4:F8  成本            <- 未實現損益的「成交均價」
    B8     現金餘額         <- 由紀錄檔的現金流水算出來

位置寫死是刻意的：這支程式只認得這個特定格式的持股管理表，
認錯了就是把數字寫進別人的格子裡，寧可一開始就對不上而報錯。

版面完全維持原樣，沒有任何輔助欄位 —— 前日餘額、每日淨收付、最後更新日、
哪一格是自動哪一格是手動，全部記在 ledger.py 管理的紀錄檔裡。
"""

import datetime
import os
import re
import shutil
from pathlib import Path

from login import app_dir
from util import to_num

DEFAULT_EXCEL = Path("dist") / "持股管理-台美股-5家.xls"

HOLDING_ROWS = range(4, 9)
COL_NAME, COL_QTY, COL_COST = 4, 5, 6
CELL_BALANCE = (8, 2)

BACKUP_KEEP = 10

# 從「台灣50(0050)」取出 0050。全形括號也吃，因為是人手打的欄位。
CODE_PATTERN = re.compile(r"[（(]\s*([0-9A-Za-z]+)\s*[)）]")


def excel_path():
    """Excel 檔位置。可用 .env 的 EXCEL_PATH 蓋過，相對路徑以 .env 所在資料夾為基準。"""
    raw = os.getenv("EXCEL_PATH", "").strip().strip('"')
    path = Path(os.path.expandvars(raw)) if raw else DEFAULT_EXCEL
    if not path.is_absolute():
        path = app_dir() / path
    return path


def stock_code_of(text):
    """「台灣50(0050)」-> 「0050」。沒有括號就回 None（空白列）。"""
    found = CODE_PATTERN.search(str(text or ""))
    return found.group(1).strip().upper() if found else None


def read_sheet(sheet):
    """
    把一個分頁讀成純 Python 資料，之後的計算都不必再碰 COM。

    這樣做不只是為了乾淨：COM 每讀一格都是一次跨行程呼叫，而且介面版之後
    要在背景執行緒跑，讓計算完全脫離 COM 物件會少掉很多麻煩。
    """
    rows = []
    for row in HOLDING_ROWS:
        label = sheet.Cells(row, COL_NAME).Value
        code = stock_code_of(label)
        if not code:
            continue                 # 空白列，這張表允許留空
        rows.append({
            "row": row,
            "label": str(label).strip(),
            "code": code,
            "qty": sheet.Cells(row, COL_QTY).Value,
            "cost": sheet.Cells(row, COL_COST).Value,
        })

    return {
        "balance": to_num(sheet.Cells(*CELL_BALANCE).Value, None),
        "rows": rows,
    }


def write_cells(sheet, writes):
    """writes 是 (row, col, value) 的清單。只會寫進記憶體，存檔是另一回事。"""
    for row, col, value in writes:
        sheet.Cells(row, col).Value = value


def find_sheet(book, name):
    """用分頁名稱找分頁。同名不只一個就不猜，直接報錯。"""
    wanted = (name or "").strip()
    if not wanted:
        return None, "沒有分頁名稱可以對應（登入後沒讀到帳戶名）"
    matches = [s for s in book.Worksheets if s.Name.strip() == wanted]
    if not matches:
        return None, f"找不到名稱為「{wanted}」的分頁"
    if len(matches) > 1:
        return None, f"有 {len(matches)} 個分頁都叫「{wanted}」，無法判斷要寫哪一個"
    return matches[0], None


def is_open_in_excel(path):
    """檔案是不是正被 Excel 開著。"""
    if (path.parent / ("~$" + path.name)).exists():
        return True
    try:
        with open(path, "r+b"):
            return False
    except OSError:
        return True


def open_workbook(path, write):
    """
    取得 Workbook，回傳 (excel, book, attached)。

    檔案已經開在 Excel 裡時就直接接上那個實例，不必為了跑這支程式先關檔。
    接上的好處不只是方便：讀到的是你畫面上的即時內容（含還沒存檔的修改），
    寫入也只進到記憶體，你看過覺得對再自己 Ctrl+S，等於多一道人工確認。

    attached=True 時絕對不能 Close 或 Quit —— 那是使用者的視窗，關掉他會很錯愕。
    """
    import win32com.client

    if is_open_in_excel(path):
        try:
            book = win32com.client.GetObject(str(path))
            if Path(str(book.FullName)).resolve() == path.resolve():
                return book.Application, book, True
        except Exception:
            pass
        raise RuntimeError(
            f"這個檔案正被 Excel 開著，但接不上那個 Excel：{path}\n"
            f"請先在 Excel 存檔並關閉，再執行一次。"
        )

    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    return excel, excel.Workbooks.Open(str(path), ReadOnly=not write), False


def close_workbook(excel, book, attached):
    """接上使用者自己開的 Excel 時，關檔與結束都不是我們該做的事。"""
    if attached:
        return
    book.Close(False)
    excel.Quit()


def backup(path):
    """寫入前先備份，只留最近幾份。"""
    folder = path.parent / "備份"
    folder.mkdir(exist_ok=True)
    dest = folder / f"{path.stem}_{datetime.datetime.now():%Y%m%d_%H%M%S}{path.suffix}"
    shutil.copy2(path, dest)

    old_files = sorted(folder.glob(f"{path.stem}_*{path.suffix}"))
    for old in old_files[:-BACKUP_KEEP]:
        old.unlink()
    return dest
