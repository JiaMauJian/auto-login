"""
Excel 讀寫。只認得持股管理表的這幾格，其餘全是公式，一律不碰：

    D4:D8  股票名稱(代號)   只讀，用來對出每一列是哪一檔
    E4:E8  股數            <- 未實現損益的「成交股數」
    F4:F8  成本            <- 未實現損益的「成交均價」
    B8     現金餘額         <- 由紀錄檔的現金流水算出來
    D1     程式維護提醒      每次寫入順便刷新，關閉程式時清掉，見 marker_enabled

位置寫死是刻意的：這支程式只認得這個特定格式的持股管理表，
認錯了就是把數字寫進別人的格子裡，寧可一開始就對不上而報錯。

版面完全維持原樣，沒有任何輔助欄位 —— 前日餘額、每日淨收付、最後更新日，
全部記在 ledger.py 管理的紀錄檔裡。
"""

import codecs
import datetime
import os
import re
import shutil
import time
from pathlib import Path

from login import app_dir
from util import env_int, to_num

ENV_FILE = ".env"
ENV_KEY = "EXCEL_PATH"

HOLDING_ROWS = range(4, 9)
COL_NAME, COL_QTY, COL_COST = 4, 5, 6
CELL_BALANCE = (8, 2)

# D1 沒被上面三處佔用，理論上可以安全借來提醒「這份檔案有程式在管」。
# 語意是「這份檔案由程式維護」的持久標記，不是「程式現在正在跑」的即時燈號——
# _write_worker 每次同步只是短暫用 COM 開檔、寫入、存檔、關檔，平常大部分時間
# 程式根本沒碰著檔案，做不出真正的即時狀態，文字用詞也刻意避開「控制中」這種
# 容易被誤會成即時狀態的說法。
CELL_MARKER = (1, 4)
MARKER_TEXT = "此檔案由程式自動維護，手動修改可能被覆蓋"
MARKER_COLOR = 0xFF0000  # COM 的 Font.Color 是 BGR，這個值是藍字
MARKER_ENV_KEY = "EXCEL_CONTROL_MARKER"


def marker_enabled():
    """.env 的 EXCEL_CONTROL_MARKER 沒設或設 1 就開著；設 0 就整個關掉這個功能。"""
    return env_int(MARKER_ENV_KEY, 1) != 0

BACKUP_KEEP = 10

# Excel 實例在半路死掉時會看到的 HRESULT。
#   0x800A01A8  物件不見了（Open 成功之後，下一句就撲空）
#   0x80010108  用戶端中斷了已啟動物件的連線
DEAD_OBJECT_CODES = (-2146827864, -2147417848)

DEAD_EXCEL_HINT = (
    "Excel 開起來之後又自己消失了。\n\n"
    "這通常是 Office 沒有啟用：啟用檢查失敗時 Excel 會把自己收掉、"
    "改用「產品啟動失敗」模式重開，程式手上的物件就跟著失效。\n\n"
    "兩個辦法：\n"
    "  1. 先自己用 Excel 把這個檔開著再跑 —— 程式會接上你那個視窗，不會另外開一個。\n"
    "  2. 把 Office 啟用起來，從根本解決。"
)


def is_dead_object(exc):
    """這個例外是不是「Excel 死在半路」。是的話錯誤訊息要多說一句，不然沒人看得懂。"""
    code = getattr(exc, "hresult", None)
    if code is None:
        args = getattr(exc, "args", ())
        code = args[0] if args and isinstance(args[0], int) else None
    return code in DEAD_OBJECT_CODES

# 從「台灣50(0050)」取出 0050。全形括號也吃，因為是人手打的欄位。
CODE_PATTERN = re.compile(r"[（(]\s*([0-9A-Za-z]+)\s*[)）]")


def excel_path():
    """
    要同步哪一份 Excel。只看 .env 的 EXCEL_PATH，沒設就回傳 None（還沒選過）。

    刻意不給預設檔名。給了預設的話，換一台機器或改個檔名，程式會安靜地去同步
    「另一個檔」或報「找不到某某檔」，而不是直接說「你還沒選」—— 對一支會改寫
    檔案的程式來說，猜錯對象是最不能接受的錯。

    相對路徑以 .env 所在資料夾為基準（打包成 exe 後就是 exe 旁邊）。
    """
    raw = os.getenv(ENV_KEY, "").strip().strip('"')
    if not raw:
        return None
    path = Path(os.path.expandvars(raw))
    return path if path.is_absolute() else app_dir() / path


def remember_excel_path(path):
    """
    把選好的檔案寫回 .env 的 EXCEL_PATH，下次開啟就直接用，不必再選一次。

    只動那一行：整個檔讀進來、找到就換掉、沒有就補在最後，其他行原封不動 ——
    .env 裡都是人寫的設定與註解，任何「重新產生一份」的做法都會把它們洗掉。
    BOM 也要留著：用記事本存過的 .env 檔頭會有 BOM，拿掉之後下次讀取
    第一個設定會失效（load_dotenv 是用 utf-8-sig 讀的）。
    """
    env = app_dir() / ENV_FILE
    raw = env.read_bytes() if env.is_file() else b""
    has_bom = raw.startswith(codecs.BOM_UTF8)
    text = raw.decode("utf-8-sig") if raw else ""
    newline = "\r\n" if "\r\n" in text else "\n"

    lines = text.splitlines()
    entry = f"{ENV_KEY}={path}"
    for i, line in enumerate(lines):
        stripped = line.strip()
        if stripped.startswith("#") or "=" not in stripped:
            continue
        if stripped.split("=", 1)[0].strip() == ENV_KEY:
            lines[i] = entry
            break
    else:
        if lines and lines[-1].strip():
            lines.append("")
        lines.append("# 要同步哪一份 Excel。介面上按「開啟EXCEL」選檔時會自動改這一行。")
        lines.append(entry)

    body = (newline.join(lines) + newline).encode("utf-8")
    env.write_bytes((codecs.BOM_UTF8 if has_bom else b"") + body)

    # 這個行程的 os.environ 也要跟著更新，否則同一次執行裡再呼叫 excel_path()
    # 拿到的還是舊的（load_dotenv 只在 import 時跑過一次）。
    os.environ[ENV_KEY] = str(path)


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
    """
    writes 是 (row, col, value) 的清單。只會寫進記憶體，存檔是另一回事。

    現金餘額那格 value 會是「=1000-107」這種公式字串（見 util.cash_formula），
    要走 Formula 屬性才會被 Excel 當成公式算，走 Value 屬性只會存成一串文字。
    """
    for row, col, value in writes:
        cell = sheet.Cells(row, col)
        if isinstance(value, str) and value.startswith("="):
            cell.Formula = value
        else:
            cell.Value = value


def write_marker(sheet):
    """在 D1 補上藍字提醒。每次同步寫入都順手刷新，反正已經在開檔，不多一次 COM 往返。"""
    cell = sheet.Cells(*CELL_MARKER)
    cell.Value = MARKER_TEXT
    cell.Font.Color = MARKER_COLOR


def clear_all_markers(path):
    """
    程式正常關閉時呼叫：把 Excel 裡所有分頁的 D1 提醒清掉。

    範圍是整份活頁簿，不只這次執行寫過的分頁——帳戶數不多，全部掃一遍成本低，
    還能順便清掉上次非正常結束（當機、被工作管理員砍掉）遺留下來的提醒。

    Best-effort：任何一步失敗都默默放棄，不能卡住關閉流程。真正的風險是相反的
    不對稱——非正常結束不會走到這裡，D1 會留著過期的提醒，那個風險是可以接受的，
    但絕不能因為清除失敗就讓程式關不掉。
    """
    if not marker_enabled() or path is None or not path.is_file():
        return
    try:
        backup(path)
        excel, workbook, attached = open_workbook(path, True)
    except Exception:
        return
    try:
        changed = False
        for sheet in workbook.Worksheets:
            cell = sheet.Cells(*CELL_MARKER)
            if str(cell.Value or "").strip() == MARKER_TEXT:
                cell.Value = None
                changed = True
        if changed:
            workbook.Save()
    except Exception:
        pass
    finally:
        close_workbook(excel, workbook, attached)


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


def open_workbook(path, write, attempts=3):
    """
    取得 Workbook，回傳 (excel, book, attached)。

    檔案已經開在 Excel 裡時就直接接上那個實例，不必為了跑這支程式先關檔。
    接上的好處不只是方便：讀到的是你畫面上的即時內容（含還沒存檔的修改）。
    寫入之後一樣會存檔 —— 留給人自己按 Ctrl+S 聽起來像多一道確認，實際上是
    「紀錄檔記成寫過了、檔案卻沒存」這個破口的來源。反悔靠寫入前那份備份。

    attached=True 時絕對不能 Close 或 Quit —— 那是使用者的視窗，關掉他會很錯愕。

    為什麼開完要再碰一下、失敗還要重開
    ----------------------------------
    Office 沒啟用時，Excel 起來之後那道啟用檢查失敗會把自己收掉重來，
    Open 明明成功，下一句碰 Worksheets 就是 OLE error 0x800A01A8。
    錯誤發生在半路最麻煩 —— 那時備份做了、可能還寫了一半。所以開完先碰一下
    確認物件是活的，死的就整個丟掉、換一個新的 Excel 再試。
    """
    trouble = ""
    for attempt in range(attempts):
        excel = book = None
        attached = False
        try:
            excel, book, attached = _open_once(path, write)
        except Exception as exc:
            trouble = str(exc)
        else:
            if _still_alive(book):
                return excel, book, attached
            trouble = DEAD_EXCEL_HINT
            close_workbook(excel, book, attached)
        if attempt < attempts - 1:
            time.sleep(1.0)

    raise RuntimeError(f"開不了 Excel：{path}\n\n{trouble}")


def _still_alive(book):
    """碰一下 COM 物件，確認它還在。死掉的物件在這裡就會丟例外。"""
    try:
        book.Worksheets.Count
        return True
    except Exception:
        return False


def _open_once(path, write):
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
    """
    接上使用者自己開的 Excel 時，關檔與結束都不是我們該做的事。

    這裡的錯誤一律吞掉。收尾失敗沒有任何補救動作可做，而 Excel 已經死掉時
    Close/Quit 只會再丟一個例外，把真正的原因蓋在「During handling of the
    above exception」底下 —— 那正是最需要看清楚的那一行。
    """
    if attached:
        return
    for action in (lambda: book.Close(False), excel.Quit):
        try:
            action()
        except Exception:
            pass


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
