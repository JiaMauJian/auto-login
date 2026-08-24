"""
在持股管理 Excel 補上模擬用的分頁（交易人A、交易人B…），測完再一鍵移除。

    python dev_tools/sim_excel.py                 只看會做什麼（試算，安全）
    python dev_tools/sim_excel.py --write          真的加分頁
    python dev_tools/sim_excel.py --count 5        只加 5 個
    python dev_tools/sim_excel.py --remove --write 把模擬分頁與它們在紀錄檔裡的資料清掉

分頁是整張複製現有那個真分頁，所以版面、格式、公式（那張表的公式全部只參照
自己這一頁，複製不會跑掉）都跟真的一模一樣，只改這幾格：

    D4:D8  股票名稱(代號)  <- 這個假帳號的 5 檔
    E4:G8  股數/成本/現金股利 <- 照 simulate.FIXED_ACCOUNTS 那個帳號的固定數字填，
                             不是隨機也不是清成 0，跟假網頁看到的數字一致
    I4:I8  股價           <- 同樣照 FIXED_ACCOUNTS 填，市值才不會是 0
    B6:B8  原始資金/今年股本/現金餘額 <- 固定的起始資金與現金（見 simulate.FIXED_ACCOUNTS）

為什麼要能一鍵移除
------------------
模擬分頁是加在正在用的那份 Excel 裡，紀錄檔（同步紀錄.json / 同步歷程.jsonl）
是跟著 Excel 檔名走的，所以假分頁的資料會跟真分頁躺在同一份紀錄檔裡。
測完手工清會清不乾淨（尤其是歷程那種一行一筆的檔案），所以清理這件事要程式做。
歷程檔是只增不改的檔案，移除模擬紀錄前會先留一份原檔（見 remove_sheets）。
"""

import json
import shutil
import sys
import traceback
from pathlib import Path

# 這支腳本可能被 login.py 用 `from dev_tools import sim_excel` 匯入，
# 也可能直接用 `python dev_tools/sim_excel.py` 執行；兩種情況 sys.path 起點不一樣，
# 這裡把專案根目錄補回去，root 底下那些模組（excel_io / login / util）才找得到。
_ROOT = Path(__file__).resolve().parent.parent
if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))

import excel_io
import ledger as ledger_mod
from dev_tools import simulate
from excel_io import CELL_BALANCE, COL_COST, COL_NAME, COL_QTY, HOLDING_ROWS, excel_path
from login import pause
from util import pad

COL_DIVIDEND = 7     # G 現金股利
COL_PRICE = 9        # I 股價
ROW_CAPITAL = 6      # B6 原始資金
ROW_YEAR = 7         # B7 今年股本

DEFAULT_COUNT = 19


def parse_args(argv):
    """回傳 (要幾個, 是否寫入, 是否移除)。"""
    write = remove = False
    count = None

    rest = list(argv[1:])
    while rest:
        arg = rest.pop(0)
        if arg == "--write":
            write = True
        elif arg == "--remove":
            remove = True
        elif arg == "--count":
            if not rest:
                print("--count 後面要接數字")
                sys.exit(1)
            count = rest.pop(0)
        elif arg.startswith("--count="):
            count = arg.split("=", 1)[1]
        else:
            print(f"看不懂的參數: {arg}")
            print(__doc__.strip())
            sys.exit(1)

    if count is None:
        # .env 設了 SIMULATE_ACCOUNTS 就以它為準 —— 分頁數跟假帳號數對不上的話，
        # 讀取時會有帳號找不到自己的分頁。
        number = simulate.simulate_count() or DEFAULT_COUNT
    else:
        try:
            number = int(count)
        except ValueError:
            print(f"--count 要是數字，收到的是: {count}")
            sys.exit(1)

    if not 1 <= number <= len(simulate.FAKE_NAMES):
        print(f"數量要在 1 ~ {len(simulate.FAKE_NAMES)} 之間。")
        sys.exit(1)

    return number, write, remove


def main():
    count, write, remove = parse_args(sys.argv)

    path = excel_path()
    if path is None:
        print("還沒指定要同步哪一份 Excel。")
        print("請在 .env 用 EXCEL_PATH 指定，或開介面（ui.py）按「開啟EXCEL」選一次。")
        sys.exit(1)
    if not path.is_file():
        print(f"找不到 Excel 檔: {path}")
        print("可以在 .env 用 EXCEL_PATH 指定位置。")
        sys.exit(1)

    # 加/刪分頁是動檔案結構，接上使用者開著的那份會跟他的畫面打架，所以直接擋掉。
    if excel_io.is_open_in_excel(path):
        print(f"這個檔案正開在 Excel 裡：{path}")
        print("加或刪分頁要動檔案結構，請先在 Excel 存檔並關閉，再執行一次。")
        sys.exit(1)

    print(f"Excel: {path}")
    print(f"模式: {'寫入' if write else '試算（不會動檔案）'}")
    print()

    excel, workbook, attached = excel_io.open_workbook(path, write)
    try:
        if remove:
            remove_sheets(path, workbook, write)
        else:
            add_sheets(path, workbook, count, write)

        if write:
            workbook.Save()
            print(f"已存檔: {path}")
    finally:
        excel_io.close_workbook(excel, workbook, attached)


def add_sheets(path, workbook, count, write):
    """複製出模擬分頁。已經存在的就跳過，不覆蓋任何東西。"""
    wanted = simulate.FAKE_NAMES[:count]
    existing = {sheet.Name.strip() for sheet in workbook.Worksheets}

    template = None
    for sheet in workbook.Worksheets:
        if sheet.Name.strip() not in simulate.FAKE_NAMES:
            template = sheet
            break
    if template is None:
        print("找不到可以當範本的真分頁（這份 Excel 裡全部都是模擬分頁）。")
        sys.exit(1)

    todo = [name for name in wanted if name not in existing]
    print(f"範本分頁: 「{template.Name}」（整張複製，只改持股與現金那幾格）")
    print(f"要新增 {len(todo)} 個分頁" + (f"，已存在跳過 {len(wanted) - len(todo)} 個" if len(todo) < len(wanted) else ""))
    print()

    print(f"  {pad('分頁', 12)}{pad('現金餘額', 14)}{'持股'}")
    for name in todo:
        account = simulate.FIXED_ACCOUNTS.get(name, {})
        cash = account.get("cash", 0)
        stocks = "、".join(f"{stock_name}({code})" for code, stock_name, _qty, _price, _mkt
                           in account.get("holdings", []))
        print(f"  {pad(name, 12)}{pad(f'{cash:,}', 14)}{stocks}")
    print()

    if not todo:
        print("沒有要新增的分頁。")
        return

    if not write:
        print("以上是試算結果，檔案沒有被動過。確認無誤後加上 --write 才會生效。")
        return

    for name in todo:
        # Copy 的參數一定要用位置傳（Before, After），不能寫 Copy(After=...)：
        # win32com 的動態呼叫吃不到這個具名參數，會被當成「兩個都沒給」，
        # 而 Worksheet.Copy 沒給位置時的行為是「複製到一個新的工作簿」——
        # 分頁一個都沒多，然後下面那行取到的是範本自己，接著就被改名。
        # （踩過一次：範本分頁被連續改名 19 次，最後整份 Excel 只剩一張分頁。）
        before = workbook.Worksheets.Count
        template.Copy(None, workbook.Worksheets(before))

        if workbook.Worksheets.Count != before + 1:
            raise RuntimeError(
                f"複製分頁沒有生效（分頁數還是 {before}）。檔案沒有存檔，"
                f"原檔沒有被動過。"
            )

        sheet = workbook.Worksheets(workbook.Worksheets.Count)
        sheet.Name = name
        fill_sheet(sheet, name)
        print(f"  已新增分頁「{name}」")

    print()
    print(f"共新增 {len(todo)} 個分頁。這些分頁已經填好 FIXED_ACCOUNTS 的固定股數與成本，")
    print("直接開持股同步介面按一次同步就會照網頁值覆蓋、現金基準也會在第一次登入時設定好：")
    print("    tbb-login.exe（或 python ui.py）")


def fill_sheet(sheet, name):
    """把複製出來的分頁改成這個假帳號的樣子。只碰那幾格，版面完全不動。

    資料全部來自 simulate.FIXED_ACCOUNTS[name]，跟假網頁（simulate.seed_data）
    看到的股數/成本/市價/現金一致，不是隨機生成也不是清成 0。
    """
    account = simulate.FIXED_ACCOUNTS.get(name, {})
    capital = account.get("original_capital", 0)
    cash = account.get("cash", 0)

    for row, (code, stock_name, qty, price, mkt) in zip(HOLDING_ROWS, account.get("holdings", [])):
        sheet.Cells(row, COL_NAME).Value = f"{stock_name}({code})"
        sheet.Cells(row, COL_QTY).Value = qty
        sheet.Cells(row, COL_COST).Value = price
        sheet.Cells(row, COL_DIVIDEND).Value = 0
        sheet.Cells(row, COL_PRICE).Value = mkt

    sheet.Cells(ROW_CAPITAL, 2).Value = capital
    sheet.Cells(ROW_YEAR, 2).Value = capital
    sheet.Cells(*CELL_BALANCE).Value = cash


def remove_sheets(path, workbook, write):
    """把模擬分頁與它們在紀錄檔、歷程裡的資料全部清掉。"""
    names = [sheet.Name.strip() for sheet in workbook.Worksheets
             if sheet.Name.strip() in simulate.FAKE_NAMES]

    ledger_path = path.parent / (path.stem + ledger_mod.LEDGER_SUFFIX)
    history_path = path.parent / (path.stem + ledger_mod.HISTORY_SUFFIX)

    in_ledger = []
    if ledger_path.is_file():
        data = json.loads(ledger_path.read_text(encoding="utf-8"))
        in_ledger = [key for key in (data.get("sheets") or {}) if key in simulate.FAKE_NAMES]

    history_lines = 0
    if history_path.is_file():
        for line in history_path.read_text(encoding="utf-8").splitlines():
            if line.strip() and json.loads(line).get("sheet") in simulate.FAKE_NAMES:
                history_lines += 1

    print(f"Excel 裡的模擬分頁: {len(names)} 個" + (f"（{'、'.join(names)}）" if names else ""))
    print(f"紀錄檔裡的模擬分頁: {len(in_ledger)} 個 -> {ledger_path.name}")
    print(f"歷程裡的模擬紀錄  : {history_lines} 行 -> {history_path.name}")
    print()

    if not (names or in_ledger or history_lines):
        print("沒有模擬資料需要清理。")
        return

    if not write:
        print("以上是試算結果，什麼都沒有被動過。確認無誤後加上 --remove --write 才會生效。")
        return

    if names:
        for name in names:
            workbook.Worksheets(name).Delete()
            print(f"  已刪除分頁「{name}」")

    if in_ledger:
        data = json.loads(ledger_path.read_text(encoding="utf-8"))
        for key in in_ledger:
            data["sheets"].pop(key, None)
        temp = ledger_path.with_name(ledger_path.name + ".tmp")
        temp.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
        temp.replace(ledger_path)
        print(f"  已從紀錄檔移除 {len(in_ledger)} 個分頁")

    if history_lines:
        # 歷程本來是只增不改的檔案，這裡是唯一的例外，所以動手前先留一份原檔。
        folder = path.parent / "備份"
        folder.mkdir(exist_ok=True)
        kept = folder / (history_path.stem + "-移除模擬前" + history_path.suffix)
        shutil.copy2(history_path, kept)

        lines = [line for line in history_path.read_text(encoding="utf-8").splitlines()
                 if line.strip() and json.loads(line).get("sheet") not in simulate.FAKE_NAMES]
        history_path.write_text("".join(line + "\n" for line in lines), encoding="utf-8")
        print(f"  已從歷程移除 {history_lines} 行（原檔備份在 {kept.name}）")

    print()
    print("清理完成。記得把 .env 的 SIMULATE_ACCOUNTS 註解掉或設成 0。")


if __name__ == "__main__":
    try:
        main()
    except SystemExit as exc:
        if exc.code:
            pause("按 Enter 關閉視窗...")
        raise
    except Exception:
        traceback.print_exc()
        pause("發生未預期的錯誤，按 Enter 關閉視窗...")
        sys.exit(1)
