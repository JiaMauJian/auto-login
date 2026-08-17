"""
把 tbbstock 的持股與淨收付寫回「持股管理」Excel。預設只印變更對照表，不動檔案。

    python update_excel.py                      只看會改什麼（試算，安全）
    python update_excel.py --write               真的寫入（會先自動備份）
    python update_excel.py 1                     只處理 .env 的第 1 組帳號
    python update_excel.py --adopt --today=applied --write

寫入的格子只有三處，其餘全是公式，一律不碰：

    E4:E8  股數    <- 未實現損益的「成交股數」(costqtyn)
    F4:F8  成本    <- 未實現損益的「成交均價」(priceavgn)
    B8     現金餘額 <- 現金基準 + 每日淨收付流水

Excel 的版面完全維持原樣，沒有任何輔助欄位。程式要記的事情（現金基準、
每日淨收付、哪一格是自動哪一格是手動、改動歷程）都放在 Excel 旁邊的紀錄檔裡，
細節見 ledger.py。

自動與手動
----------
每一格各自有狀態。程式記得自己上次寫了什麼，所以 Excel 上跟記憶不符的數字
一定是人改的 —— 這種格子會自動轉成「手動」，之後程式就不再碰它。

要交還給程式管理是 --adopt。持股沒有副作用（網頁庫存就是唯一真相，交還等於
重抄一次）；現金則必須同時說清楚 Excel 上那個數字含不含今天的淨收付：

    --today=applied   B8 現在的數字已經含今天的淨收付了
    --today=pending   還沒含，程式待會兒會幫你加上去

猜錯的後果是多扣或少扣一次，所以程式不猜，沒講就不動現金那一格。
第一次使用（紀錄檔還不存在）也是用 --adopt 把現況接管起來。

分頁怎麼對到帳號
----------------
分頁名稱就是帳戶名（例如「簡嘉懋」），對應 sessionStorage 的 account。
另外還會驗回應裡每一列都帶的 bhno（分公司）與 cseq（客戶號）是否等於這個分頁
登入的 branch_id/cust_id —— 多帳號共用同一個瀏覽器時 JSESSIONID 是共用的，
後登入的會把前一個的 server session 頂掉，這道檢查就是用來擋「抓到別人資料」。
"""

import datetime
import os
import sys
import traceback
from pathlib import Path

from playwright.sync_api import Error as PlaywrightError
from playwright.sync_api import sync_playwright

import excel_io
import ledger
import planner
from fetch import collect
from login import app_dir, configure_browsers_path, load_accounts, open_context, pause, wait_until_finished
from util import pad, show

DEFAULT_EXCEL = Path("dist") / "持股管理-台美股-5家.xls"

STATUS_NAMES = {
    ledger.AUTO: "自動",
    ledger.MANUAL: "手動",
    ledger.UNTRACKED: "未接管",
    planner.WEB_MISSING: "網頁沒有",
}


def excel_path():
    """Excel 檔位置。可用 .env 的 EXCEL_PATH 蓋過，相對路徑以 .env 所在資料夾為基準。"""
    raw = os.getenv("EXCEL_PATH", "").strip().strip('"')
    path = Path(os.path.expandvars(raw)) if raw else DEFAULT_EXCEL
    if not path.is_absolute():
        path = app_dir() / path
    return path


def parse_args(argv):
    """回傳 (要跑第幾組或 None, 是否寫入, 是否交接, 今日淨收付含了沒)。"""
    write = adopt = False
    which = None
    today_included = None

    for arg in argv[1:]:
        if arg == "--write":
            write = True
        elif arg == "--adopt":
            adopt = True
        elif arg == "--today=applied":
            today_included = True
        elif arg == "--today=pending":
            today_included = False
        else:
            try:
                which = int(arg)
            except ValueError:
                print(f"看不懂的參數: {arg}")
                print(__doc__.strip())
                sys.exit(1)

    return which, write, adopt, today_included


def main():
    which, write, adopt, today_included = parse_args(sys.argv)

    path = excel_path()
    if not path.is_file():
        print(f"找不到 Excel 檔: {path}")
        print("可以在 .env 用 EXCEL_PATH 指定位置。")
        sys.exit(1)

    accounts = load_accounts()
    if not accounts:
        print(f"找不到帳號設定。請在 {app_dir()} 放一個 .env 檔（可複製 .env.example）。")
        sys.exit(1)

    numbered = list(enumerate(accounts, start=1))
    if which is not None:
        if not 1 <= which <= len(accounts):
            print(f".env 裡目前有 {len(accounts)} 組帳號，第 {which} 組不存在。")
            sys.exit(1)
        numbered = [numbered[which - 1]]

    today = datetime.date.today()
    print(f"Excel: {path}")
    print(f"今天: {today:%Y/%m/%d}    模式: {'寫入' if write else '試算（不會動檔案）'}")
    if excel_io.is_open_in_excel(path):
        print("偵測到這個檔案正開在 Excel 裡，會直接讀寫你畫面上的那一份。")
    print()

    configure_browsers_path()

    with sync_playwright() as p:
        context, browser = open_context(p)
        try:
            records = collect(context, numbered)
            run(path, records, today, write, adopt, today_included)
        except RuntimeError as exc:
            print(exc)
        finally:
            print()
            wait_until_finished(context)
            try:
                context.close()
                if browser is not None:
                    browser.close()
            except PlaywrightError:
                pass


def run(path, records, today, write, adopt, today_included):
    """開 Excel 算出變更、印出來，必要時寫入。"""
    book_ledger = ledger.Ledger(path)
    if not book_ledger.existed:
        print(f"還沒有紀錄檔，會在第一次寫入時建立: {book_ledger.path.name}")
        print("第一次請加上 --adopt（現金還要 --today=applied 或 --today=pending）把現況接管起來。")
        print()

    excel, workbook, attached = excel_io.open_workbook(path, write)

    pending = []
    adopt_events = []
    blocked = False
    needs_today = False
    at = datetime.datetime.now().isoformat(timespec="seconds")

    try:
        for record in records:
            print("=" * 78)
            head = f"第 {record['order']} 組帳號"
            if record.get("account_code"):
                head += f"    帳號 {record['account_code']}    分頁「{record.get('sheet_name')}」"
            print(head)
            print("=" * 78)

            if record["problems"]:
                blocked = True
                for problem in record["problems"]:
                    print(f"  [中止] {problem}")
                print()
                continue

            sheet, error = excel_io.find_sheet(workbook, record["sheet_name"])
            if sheet is None:
                blocked = True
                print(f"  [中止] {error}")
                print()
                continue

            sheet_data = excel_io.read_sheet(sheet)
            book = book_ledger.sheet(record["sheet_name"])
            book["account_code"] = record["account_code"]

            proposals, warnings = planner.plan(sheet_data, record, book, today)

            if adopt:
                messages, events, missing = planner.adopt(
                    proposals, book, record["sheet_name"], today, today_included, at
                )
                needs_today = needs_today or missing
                adopt_events.extend(events)
                for message in messages:
                    print(f"  [交接] {message}")
                if messages:
                    # 交接改變了「誰在管這一格」，整份提案要重算一次。
                    proposals, warnings = planner.plan(sheet_data, record, book, today)
                    print()

            show_table(proposals)
            for warning in warnings:
                print(f"  [提醒] {warning}")

            pending.append((sheet, proposals, book, record["sheet_name"]))
            print()

        if needs_today:
            print("有現金格要交接，但沒說今天的淨收付算過了沒。")
            print("  Excel 上的 B8 已經含今天的淨收付 -> 加 --today=applied")
            print("  還沒含，要程式幫你加                -> 加 --today=pending")
            print()

        if not write:
            print("以上是試算結果，檔案與紀錄檔都沒有被動過。確認無誤後加上 --write 才會生效。")
            return

        if blocked:
            print("有分頁沒過檢查，整份都不寫入。請先處理上面的 [中止] 項目。")
            return

        write_all(path, book_ledger, pending, adopt_events, today, at, attached)

    finally:
        excel_io.close_workbook(excel, workbook, attached)


def show_table(proposals):
    print(f"  {pad('格子', 6)}{pad('項目', 30)}{pad('Excel', 12)}{pad('網頁', 12)}"
          f"{pad('會變成', 12)}{pad('狀態', 10)}")
    for item in proposals:
        mark = "  <= 會改" if item["will_write"] else ""
        proposed = show(item["proposed"]) if item["will_write"] else "—"
        print(
            f"  {pad(item['cell'], 6)}{pad(item['label'], 30)}"
            f"{pad(show(item['current']), 12)}{pad(show(item['web']), 12)}"
            f"{pad(proposed, 12)}{pad(STATUS_NAMES.get(item['status'], item['status']), 10)}{mark}"
        )
        if item["note"]:
            print(f"  {' ' * 6}{item['note']}")


def write_all(path, book_ledger, pending, adopt_events, today, at, attached):
    """真的寫入：先備份，再寫 Excel，最後才更新紀錄檔與歷程。"""
    writes = [
        (sheet, [(item["row"], item["col"], item["proposed"]) for item in proposals if item["will_write"]])
        for sheet, proposals, _, _ in pending
    ]
    total = sum(len(cells) for _, cells in writes)

    saved = excel_io.backup(path)
    print(f"已備份: {saved}")

    for sheet, cells in writes:
        excel_io.write_cells(sheet, cells)

    # 紀錄檔在 Excel 寫完之後才更新。順序反過來的話，Excel 寫入失敗會留下一份
    # 「以為自己寫過了」的帳本，之後每次比對都判定成人工改動。
    #
    # 接上使用者開著的 Excel 時仍然會更新紀錄檔，即使他最後選擇不存檔。
    # 這不是漏洞：下次執行時 Excel 上的數字跟紀錄檔對不起來，那幾格會被判成手動、
    # 程式停手等人確認 —— 剛好就是「你自己撤銷了程式的修改」該有的結果。
    events = list(adopt_events)
    for _, proposals, book, sheet_name in pending:
        events.extend(planner.commit(proposals, book, sheet_name, today, at))

    book_ledger.save()
    book_ledger.append_history(events)

    print(f"已更新紀錄檔: {book_ledger.path.name}（歷程 {len(events)} 筆）")

    if total == 0:
        print("Excel 沒有任何格子需要改。")
    elif attached:
        # 刻意不代為存檔：讓使用者在自己的畫面上看過再決定。
        print()
        print(f"已把 {total} 格寫進你開著的 Excel，但還沒存檔：")
        print("  滿意   -> 切到 Excel 按 Ctrl+S")
        print("  不滿意 -> 關閉 Excel 時選「不要儲存」，就會回到存檔前的版本")
        print("  （提醒：程式做的變更沒辦法用 Ctrl+Z 復原，只能靠不存檔或上面那份備份）")
    else:
        print(f"已寫入 {total} 格: {path}")


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
