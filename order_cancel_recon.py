"""
唯讀偵察腳本：把「委託查詢」頁（order/layoutRWD.jsp?type=2）取消掛單那條路
從頭到尾看一遍——勾選、按「終止委託單」、把刪單確認視窗裡的每一列倒出來，
**到這裡就停手，絕對不按視窗裡的「確認」**，掛在外面的委託一張都不會被刪掉。

定位跟 order_recon.py 一樣（偵察腳本，不是正式版），只是它連 AJAX 都不打：
刪單走的是頁面本身那條路（憑證簽章在瀏覽器裡做，見 docs/介面規劃.md 10.3
第二點），所以要看的東西全在 DOM 跟頁面自己的 JS 裡。

要解開的是 10.3 第十點那四件事，一次跑完就都看得到：

    1. 逐列 checkbox 的 selector，以及怎麼從那一列對回委託書號
       （「取消全部買單／賣單」要逐列勾，只有全選的 #numall 不夠）
    2. 全選（check_all(this,'c')）會不會勾出重複的列
       ——2026/08/31 的刪單確認截圖裡，同一筆 J0124 出現了兩列
    3. layer 標題列的字是不是真的「刪單確認」（等視窗要靠它）
    4. 預約單（ordstatus=1）在不在這一頁

前提：**要有一張真的還掛在外面的委託**（有效數量 > 0），不然按「終止委託單」
只會跳「請先勾選」之類的訊息，什麼都看不到。

用法：
    python order_cancel_recon.py <第幾組帳號>

輸出在 偵察資料\\ 資料夾（已加進 .gitignore）：一份摘要、一份頁面原始 HTML、
一份頁面自己的 JS（check_all／#openConfirm 的 handler／renderTable）。
"""

import json
import sys
from datetime import datetime

from playwright.sync_api import Error as PlaywrightError, TimeoutError as PlaywrightTimeoutError
from playwright.sync_api import sync_playwright

from login import (app_dir, configure_browsers_path, do_login, load_accounts, open_context,
                   pause, wait_until_finished)
from order_recon import ORDER_PAGE, SESSION_JS
from recon import OUTPUT_DIR_NAME, pick_accounts

# 刪單確認視窗裡那顆真的會送出去的按鈕。這支腳本從頭到尾不碰它，寫在這裡只是
# 為了讓「不要點這個」變成看得見的一行——改這支腳本的人請維持這條界線。
NEVER_CLICK = "#submit"

# 頁面上所有表格、所有列、所有 checkbox 一次倒出來。分兩件事看：
#   一、同一筆委託是不是在頁面上出現兩次（RWD 頁面常常桌機一份、手機一份），
#       那正是 10.3 第七點那個「一筆變兩列」最像的成因；
#   二、每一列的 checkbox 到底長什麼樣、身上帶不帶得回委託書號。
# offsetParent 用來分辨「畫面上真的看得到的那一份」——兩份都在 DOM 裡的話，
# 隱藏的那一份照樣會被 check_all 勾到，但人看不到。
DUMP_TABLES_JS = """
() => {
    const dumpBox = (b) => ({
        id: b.id, name: b.name, cls: b.className, value: b.value, checked: b.checked,
        onclick: b.getAttribute('onclick'),
        attrs: [...b.attributes].map(a => a.name + '=' + a.value),
        visible: b.offsetParent !== null,
    });
    return [...document.querySelectorAll('table')].map((table, ti) => ({
        index: ti,
        id: table.id, cls: table.className,
        visible: table.offsetParent !== null,
        rows: [...table.querySelectorAll('tr')].map((tr, ri) => ({
            index: ri,
            cls: tr.className, id: tr.id,
            visible: tr.offsetParent !== null,
            cells: [...tr.querySelectorAll('th,td')].map(td => td.innerText.trim().replace(/\\s+/g, ' ')),
            boxes: [...tr.querySelectorAll('input[type=checkbox]')].map(dumpBox),
        })),
    }));
}
"""

# 頁面自己的程式：全選怎麼勾、「終止委託單」按下去做了什麼、那張表怎麼畫的。
# 這幾支就是「勾選」跟「送出去的那一批」中間隔的那一層（10.3 第七點），讀它
# 比在 DOM 上猜可靠。jQuery 的事件處理器要從 jQuery._data 拿，不是 DOM 屬性。
DUMP_PAGE_JS_JS = """
() => {
    const out = {};
    out.check_all = (typeof check_all === 'function') ? check_all.toString() : '(沒有這個函式)';
    out.renderTable = (typeof renderTable === 'function') ? renderTable.toString() : '(沒有這個函式)';
    out.mod = (typeof mod === 'undefined') ? '(undefined)' : String(mod);
    out.orderArray = (typeof orderArray === 'undefined') ? '(undefined)' : JSON.stringify(orderArray);
    const btn = document.getElementById('openConfirm');
    out.openConfirm_html = btn ? btn.outerHTML : '(找不到 #openConfirm)';
    try {
        const events = (btn && window.jQuery) ? jQuery._data(btn, 'events') : null;
        out.openConfirm_handlers = events
            ? Object.keys(events).map(k => k + ': ' + events[k].map(e => e.handler.toString()).join('\\n'))
            : ['(抓不到 jQuery 事件)'];
    } catch (e) {
        out.openConfirm_handlers = ['(讀事件時出錯: ' + e + ')'];
    }
    out.inline_scripts = [...document.querySelectorAll('script')].filter(s => !s.src)
        .map(s => s.textContent);
    out.script_src = [...document.querySelectorAll('script')].filter(s => s.src).map(s => s.src);
    return out;
}
"""

# 全選之後實際被勾起來的是哪幾個。value 通常就是那一列的識別（是不是委託書號
# 正是要看的事），visible 分得出「畫面上看得到的那一份」跟藏起來的另一份。
DUMP_CHECKED_JS = """
() => [...document.querySelectorAll('input[type=checkbox]')]
        .filter(b => b.checked)
        .map(b => ({id: b.id, name: b.name, cls: b.className, value: b.value,
                    visible: b.offsetParent !== null,
                    row: b.closest('tr') ? b.closest('tr').innerText.trim().replace(/\\s+/g, ' ') : ''}))
"""

# 刪單確認視窗（orderConfirmRWD.html）真正要送出去的那一批就是 parent.orderArray
# ——「勾了幾列」跟「會送幾筆」是不是同一個數字，答案在這裡。
DUMP_DIALOG_JS = """
() => {
    const out = {mod: (typeof mod === 'undefined') ? '(undefined)' : String(mod)};
    out.orderArray = (typeof orderArray === 'undefined') ? '(undefined)' : JSON.stringify(orderArray);
    out.layer_titles = [...document.querySelectorAll('.layui-layer-title')].map(t => t.innerText.trim());
    const frame = document.querySelector("iframe[src*='orderConfirmRWD']");
    out.iframe_src = frame ? frame.getAttribute('src') : '(找不到 orderConfirmRWD 的 iframe)';
    return out;
}
"""


def _save(out_dir, name, text, report):
    (out_dir / name).write_text(text, encoding="utf-8")
    report.append(f"已存檔: {name}")


def recon_one(page, out_dir, stamp, tag, report):
    """在一個已登入的分頁上把取消掛單那條路看一遍。不按確認。"""
    page.goto(ORDER_PAGE, wait_until="domcontentloaded")
    # 這一頁的表格是 renderTable() 打完 queryOrder 才畫出來的，goto 回來的當下
    # 通常還是空的。等「頁面上出現任何一個 checkbox」就夠——沒有掛在外面的委託
    # 時本來就不會有，所以等不到也只是印一句，不是壞掉。
    try:
        page.wait_for_function(
            "() => document.querySelectorAll('input[type=checkbox]').length > 0", timeout=15000)
    except (PlaywrightError, PlaywrightTimeoutError):
        report.append("等不到任何 checkbox——這個帳號今天可能沒有還掛在外面的委託。")

    session = page.evaluate(SESSION_JS)
    report.append(f"帳號代碼 = 1{session.get('branch_id')}-{session.get('cust_id')}"
                  f"   帳戶名 = {session.get('account')}")

    _save(out_dir, f"{stamp}_{tag}_委託查詢頁.html", page.content(), report)

    page_js = page.evaluate(DUMP_PAGE_JS_JS)
    _save(out_dir, f"{stamp}_{tag}_委託查詢頁_JS.txt", "\n\n".join([
        "=== check_all ===", page_js["check_all"],
        "=== #openConfirm 這顆按鈕 ===", page_js["openConfirm_html"],
        "=== #openConfirm 的事件處理器 ===", "\n".join(page_js["openConfirm_handlers"]),
        "=== renderTable ===", page_js["renderTable"],
        "=== 外部 JS ===", "\n".join(page_js["script_src"]),
        "=== 內嵌 JS ===", "\n\n---\n\n".join(page_js["inline_scripts"]),
    ]), report)

    tables = page.evaluate(DUMP_TABLES_JS)
    _save(out_dir, f"{stamp}_{tag}_表格與checkbox.json",
          json.dumps(tables, ensure_ascii=False, indent=2), report)

    report.append("")
    report.append("--- 頁面上有幾份表格（同一筆委託出現兩次的話，在這裡就看得出來）---")
    for table in tables:
        boxes = sum(len(row["boxes"]) for row in table["rows"])
        report.append(f"  表格[{table['index']}] id={table['id'] or '-'} class={table['cls'] or '-'}"
                      f" 看得到={table['visible']} 列數={len(table['rows'])} checkbox={boxes}")
        for row in table["rows"]:
            if not row["boxes"]:
                continue
            box = row["boxes"][0]
            report.append(f"      列[{row['index']}] 看得到={row['visible']}"
                          f" checkbox: id={box['id'] or '-'} name={box['name'] or '-'}"
                          f" class={box['cls'] or '-'} value={box['value'] or '-'}")
            report.append(f"          屬性: {box['attrs']}")
            report.append(f"          這一列: {' | '.join(row['cells'])[:200]}")

    # ---- 全選 ----
    report.append("")
    report.append("--- 按下全選（#numall，onclick=check_all(this,'c')）---")
    try:
        page.evaluate("() => { const b = document.getElementById('numall'); if (b) { b.click(); } }")
        page.wait_for_timeout(300)
    except PlaywrightError as exc:
        report.append(f"勾全選失敗：{exc}")

    checked = page.evaluate(DUMP_CHECKED_JS)
    report.append(f"勾起來的 checkbox 共 {len(checked)} 個"
                  f"（其中畫面上看得到的 {sum(1 for c in checked if c['visible'])} 個）：")
    for box in checked:
        report.append(f"  id={box['id'] or '-'} name={box['name'] or '-'} class={box['cls'] or '-'}"
                      f" value={box['value'] or '-'} 看得到={box['visible']}")
        report.append(f"      {box['row'][:200]}")

    # ---- 開刪單確認視窗（不按確認）----
    report.append("")
    report.append("--- 按「終止委託單」（#openConfirm）開視窗，不按確認 ---")
    try:
        page.click("#openConfirm")
    except PlaywrightError as exc:
        report.append(f"按不下去：{exc}")
        return

    try:
        page.locator(".layui-layer-title").first.wait_for(state="visible", timeout=10000)
    except (PlaywrightError, PlaywrightTimeoutError):
        report.append("等不到 layer 視窗——可能跳的是「請先勾選」之類的提示，看瀏覽器畫面。")

    page.wait_for_timeout(800)   # 讓 iframe 裡的表格畫完
    dialog = page.evaluate(DUMP_DIALOG_JS)
    report.append(f"layer 標題列的字: {dialog['layer_titles']}"
                  f"   （10.3 等視窗就是靠這個字，應該是「刪單確認」）")
    report.append(f"parent.mod = {dialog['mod']}   （刪單應該是 '3'）")
    report.append(f"iframe: {dialog['iframe_src']}")
    _save(out_dir, f"{stamp}_{tag}_orderArray.json", dialog["orderArray"], report)

    try:
        order_array = json.loads(dialog["orderArray"])
    except (TypeError, ValueError):
        order_array = None
    if isinstance(order_array, list):
        report.append(f"orderArray 有 {len(order_array)} 筆（＝按下確認會送幾次刪單）：")
        seen = {}
        for i, item in enumerate(order_array):
            ordno = item.get("ordno")
            seen[ordno] = seen.get(ordno, 0) + 1
            report.append(f"  [{i}] 委託書號={ordno} 股票={item.get('stockId')}"
                          f" 買賣={item.get('buysell')} 原委託={item.get('orgqty')}"
                          f" 已取消={item.get('celqty')} 已成交={item.get('matqty')}"
                          f" ordstatus={item.get('ordstatus')}（1預約單/2盤中單）")
        dup = {k: v for k, v in seen.items() if v > 1}
        report.append(f"  *** 重複的委託書號: {dup if dup else '沒有，一筆一列'} ***")

    frame = page.frame_locator("iframe[src*='orderConfirmRWD']")
    try:
        report.append("")
        report.append("--- 視窗裡那張表（人看到的樣子）---")
        report.append(frame.locator("#aTable").inner_text())
    except PlaywrightError as exc:
        report.append(f"讀不到視窗裡的表格：{exc}")

    # ---- 關掉視窗。從頭到尾沒有碰過 NEVER_CLICK 那顆。----
    try:
        frame.locator("#cancel").click()
        page.wait_for_timeout(500)
        report.append("")
        report.append("已按視窗裡的「取消」關掉——沒有送出任何刪單。")
    except PlaywrightError as exc:
        report.append(f"關視窗失敗（自己去瀏覽器按「取消」）：{exc}")


def main():
    accounts = load_accounts()
    if not accounts:
        print(f"找不到帳號設定。請在 {app_dir()} 放一個 .env 檔。")
        sys.exit(1)

    selected = pick_accounts(accounts, sys.argv)

    out_dir = app_dir() / OUTPUT_DIR_NAME
    out_dir.mkdir(exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M")

    configure_browsers_path()

    report = [f"偵察時間: {datetime.now():%Y/%m/%d %H:%M:%S}（取消掛單那條路，唯讀，不按確認）"]

    with sync_playwright() as p:
        context, browser = open_context(p)
        spare_page = context.pages[0] if context.pages else None

        for index, account in selected:
            report.append("")
            report.append("=" * 70)
            report.append(f"第 {index} 組帳號")
            report.append("=" * 70)
            try:
                page = do_login(context, account["id"], account["password"], spare_page)
                spare_page = None
                recon_one(page, out_dir, stamp, f"第{index}組", report)
            except PlaywrightTimeoutError:
                report.append("登入逾時，找不到欄位，網站版面可能已變更。")
            except PlaywrightError as exc:
                report.append(f"瀏覽器操作失敗：{exc}")
            except Exception as exc:
                # 這裡吞掉的是「這一組出事」，不是整支腳本出事——後面還有帳號要跑，
                # 而且不管跑到哪裡，下面那份摘要都一定要寫得出來。
                report.append(f"發生錯誤：{exc}")

        report_path = out_dir / f"{stamp}_取消掛單偵察_摘要.txt"
        report_path.write_text("\n".join(report), encoding="utf-8")

        print("\n".join(report))
        print()
        print(f"原始資料與摘要都在: {out_dir}")
        print("這支腳本沒有按過刪單確認視窗裡的「確認」，掛在外面的委託一張都沒動。")

        wait_until_finished(context)
        try:
            context.close()
            if browser is not None:
                browser.close()
        except PlaywrightError:
            pass


if __name__ == "__main__":
    try:
        main()
    except SystemExit as exc:
        if exc.code:
            pause("按 Enter 關閉視窗...")
        raise
