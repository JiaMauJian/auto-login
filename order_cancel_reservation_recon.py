"""
唯讀偵察腳本：把「預約查詢」頁（order/layoutRWD.jsp?type=4）的刪除機制看一遍。

跟 order_cancel_recon.py 的關係：那支看的是「委託查詢」（type=2）取消掛單那條路，
這支看的是同一件事在「預約查詢」（type=4）上長什麼樣子——從螢幕截圖看，這一頁
每一列自己就有一顆「刪除」鈕，不是勾選＋一顆批次「終止委託單」，跟 type=2 那套
UI 形狀不一樣，目前完全沒有程式碰過。

**這支從頭到尾不點「刪除」，也不點任何 checkbox。** 只把 DOM 跟頁面自己的內嵌
JS 原封不動存下來、印出來看，用讀原始碼取代用猜的——「刪除」鈕的 onclick 可能
是直接送出、也可能跳確認視窗，點下去之前得先知道是哪一種，不然萬一是前者，
一點就把預約單刪了，不可逆。等看過這支印出來的東西、確認安全的下一步之後，
才輪到下一支腳本去試著真的按（比照 order_cancel_recon.py 那樣，按到跳確認視窗
就停手）。

要解開的問題：
    1. 這一頁的表格／checkbox／刪除鈕的 selector 長什麼樣，怎麼從一列對回
       委託書號或預約書號（螢幕上看到的是 P0638918 這種格式，不是 ordno）。
    2. 「刪除」鈕的 onclick 呼叫的是哪個函式、那個函式做了什麼（直接 AJAX？
       跳確認視窗？要不要簽章？）。
    3. 有沒有全選／批次刪除的機制，還是真的一列一次。
    4. 這一頁的資料來源是哪支 CMD（不一定是 queryOrder，畢竟 order_query.py
       目前查到的 queryOrder 回應裡，預約單那幾列的 ordno 是空的）。

前提：要有一張真的還在外面的預約單（有效交易日是下一個交易日那種），不然這頁
可能整個是空的，看不到刪除鈕長什麼樣。

用法：
    python order_cancel_reservation_recon.py <第幾組帳號>

輸出在 偵察資料\\ 資料夾（已加進 .gitignore）：一份摘要、一份頁面原始 HTML、
一份頁面自己的內嵌 JS。
"""

import json
import sys
from datetime import datetime

from playwright.sync_api import Error as PlaywrightError, TimeoutError as PlaywrightTimeoutError
from playwright.sync_api import sync_playwright

from login import (app_dir, configure_browsers_path, do_login, load_accounts, open_context,
                   pause, wait_until_finished)
from order_recon import RESERVE_PAGE, SESSION_JS
from recon import OUTPUT_DIR_NAME, pick_accounts

# 刪單確認視窗裡那顆真的會送出去的按鈕。這支腳本從頭到尾不碰它——理由跟
# order_cancel_recon.py 的 NEVER_CLICK 一樣：開視窗只是把 orderArray 準備好、
# 呼叫 layer.open()，沒有任何 AJAX；只有這顆按鈕才會真的送出刪單。
NEVER_CLICK = "#submit"

DUMP_DIALOG_JS = """
() => {
    const out = {mod: (typeof mod === 'undefined') ? '(undefined)' : String(mod)};
    out.orderArray = (typeof orderArray === 'undefined') ? '(undefined)' : JSON.stringify(orderArray);
    out.layer_titles = [...document.querySelectorAll('.layui-layer-title')].map(t => t.innerText.trim());
    const frame = document.querySelector("iframe[src*='orderConfirm']");
    out.iframe_src = frame ? frame.getAttribute('src') : '(找不到 orderConfirm 開頭的 iframe)';
    return out;
}
"""

# 表格、列、checkbox、按鈕（含 onclick 原始碼）一次倒出來。按鈕不點，只讀
# onclick 屬性的文字——onclick="delOrder(...)" 這種寫法，看屬性就夠，不用執行。
DUMP_TABLES_JS = """
() => {
    const dumpEl = (el) => ({
        tag: el.tagName, id: el.id, name: el.name, cls: el.className,
        value: el.value !== undefined ? el.value : null,
        text: el.innerText ? el.innerText.trim().replace(/\\s+/g, ' ') : '',
        onclick: el.getAttribute('onclick'),
        attrs: [...el.attributes].map(a => a.name + '=' + a.value),
        visible: el.offsetParent !== null,
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
            boxes: [...tr.querySelectorAll('input[type=checkbox]')].map(dumpEl),
            buttons: [...tr.querySelectorAll('button, a[onclick], input[type=button]')].map(dumpEl),
        })),
    }));
}
"""

# 這一頁所有內嵌 <script>，以及所有看起來跟刪除／取消有關的全域函式原始碼
# （名字含 del/cancel/giveup/remove，不分大小寫，把可能相關的都撈出來——
# 這一頁的函式名字目前完全沒人看過，用列舉關鍵字取代用猜的）。
DUMP_PAGE_JS_JS = """
() => {
    const out = {};
    out.inline_scripts = [...document.querySelectorAll('script')].filter(s => !s.src)
        .map(s => s.textContent);
    out.script_src = [...document.querySelectorAll('script')].filter(s => s.src).map(s => s.src);
    const names = Object.getOwnPropertyNames(window).filter(
        n => /del|cancel|giveup|remove/i.test(n) && typeof window[n] === 'function');
    out.candidate_functions = names.map(n => n + ':\\n' + window[n].toString());
    return out;
}
"""


def _save(out_dir, name, text, report):
    (out_dir / name).write_text(text, encoding="utf-8")
    report.append(f"已存檔: {name}")


def recon_one(page, out_dir, stamp, tag, report):
    """在一個已登入的分頁上把預約查詢那一頁看一遍。不點任何刪除鈕、不點任何 checkbox。"""
    page.goto(RESERVE_PAGE, wait_until="domcontentloaded")
    try:
        page.wait_for_function(
            "() => document.querySelectorAll('table tr').length > 1", timeout=15000)
    except (PlaywrightError, PlaywrightTimeoutError):
        report.append("等不到表格畫出東西來——這個帳號可能今天沒有預約單。")

    session = page.evaluate(SESSION_JS)
    report.append(f"帳號代碼 = 1{session.get('branch_id')}-{session.get('cust_id')}"
                  f"   帳戶名 = {session.get('account')}")

    _save(out_dir, f"{stamp}_{tag}_預約查詢頁.html", page.content(), report)

    page_js = page.evaluate(DUMP_PAGE_JS_JS)
    _save(out_dir, f"{stamp}_{tag}_預約查詢頁_JS.txt", "\n\n".join([
        "=== 名字看起來跟刪除／取消有關的全域函式 ===",
        "\n\n---\n\n".join(page_js["candidate_functions"]) or "(一個都沒找到)",
        "=== 外部 JS ===", "\n".join(page_js["script_src"]),
        "=== 內嵌 JS ===", "\n\n---\n\n".join(page_js["inline_scripts"]),
    ]), report)

    tables = page.evaluate(DUMP_TABLES_JS)
    _save(out_dir, f"{stamp}_{tag}_表格與按鈕.json",
          json.dumps(tables, ensure_ascii=False, indent=2), report)

    report.append("")
    report.append("--- 頁面上有幾份表格、每一列有沒有 checkbox／按鈕 ---")
    for table in tables:
        boxes = sum(len(row["boxes"]) for row in table["rows"])
        buttons = sum(len(row["buttons"]) for row in table["rows"])
        report.append(f"  表格[{table['index']}] id={table['id'] or '-'} class={table['cls'] or '-'}"
                      f" 看得到={table['visible']} 列數={len(table['rows'])}"
                      f" checkbox={boxes} 按鈕={buttons}")
        for row in table["rows"]:
            if not row["boxes"] and not row["buttons"]:
                continue
            report.append(f"      列[{row['index']}] 看得到={row['visible']}"
                          f"　這一列: {' | '.join(row['cells'])[:200]}")
            for box in row["boxes"]:
                report.append(f"          checkbox: id={box['id'] or '-'} name={box['name'] or '-'}"
                              f" value={box['value'] or '-'} onclick={box['onclick'] or '-'}")
            for btn in row["buttons"]:
                report.append(f"          按鈕: <{btn['tag']}> text={btn['text'] or '-'}"
                              f" id={btn['id'] or '-'} onclick={btn['onclick'] or '-'}")

    report.append("")
    report.append("*** 表格與按鈕的部分到此為止沒有點過任何東西。"
                  "接下來會真的點一下第一列的「刪除」鈕看跳出來的視窗，"
                  "但視窗裡的「確認」／#submit 從頭到尾不會點。***")

    # ---- 點第一列的「刪除」鈕，看跳出來的視窗，不按確認 ----
    row_button = page.locator("#qOrderTable .delRow").first
    if row_button.count() == 0:
        report.append("")
        report.append("這個帳號現在沒有可以點「刪除」的預約單列，這一段跳過。")
        return

    report.append("")
    report.append("--- 點第一列的「刪除」鈕（.delRow），不按視窗裡的確認 ---")
    try:
        row_button.click()
    except PlaywrightError as exc:
        report.append(f"點不下去：{exc}")
        return

    try:
        page.locator(".layui-layer-title").first.wait_for(state="visible", timeout=10000)
    except (PlaywrightError, PlaywrightTimeoutError):
        report.append("等不到 layer 視窗——看瀏覽器畫面發生了什麼。")

    page.wait_for_timeout(800)
    dialog = page.evaluate(DUMP_DIALOG_JS)
    report.append(f"layer 標題列的字: {dialog['layer_titles']}")
    report.append(f"parent.mod = {dialog['mod']}   （單筆刪除應該是 '3'）")
    report.append(f"iframe src = {dialog['iframe_src']}")
    _save(out_dir, f"{stamp}_{tag}_delRow_orderArray.json", dialog["orderArray"], report)

    try:
        order_array = json.loads(dialog["orderArray"])
    except (TypeError, ValueError):
        order_array = None
    if isinstance(order_array, list):
        report.append(f"orderArray 有 {len(order_array)} 筆：")
        for i, item in enumerate(order_array):
            report.append(f"  [{i}] {json.dumps(item, ensure_ascii=False)}")

    frame = page.frame_locator("iframe[src*='orderConfirm']")
    for probe_id in ("#aTable", "#submit", "#cancel", "#result0"):
        try:
            count = frame.locator(probe_id).count()
            text = frame.locator(probe_id).inner_text() if count and probe_id == "#aTable" else ""
        except PlaywrightError as exc:
            count, text = f"讀取失敗: {exc}", ""
        report.append(f"  視窗 iframe 裡 {probe_id}：count={count}"
                      + (f"　內容={text[:300]}" if text else ""))

    try:
        frame.locator("#cancel").click()
        page.wait_for_timeout(500)
        report.append("")
        report.append("已按視窗裡的「取消」關掉——沒有送出任何刪單，這筆預約單一動都沒動。")
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

    report = [f"偵察時間: {datetime.now():%Y/%m/%d %H:%M:%S}（預約查詢頁，唯讀，不點刪除）"]

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
                report.append(f"發生錯誤：{exc}")

        report_path = out_dir / f"{stamp}_預約查詢偵察_摘要.txt"
        report_path.write_text("\n".join(report), encoding="utf-8")

        print("\n".join(report))
        print()
        print(f"原始資料與摘要都在: {out_dir}")
        print("這支腳本沒有點過任何刪除鈕，預約單一張都沒動。")

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
