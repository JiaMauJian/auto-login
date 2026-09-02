"""
唯讀偵察腳本：把下單頁那幾個下拉選單有哪些選項全部倒出來，回答「零股要怎麼下」。

為什麼需要這一支
----------------
*** 2026/09/01 09:18 已經問到答案了，這支的任務完成 ***
報告：偵察資料60901_0918_下單表單選項.txt。結論：

    交易盤別（.tab1）：'1' 整股、'2' 盤後、'5' 盤中零股、'3' 盤後零股
    零股的數量欄旁邊寫「股」（整股寫「張」）
    零股的委託別只剩 'R'（ROD）、交易別只剩 '0'（現股）

order_fill.TAB1_ODD 因此填 "5"（取盤中零股的理由見 docs/介面規劃.md 9.4）。
這支留著不是待辦，是「網站改版之後再倒一次來對照」用的——下面那四個問題就是
當初要問的東西，答案現在都寫在 order_fill.py 的常數旁邊。

當初要回答的四件事（現在都有答案了，留著當對照）：

    1. `.tab1`（交易盤別）有哪些選項，盤中零股是哪一個值
    2. 換成零股之後 `#bsFlag`（委託別）剩哪些——盤中零股一般只收 ROD，
       如果真的只剩 ROD，那 orders.BS_FLAG_INTRADAY 那條 IOC 的路對零股不適用。
       這一項只影響「出清・零股」：買賣股票整批固定 ROD（規劃文件明講），
       不管零股還剩哪些選項都一樣
    3. `#qty` 的單位是張還是股（整股是「張」，零股一定是「股」，
       填錯就是差 1000 倍，而且不會報錯）。程式目前就是照「零股填股」寫的
       （order_fill._check_qty 零股只收 1~999），這一項是去確認那個假設
    4. 換成零股之後表單還有沒有別的欄位跟著變

安全設計
--------
全程只做「登入 → 開下單頁 → 讀下拉選單的選項 → 換選項再讀一次」，
**不填股票、不填數量、不填價格，也絕對不點「確認下單」**（#openConfirm1 從頭到尾
沒有被碰到）。換 `.tab1` 會觸發頁面自己那段把股票清空的 handler（見
order_fill.open_order_form 的說明），但本來就沒填股票，沒有東西可以被清掉。

用法：`python recon_order_form.py [第幾組帳號]`（預設第 1 組）。
輸出在 偵察資料\\ 資料夾。
"""

import sys
from datetime import datetime

from playwright.sync_api import Error as PlaywrightError, TimeoutError as PlaywrightTimeoutError
from playwright.sync_api import sync_playwright

from login import (
    app_dir,
    configure_browsers_path,
    do_login,
    load_accounts,
    open_context,
    pause,
    wait_until_finished,
)
from order_fill import ORDER_ENTRY_PAGE

OUTPUT_DIR_NAME = "偵察資料"

# 要倒出來的下拉選單。名字是給人看的，選擇器是頁面上的。
SELECTS = {
    "交易盤別（.tab1）": ".tab1",
    "交易別（#tradeType）": "#tradeType",
    "委託別（#bsFlag）": "#bsFlag",
    "價格類別（#priceRadio）": "#priceRadio",
}

DUMP_JS = """
(selector) => {
    const el = document.querySelector(selector);
    if (!el) return null;
    return {
        value: el.value,
        disabled: el.disabled,
        options: Array.from(el.options).map(o => ({
            value: o.value, text: (o.text || '').trim(), disabled: o.disabled,
        })),
    };
}
"""

# 數量欄的單位只能從它旁邊的文字看出來（張／股），欄位本身看不出來。
QTY_JS = """
() => {
    const el = document.querySelector('#qty');
    if (!el) return null;
    const box = el.closest('div, td, li, tr') || el.parentElement;
    return {
        placeholder: el.placeholder || '',
        title: el.title || '',
        value: el.value,
        maxlength: el.getAttribute('maxlength'),
        nearby: (box ? box.innerText : '').replace(/\\s+/g, ' ').trim().slice(0, 200),
    };
}
"""


def dump_form(page):
    """把目前表單狀態讀成幾行字。只讀，不動任何欄位。"""
    lines = []
    for title, selector in SELECTS.items():
        info = page.evaluate(DUMP_JS, selector)
        if info is None:
            lines.append(f"  {title}: 找不到這個元素")
            continue
        state = "（disabled）" if info["disabled"] else ""
        lines.append(f"  {title}{state} 目前選 {info['value']!r}")
        for option in info["options"]:
            mark = "  <- 目前選這個" if option["value"] == info["value"] else ""
            skip = "（disabled）" if option["disabled"] else ""
            lines.append(f"      {option['value']!r:>8} = {option['text']}{skip}{mark}")

    qty = page.evaluate(QTY_JS)
    if qty is None:
        lines.append("  數量欄（#qty）: 找不到")
    else:
        lines.append(f"  數量欄（#qty） placeholder={qty['placeholder']!r} "
                     f"title={qty['title']!r} maxlength={qty['maxlength']!r}")
        lines.append(f"      旁邊的文字: {qty['nearby']}")
    return lines


def main():
    accounts = load_accounts()
    if not accounts:
        print(f"找不到帳號設定。請在 {app_dir()} 放一個 .env 檔（可複製 .env.example）。")
        sys.exit(1)

    which = 1
    if len(sys.argv) >= 2:
        try:
            which = int(sys.argv[1])
        except ValueError:
            print(f"參數要是數字（第幾組帳號），收到的是: {sys.argv[1]}")
            sys.exit(1)
    if not 1 <= which <= len(accounts):
        print(f".env 裡目前有 {len(accounts)} 組帳號，第 {which} 組不存在。")
        sys.exit(1)
    account = accounts[which - 1]

    out_dir = app_dir() / OUTPUT_DIR_NAME
    out_dir.mkdir(exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M")

    configure_browsers_path()

    report = [
        f"偵察時間: {datetime.now():%Y/%m/%d %H:%M:%S}",
        f"第 {which} 組帳號",
        "目的：下單頁的交易盤別／委託別有哪些選項，零股要選哪個值。",
        "本次只讀下拉選單、只換選項，不填任何欄位、不點「確認下單」。",
    ]

    with sync_playwright() as p:
        context, browser = open_context(p)
        spare = context.pages[0] if context.pages else None

        try:
            page = do_login(context, account["id"], account["password"], spare)
        except PlaywrightTimeoutError:
            report.append("登入逾時，找不到欄位，網站版面可能已變更。")
            print("\n".join(report))
            sys.exit(1)
        except PlaywrightError as exc:
            report.append(f"瀏覽器操作失敗：{exc}")
            print("\n".join(report))
            sys.exit(1)

        report.append(f"登入完成，目前頁面: {page.url}")

        try:
            page.goto(ORDER_ENTRY_PAGE, wait_until="domcontentloaded")
            page.wait_for_selector(".tab1", timeout=15000)

            report.append("")
            report.append("=" * 70)
            report.append("開下單頁之後的原始狀態（什麼都還沒動）")
            report.append("=" * 70)
            report.extend(dump_form(page))

            # 逐一切換交易盤別，看每一種底下的委託別剩哪些。這是這支腳本真正要
            # 回答的問題：零股那一種能不能用 IOC，還是只剩 ROD。
            tab_info = page.evaluate(DUMP_JS, ".tab1")
            for option in (tab_info or {}).get("options", []):
                if option["disabled"]:
                    continue
                report.append("")
                report.append("=" * 70)
                report.append(f"把交易盤別切成 {option['value']!r} = {option['text']} 之後")
                report.append("=" * 70)
                try:
                    page.select_option(".tab1", option["value"])
                    # 頁面自己有 change handler 要跑（會清空股票欄），等它一下
                    page.wait_for_timeout(500)
                    report.extend(dump_form(page))
                except PlaywrightError as exc:
                    report.append(f"  切不過去：{exc}")
        except PlaywrightError as exc:
            report.append(f"讀取下單頁失敗：{exc}")

        report_path = out_dir / f"{stamp}_下單表單選項.txt"
        report_path.write_text("\n".join(report), encoding="utf-8")

        print("\n".join(report))
        print()
        print(f"報告已存檔: {report_path}")
        print("瀏覽器留著，看完自己關掉。")

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
