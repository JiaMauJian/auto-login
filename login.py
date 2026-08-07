"""
自動登入 tbbstock.com.tw（支援多組帳號）
流程：自動開啟瀏覽器 -> 依 .env 內的每組帳號各開一個獨立 context -> 自動填入身分證、密碼與驗證碼
     -> 使用者確認後，一次把所有已開啟的帳號都送出登入

多組帳號設定：.env 用 TBB_ID_1/TBB_PASSWORD_1、TBB_ID_2/TBB_PASSWORD_2... 依序編號（見 .env.example）。
每組帳號開在同一個瀏覽器視窗的不同分頁（共用同一個 browser context，因此也共用同一組 cookie/session）。
注意：這個網站是 Java/Servlet 架構，登入狀態通常是用 JSESSIONID 這類 cookie 辨識；
共用 context 代表多帳號同時登入時，有可能其中一個分頁登入後把另一個分頁的 session 頂掉。
如果實測發現會互踢，需要改回每組帳號各自獨立 context（browser.new_context()）。

驗證碼取得方式：頁面載入過程中瀏覽器會呼叫 VerifyNumberServlet，
該次請求的回應內容本身就是明文數字（例如 "86176"），網站再用這組數字畫成 canvas 圖案顯示。
因此直接攔截這次請求的回應即可取得驗證碼，不需要做圖片辨識。
"""

import os
import sys

from dotenv import load_dotenv
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError

load_dotenv()

LOGIN_URL = "https://www.tbbstock.com.tw/tbb/index/home.jsp"


def load_accounts():
    accounts = []
    i = 1
    while True:
        tbb_id = os.getenv(f"TBB_ID_{i}")
        tbb_password = os.getenv(f"TBB_PASSWORD_{i}")
        if not tbb_id or not tbb_password:
            break
        accounts.append({"id": tbb_id, "password": tbb_password})
        i += 1
    return accounts


def prepare_login(context, tbb_id, tbb_password):
    """開一個新分頁，填入帳密與驗證碼，回傳該分頁物件。"""
    page = context.new_page()

    verify_number = {}

    def on_verify_response(resp):
        if "VerifyNumberServlet" in resp.url:
            try:
                verify_number["value"] = resp.text().strip()
            except Exception:
                pass

    page.on("response", on_verify_response)
    page.goto(LOGIN_URL)

    # 頁面預設是「帳號登入」模式，欄位 placeholder 為「帳號」；
    # 必須先點「身份證登入」，網站的 JS 才會把同一個欄位切換成「身分證」模式。
    mode_link = page.locator("a:visible", has_text="身份證登入").first
    mode_link.wait_for(state="visible", timeout=15000)
    mode_link.click()

    # 頁面上有兩個 id="id" 的重複欄位，用父層範圍鎖定可見的那個。
    id_input = page.locator("#ind-tab1 #id")
    id_input.wait_for(state="visible", timeout=15000)
    page.wait_for_function(
        "document.querySelector('#ind-tab1 #id').placeholder === '身分證'",
        timeout=5000,
    )
    id_input.fill(tbb_id)

    page.locator("#pass").fill(tbb_password)

    # 等驗證碼圖案畫出來，確認頁面已經準備好
    page.locator("#VerifyNumber canvas").wait_for(state="visible", timeout=15000)

    # 等待攔截到 VerifyNumberServlet 的回應（回應內容就是明文驗證碼）
    for _ in range(30):
        if "value" in verify_number:
            break
        page.wait_for_timeout(100)

    if "value" not in verify_number:
        print(f"[{tbb_id}] 沒有攔截到 VerifyNumberServlet 的回應，請確認網站是否改版，改回手動輸入驗證碼。")
    else:
        page.locator("#NumberLabel").fill(verify_number["value"])
        print(f"[{tbb_id}] 已自動取得並填入驗證碼: {verify_number['value']}")

    return page


def main():
    accounts = load_accounts()
    if not accounts:
        print("請先複製 .env.example 為 .env，並依 TBB_ID_1/TBB_PASSWORD_1、TBB_ID_2/TBB_PASSWORD_2... 填入至少一組帳號。")
        sys.exit(1)

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        context = browser.new_context()

        pages = []
        try:
            for account in accounts:
                page = prepare_login(context, account["id"], account["password"])
                pages.append((account["id"], page))
        except PlaywrightTimeoutError:
            print("找不到登入欄位，網站版面可能已變更，請檢查 login.py 中的選擇器。")
            browser.close()
            sys.exit(1)

        print("=" * 60)
        print(f"共 {len(pages)} 組帳號已自動填入身分證、密碼與驗證碼。")
        print("請切換到跳出的瀏覽器視窗逐一確認內容無誤，回到這個終端機視窗按 Enter 鍵，會一次把所有帳號送出登入。")
        print("=" * 60)
        input("按 Enter 繼續送出登入...")

        for tbb_id, page in pages:
            page.locator("#Image22").click()
            page.wait_for_timeout(3000)
            print(f"[{tbb_id}] 目前頁面網址: {page.url}")

        input("按 Enter 結束並關閉瀏覽器...")

        browser.close()


if __name__ == "__main__":
    main()
