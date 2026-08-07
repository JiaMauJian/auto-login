"""
自動登入 tbbstock.com.tw（支援多組帳號）
流程：自動開啟瀏覽器 -> 依 .env 內的每組帳號各開一個分頁 -> 自動填入身分證、密碼與驗證碼
     -> 依 .env 的 AUTO_SUBMIT 決定要直接送出登入，還是等使用者確認後才一次送出。

AUTO_SUBMIT=true（預設）：填完就直接按下登入。
AUTO_SUBMIT=false：全部填完後停在終端機等按 Enter，方便先人工核對內容，再一次送出所有帳號。

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
import subprocess
import sys
import traceback
from pathlib import Path

from dotenv import load_dotenv
from playwright.sync_api import (
    sync_playwright,
    Error as PlaywrightError,
    TimeoutError as PlaywrightTimeoutError,
)


def app_dir():
    """打包成 exe 後回傳 exe 所在資料夾，直接跑 .py 時回傳原始碼資料夾。.env 要放在這裡。"""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).parent
    return Path(__file__).parent


# 打包成 exe 時工作目錄不一定等於 exe 位置，所以明確指定 .env 路徑。
# utf-8-sig：用記事本存 .env 會在檔頭加上 BOM，不處理的話第一個設定會讀不到。
load_dotenv(app_dir() / ".env", encoding="utf-8-sig")

LOGIN_URL = "https://www.tbbstock.com.tw/tbb/index/home.jsp"

# 攔截到驗證碼後多等一下（毫秒），讓頁面可能的第二次 VerifyNumberServlet 請求先回來，
# 以最後一次的值為準；送出登入前也再等同樣的時間，避免驗證碼還沒套用就按下登入。
VERIFY_SETTLE_MS = 200


def pause(message):
    """等使用者按 Enter；沒有可用的標準輸入時（例如被程式呼叫）直接略過。"""
    try:
        input(message)
    except EOFError:
        pass


def configure_browsers_path():
    """
    決定要去哪裡找 Chromium。

    打包成 exe 後，Playwright 會自動把 PLAYWRIGHT_BROWSERS_PATH 設成 "0"
    （見 playwright/_impl/_transport.py），意思是「只找打包內容裡的瀏覽器」，
    結果是明明電腦上已經裝過 Chromium 也找不到。所以這裡自己指定：

    1. exe 旁邊（或打包內容裡）有 ms-playwright 資料夾就用它 —— 整包複製到別台電腦也能跑；
    2. 否則用 Windows 的標準位置 %LOCALAPPDATA%\\ms-playwright，跟 playwright install 裝的共用。
    """
    if os.getenv("PLAYWRIGHT_BROWSERS_PATH"):
        return  # 使用者自己在環境變數指定的優先

    candidates = [app_dir() / "ms-playwright"]
    meipass = getattr(sys, "_MEIPASS", None)
    if meipass:
        candidates.append(Path(meipass) / "ms-playwright")

    for path in candidates:
        if path.is_dir():
            os.environ["PLAYWRIGHT_BROWSERS_PATH"] = str(path)
            return

    local_appdata = os.getenv("LOCALAPPDATA")
    if local_appdata:
        os.environ["PLAYWRIGHT_BROWSERS_PATH"] = str(Path(local_appdata) / "ms-playwright")


def install_chromium():
    """呼叫 Playwright 內建的 node driver 下載 Chromium（exe 內沒有 pip/python 可用，所以直接跑 driver）。"""
    from playwright._impl._driver import compute_driver_executable, get_driver_env

    node, cli = compute_driver_executable()
    subprocess.run([node, cli, "install", "chromium"], env=get_driver_env(), check=True)


def launch_browser(p):
    """開啟 Chromium；第一次在新電腦上執行、瀏覽器還沒下載時自動補裝。"""
    try:
        return p.chromium.launch(headless=False)
    except PlaywrightError as exc:
        if "Executable doesn't exist" not in str(exc):
            raise
        print("找不到 Chromium，第一次執行需要下載（約 150MB，只需一次）...")
        install_chromium()
        return p.chromium.launch(headless=False)


def env_flag(name, default=True):
    """讀取 .env 的布林設定，接受 true/1/yes/y/on 這類寫法（不分大小寫）。"""
    raw = os.getenv(name)
    if raw is None or raw.strip() == "":
        return default
    return raw.strip().lower() in ("1", "true", "yes", "y", "on")


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


def submit_login(tbb_id, page):
    page.wait_for_timeout(VERIFY_SETTLE_MS)
    page.locator("#Image22").click()
    page.wait_for_timeout(3000)
    print(f"[{tbb_id}] 已送出登入，目前頁面網址: {page.url}")


def do_login(context, tbb_id, tbb_password, auto_submit=True):
    """開一個新分頁並填入帳密與驗證碼；auto_submit 為 True 時直接送出登入。回傳該分頁物件。"""
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

    # 頁面載入過程可能會再打一次 VerifyNumberServlet 換新的驗證碼，
    # 先等一下讓它安定下來，再用「最後一次」攔截到的值填入，避免填到已經失效的舊驗證碼。
    page.wait_for_timeout(VERIFY_SETTLE_MS)

    if "value" not in verify_number:
        print(f"[{tbb_id}] 沒有攔截到 VerifyNumberServlet 的回應，請確認網站是否改版，改回手動輸入驗證碼。")
    else:
        page.locator("#NumberLabel").fill(verify_number["value"])
        print(f"[{tbb_id}] 已自動取得並填入驗證碼: {verify_number['value']}")

    if auto_submit:
        submit_login(tbb_id, page)

    return page


def main():
    accounts = load_accounts()
    if not accounts:
        print(f"找不到帳號設定。請在 {app_dir()} 放一個 .env 檔（可複製 .env.example），")
        print("並依 TBB_ID_1/TBB_PASSWORD_1、TBB_ID_2/TBB_PASSWORD_2... 填入至少一組帳號。")
        sys.exit(1)

    auto_submit = env_flag("AUTO_SUBMIT", default=True)
    configure_browsers_path()

    with sync_playwright() as p:
        browser = launch_browser(p)
        context = browser.new_context()

        pages = []
        try:
            for account in accounts:
                page = do_login(context, account["id"], account["password"], auto_submit)
                pages.append((account["id"], page))
        except PlaywrightTimeoutError:
            print("找不到登入欄位，網站版面可能已變更，請檢查 login.py 中的選擇器。")
            browser.close()
            sys.exit(1)

        print("=" * 60)
        if auto_submit:
            print(f"共 {len(pages)} 組帳號已自動送出登入。")
        else:
            print(f"共 {len(pages)} 組帳號已自動填入身分證、密碼與驗證碼（AUTO_SUBMIT=false，尚未送出）。")
            print("請切換到瀏覽器視窗逐一確認內容無誤，再回到終端機按 Enter，會一次把所有帳號送出登入。")
            print("=" * 60)
            pause("按 Enter 送出登入...")
            for tbb_id, page in pages:
                submit_login(tbb_id, page)
        print("=" * 60)
        pause("按 Enter 結束並關閉瀏覽器...")

        browser.close()


if __name__ == "__main__":
    # 打包成 exe 用滑鼠雙擊執行時，程式一結束視窗就會關掉，
    # 所以出錯時要停下來讓使用者看得到訊息。
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
