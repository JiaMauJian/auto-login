"""
登入 tbbstock.com.tw 的核心邏輯，供持股同步 GUI（ui.py，經由 fetch.py）呼叫。

do_login()：開一個分頁，自動填入身分證、密碼與驗證碼並送出登入，回傳該分頁。
多組帳號設定：.env 用 TBB_ID_1/TBB_PASSWORD_1、TBB_ID_2/TBB_PASSWORD_2... 依序編號（見 .env.example）。
多帳號共用同一個瀏覽器 context（因此也共用同一組 cookie/session）：這個網站是
Java/Servlet 架構，登入狀態靠 JSESSIONID 這類 cookie 辨識，換帳號登入會把前一個
的 session 頂掉。GUI 靠 fetch._ensure_one 在換帳號前先把上一組的 cookie
收下來，需要用時再換回去，才不必每次都重新跑一遍登入流程；
也因為 cookie 只有一組，抓資料一定是「一組登入完就立刻抓完他的」（見 fetch.collect）。

瀏覽器設定：預設用 Playwright 自己下載的 Chromium、每次都是全新的空白 profile（沒有任何登入狀態）。
想改用電腦上已安裝的 Chrome，在 .env 設 BROWSER_CHANNEL=chrome；
想保留狀態（Cookie、tbbstock 數位憑證），再加上 USER_DATA_DIR 指定使用者資料夾（見 .env.example）。

驗證碼取得方式：頁面載入過程中瀏覽器會呼叫 VerifyNumberServlet，
該次請求的回應內容本身就是明文數字（例如 "86176"），網站再用這組數字畫成 canvas 圖案顯示。
因此直接攔截這次請求的回應即可取得驗證碼，不需要做圖片辨識。
"""

import os
import subprocess
import sys
import traceback
from datetime import datetime
from pathlib import Path

from dotenv import load_dotenv
from playwright.sync_api import (
    Error as PlaywrightError,
    TimeoutError as PlaywrightTimeoutError,
)

from dev_tools import simulate
from util import env_int


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
#
# 網站慢的時候第二次請求可能還沒回來就被填掉，可以在 .env 設 VERIFY_SETTLE_MS 加大。
# 上限 5000：再多就不是「等它安定」而是每個帳號都白白多等五秒，通常代表數字填錯了。
VERIFY_SETTLE_MS = max(0, min(5000, env_int("VERIFY_SETTLE_MS", 200)))

# 登入表單最多等多久（毫秒）才判定「這個瀏覽器裡已經有人登入著」。等不到就清 cookie
# 重來一次（見 do_login），所以這一段不能設太長 —— 換交易人時每次都要先耗掉它。
SWITCH_USER_TIMEOUT_MS = 6000

# 送出登入後最多等多久（毫秒）讓頁面換完。逾時不算失敗，只代表網站沒有換頁
# （例如驗證碼錯誤被擋在原頁），程式照樣往下印出目前網址讓使用者自己判斷。
LOGIN_NAV_TIMEOUT_MS = 10000


def pause(message):
    """等使用者按 Enter；沒有可用的標準輸入時（例如被程式呼叫）直接略過。"""
    try:
        input(message)
    except EOFError:
        pass


def log_crash(detail):
    """
    把啟動階段的例外寫進 exe 旁邊的 crash.log。

    exe 打包成 --windowed 後沒有主控台，print()/pause() 使用者完全看不到，
    只能寫成檔案，事後請使用者把這個檔案的內容貼給你。
    """
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    try:
        with open(app_dir() / "crash.log", "a", encoding="utf-8") as f:
            f.write(f"\n[{timestamp}]\n{detail}\n")
    except OSError:
        pass  # 連寫檔都失敗（例如資料夾沒有寫入權限），沒有更後路了


def _enter_pressed():
    """
    有沒有按下 Enter（不阻塞）。回傳 None 代表這個環境讀不到鍵盤，只能改用別的方式結束。

    不能用 input()，那會整個卡住 —— 卡住的後果見 wait_until_finished 的說明。
    """
    if not sys.stdin or not sys.stdin.isatty():
        return None
    try:
        import msvcrt
    except ImportError:
        return None
    try:
        pressed = False
        while msvcrt.kbhit():
            if msvcrt.getwch() in ("\r", "\n"):
                pressed = True
        return pressed
    except OSError:
        return None


def wait_until_finished(context):
    """
    登入完成後停在這裡，直到使用者按 Enter 或自己把瀏覽器關掉。

    這段等待「必須持續呼叫 Playwright」，也就是下面那個 wait_for_timeout。
    Playwright 的同步 API 只有在程式進入 Playwright 呼叫時才會去處理瀏覽器送來的事件，
    而網站用 window.open 開出來的新視窗要等 Playwright 接手初始化之後才會開始載入。

    原本這裡是 input()，一卡住就整個停擺，於是使用者在網站上點「簡易看盤下單」
    （<a href="javascript:fastQuoteUtil.openWinURL('../FastQuote/index.jsp')">，
    內部是 window.open）跳出來的視窗會一直停在 about:blank 空白，
    按了 Enter 又會直接連瀏覽器一起關掉，永遠看不到它載入。
    """
    print("按 Enter 結束並關閉瀏覽器（也可以直接把瀏覽器視窗關掉）...", flush=True)

    while True:
        pressed = _enter_pressed()
        if pressed is None:
            return          # 讀不到鍵盤（例如被別的程式呼叫），維持原本「不等待」的行為
        if pressed:
            return

        # 分頁全關掉 = 使用者自己關了瀏覽器，不要再等下去。
        try:
            pages = [pg for pg in context.pages if not pg.is_closed()]
            if not pages:
                print("瀏覽器已關閉。")
                return
            pages[0].wait_for_timeout(150)   # 這行同時是「驅動 Playwright」的關鍵
        except PlaywrightError:
            print("瀏覽器已關閉。")
            return


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


def launch_options():
    """
    依 .env 決定要開哪一個瀏覽器。

    BROWSER_CHANNEL：chrome / msedge / chrome-beta ...，用電腦上已安裝的正式版瀏覽器；
                     留空則用 Playwright 自己下載的 Chromium。
    BROWSER_PATH：直接指定執行檔完整路徑（例如某些綠色版 Chrome），會蓋過 BROWSER_CHANNEL。
    BROWSER_PROFILE_DIR：Chrome 使用者資料夾底下的哪一個設定檔（Default、Profile 1...），
                         只有搭配 USER_DATA_DIR 才有意義。

    --start-maximized：讓視窗一開起來就最大化（要搭配 no_viewport 才有效，見 open_context）。
    """
    options = {"headless": False}
    args = ["--start-maximized"]

    executable_path = os.getenv("BROWSER_PATH", "").strip()
    channel = os.getenv("BROWSER_CHANNEL", "").strip()
    if executable_path:
        options["executable_path"] = executable_path
    elif channel:
        options["channel"] = channel

    profile_dir = os.getenv("BROWSER_PROFILE_DIR", "").strip()
    if profile_dir:
        args.append(f"--profile-directory={profile_dir}")

    options["args"] = args

    return options


def user_data_dir():
    """
    回傳 .env 的 USER_DATA_DIR（使用者資料夾）；留空代表每次都用全新的暫時 profile。

    指定資料夾後，Cookie 與 tbbstock 的數位憑證都保存在這個資料夾，下次執行就直接沿用。
    可以用相對路徑（相對於 .env 所在資料夾），也支援 %LOCALAPPDATA% 這類環境變數寫法。
    """
    raw = os.getenv("USER_DATA_DIR", "").strip().strip('"')
    if not raw:
        return None
    path = Path(os.path.expandvars(os.path.expanduser(raw)))
    if not path.is_absolute():
        path = app_dir() / path
    return path


def open_context(p):
    """
    開啟瀏覽器並回傳 (context, browser)。

    有指定 USER_DATA_DIR 時用 launch_persistent_context（登入狀態會留在該資料夾，
    此時沒有獨立的 browser 物件，回傳的 browser 為 None）；
    沒指定時維持原本行為：開全新的暫時 profile。
    第一次在新電腦上執行、Playwright 的 Chromium 還沒下載時會自動補裝。

    no_viewport=True：Playwright 預設會把網頁鎖在 1280x720 的模擬 viewport，
    視窗再大網頁也只畫在左上角一小塊、右邊下面留白，而且拉大視窗也不會跟著變。
    設 True 之後 viewport 就跟著實際視窗大小走（配合 --start-maximized 開起來就是滿版）。
    注意這是 context 選項：launch() 不吃，只能放在 new_context()；
    launch_persistent_context() 則是 launch 與 context 選項合併，可以直接帶。
    """
    options = launch_options()
    profile_path = user_data_dir()

    def launch():
        if profile_path is None:
            browser = p.chromium.launch(**options)
            return browser.new_context(no_viewport=True), browser
        profile_path.mkdir(parents=True, exist_ok=True)
        return p.chromium.launch_persistent_context(str(profile_path), no_viewport=True, **options), None

    try:
        return launch()
    except PlaywrightError as exc:
        message = str(exc)
        if "Executable doesn't exist" in message and "channel" not in options and "executable_path" not in options:
            print("找不到 Chromium，第一次執行需要下載（約 150MB，只需一次）...")
            install_chromium()
            return launch()
        if profile_path is not None:
            print(f"開啟瀏覽器失敗，使用者資料夾: {profile_path}")
            print("如果這是你平常在用的 Chrome 資料夾，請先把所有 Chrome 視窗完全關掉再執行一次")
            print("（同一個資料夾不能同時被兩個 Chrome 開著）。")
        raise


def load_accounts():
    """
    .env 裡的帳號設定，依序編號。

    後面會接上模擬用的假帳號（.env 的 SIMULATE_ACCOUNTS，沒設就一個都沒有，
    正式部署的機器上等於這件事不存在）。假帳號帶 fake 旗標，
    不會去登入任何網站，詳見 dev_tools/simulate.py。
    """
    accounts = []
    i = 1
    while True:
        tbb_id = os.getenv(f"TBB_ID_{i}")
        tbb_password = os.getenv(f"TBB_PASSWORD_{i}")
        if not tbb_id or not tbb_password:
            break
        accounts.append({"id": tbb_id, "password": tbb_password})
        i += 1

    accounts.extend(simulate.fake_accounts())
    return accounts


def do_login(context, tbb_id, tbb_password, page=None):
    """
    開一個新分頁，填入帳密與驗證碼後送出登入。回傳該分頁物件。
    傳入 page 可以沿用既有分頁（persistent context 啟動時會自帶一個空白分頁）。
    """
    if page is None:
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
    #
    # 表單沒出現的話多半不是網站改版，而是這個瀏覽器裡已經有人登入著：網站看到
    # cookie 就把人導去別的頁，登入表單根本不會畫出來。換交易人就一定會遇到
    # （整個瀏覽器只有一組 session，見 fetch.ensure_logged_in）。所以先清掉
    # cookie 再回登入頁一次 —— 那是網站認人的唯一依據，清掉就等於登出。
    # 憑證不在 cookie 裡（它在 Windows 的憑證存放區與這個 Chrome profile 裡），
    # 不會被這一步弄掉。
    mode_link = page.locator("a:visible", has_text="身份證登入").first
    try:
        mode_link.wait_for(state="visible", timeout=SWITCH_USER_TIMEOUT_MS)
    except PlaywrightTimeoutError:
        print(f"[{tbb_id}] 登入頁沒有出現登入表單，清掉 cookie（等於登出）再試一次。")
        context.clear_cookies()
        page.goto(LOGIN_URL)
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

    # 送出登入前再等一下，避免驗證碼還沒套用就按下登入。
    page.wait_for_timeout(VERIFY_SETTLE_MS)

    # 等待要在按下登入「之前」就開始，否則換頁太快會來不及攔到。
    # 換頁一完成就往下走，不用固定乾等；沒換頁時（登入被擋在原頁）逾時當作正常情況吞掉。
    try:
        with page.expect_navigation(wait_until="load", timeout=LOGIN_NAV_TIMEOUT_MS):
            page.locator("#Image22").click()
    except PlaywrightTimeoutError:
        print(f"[{tbb_id}] 送出後頁面沒有跳轉，請確認是否登入失敗。")

    print(f"[{tbb_id}] 已送出登入，目前頁面網址: {page.url}")

    return page


def route():
    """
    只打包一個 exe，靠參數決定要做什麼：

        tbb-login.exe              持股同步介面（不帶參數就是它，雙擊直接開 GUI）
        tbb-login.exe --sim-excel  在 Excel 補上模擬用的分頁（測試用，見 dev_tools/sim_excel.py）

    不做成好幾個 exe，是因為 Playwright 那包東西會被各塞一份，dist 直接肥好幾倍，
    而部署方式是整包資料夾複製到目標電腦。

    GUI 是預設行為，因為登入、抓網頁、寫 Excel 現在全部都在 GUI 裡做（見
    ui_background.py），GUI 錯誤也全部走 messagebox、不靠印在主控台上，exe
    也打包成 --windowed，讓使用者可以直接雙擊 tbb-login.exe 開介面。

    ui 刻意在函式裡才 import：它反過來 import 這個模組，寫在檔案最上面會變成
    循環匯入。
    """
    args = sys.argv[1:]

    if args and args[0] == "--sim-excel":
        from dev_tools import sim_excel
        sys.argv = [sys.argv[0]] + args[1:]
        sim_excel.main()
    else:
        import ui
        ui.main()


if __name__ == "__main__":
    # exe 打包成 --windowed（沒有主控台），route() 丟出來、沒被 ui.py 自己的
    # messagebox 接住的例外（例如還沒進到 GUI 就炸掉），只能寫進 crash.log，
    # print()/pause() 在這裡使用者根本看不到。
    try:
        route()
    except SystemExit as exc:
        if exc.code:
            log_crash(f"SystemExit: {exc.code!r}")
        raise
    except Exception:
        log_crash(traceback.format_exc())
        sys.exit(1)
