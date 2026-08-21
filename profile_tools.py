"""
Chrome Profile 的建立與 tbbstock 數位憑證的搬遷／狀態，供 ui.py 的「憑證」分頁使用。

這裡把 setup-profile.ps1 與 migrate-cert.ps1 的邏輯搬進 Python：
    - setup-profile.ps1 -> create_profile / launch_manual_chrome / profile_initialized
    - migrate-cert.ps1  -> scan_cert_sources / copy_cert / scan_profile_dir

原本兩支 .ps1 還是留著（給不方便開介面、或要跨機複製憑證的情況用），
這裡只重做「介面按一按就會用到」的那條路：本機掃描、本機複製、
在同一台電腦上建立/重建資料夾。

為什麼跟 excel_io.py 一樣自己組 .env 的讀寫，而不是共用同一個函式：
兩邊要改的 key 不一樣（EXCEL_PATH／USER_DATA_DIR），硬拉一個共用函式出來
會多一層參數化，換不到什麼好處。
"""

import codecs
import datetime
import os
import shutil
import subprocess
from pathlib import Path

from login import app_dir

ENV_FILE = ".env"
ENV_KEY = "USER_DATA_DIR"
DEFAULT_NAME = "chrome-profile"

HOME_URL = "https://www.tbbstock.com.tw/tbb/index/home.jsp"

# leveldb 檔案裡出現這個字串，代表這個 profile 存過 tbbstock 的憑證資料（見 migrate-cert.ps1 檔頭）。
# 不能拿 "tbbstock" 本身當標記 —— localStorage key 是照 origin 存的，只要開過
# https://www.tbbstock.com.tw 這個網頁，不管有沒有憑證，這個字串就會寫進去。
MARKERS = ("TWCACertIdxRef",)


def current_raw():
    """.env 現在的 USER_DATA_DIR 原始字串（未展開），沒設就回預設資料夾名稱。"""
    return os.getenv(ENV_KEY, "").strip().strip('"') or DEFAULT_NAME


def resolve_path(raw):
    """
    跟 login.user_data_dir() 同一套規則：支援 %LOCALAPPDATA% 這類寫法，
    相對路徑以 .env 所在資料夾為準。raw 是空字串時回 None（代表不留存 profile）。
    """
    raw = (raw or "").strip().strip('"')
    if not raw:
        return None
    path = Path(os.path.expandvars(os.path.expanduser(raw)))
    if not path.is_absolute():
        path = app_dir() / path
    return path


def profile_subdir_name():
    """USER_DATA_DIR 底下實際存放憑證資料的子資料夾（Default、Profile 1...），對應 BROWSER_PROFILE_DIR（見 login.py）。"""
    return os.getenv("BROWSER_PROFILE_DIR", "").strip().strip('"') or "Default"


def default_chrome_dir():
    return Path(os.environ.get("LOCALAPPDATA", "")) / "Google" / "Chrome" / "User Data"


def is_default_chrome_dir(path):
    """擋掉指向 Chrome 預設使用者資料夾的路徑 —— Chrome 136 之後自動化連不上它（見 setup-profile.ps1）。"""
    try:
        return Path(path).resolve() == default_chrome_dir().resolve()
    except OSError:
        return False


def remember_user_data_dir(raw):
    """
    把選好的資料夾名稱寫回 .env 的 USER_DATA_DIR，只動那一行，其餘原封不動；
    BOM 也保留（記事本存過的 .env 檔頭會有 BOM，拿掉下次讀取會少第一個設定）。
    """
    env = app_dir() / ENV_FILE
    data = env.read_bytes() if env.is_file() else b""
    has_bom = data.startswith(codecs.BOM_UTF8)
    text = data.decode("utf-8-sig") if data else ""
    newline = "\r\n" if "\r\n" in text else "\n"

    lines = text.splitlines()
    entry = f"{ENV_KEY}={raw}"
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
        lines.append("# 使用者資料夾。介面上按「建立 Profile」時會自動改這一行。")
        lines.append(entry)

    body = (newline.join(lines) + newline).encode("utf-8")
    env.write_bytes((codecs.BOM_UTF8 if has_bom else b"") + body)
    os.environ[ENV_KEY] = raw


def find_chrome():
    """一般模式（非自動化）用的 chrome.exe 路徑，找不到回 None（見 setup-profile.ps1 的候選清單）。"""
    candidates = (
        Path(os.environ.get("ProgramFiles", r"C:\Program Files")) / "Google" / "Chrome" / "Application" / "chrome.exe",
        Path(os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)")) / "Google" / "Chrome" / "Application" / "chrome.exe",
        Path(os.environ.get("LOCALAPPDATA", "")) / "Google" / "Chrome" / "Application" / "chrome.exe",
    )
    return next((c for c in candidates if c.is_file()), None)


def scan_profile_dir(profile_dir):
    """
    這一個 Chrome/Edge profile（例如 .../User Data/Default）的 Local Storage 裡
    有沒有 tbbstock 憑證的痕跡。回傳 (found, hits, file_count, total_bytes)。

    .ldb 是 Snappy 壓縮過的，字串搜尋可能漏；.log（最近寫入）是未壓縮的 ——
    所以「找到」很可靠，「找不到」只能當成「很可能沒有」，跟 migrate-cert.ps1 同一個限制。
    """
    leveldb = Path(profile_dir) / "Local Storage" / "leveldb"
    if not leveldb.is_dir():
        return False, [], 0, 0

    files = list(leveldb.glob("*.log")) + list(leveldb.glob("*.ldb"))
    hits, total_bytes = set(), 0
    for f in files:
        try:
            data = f.read_bytes()
        except OSError:
            continue
        total_bytes += len(data)
        # ISO-8859-1：位元組原樣轉字元，不會因為切到多位元組字元中間而丟例外。
        text = data.decode("latin-1", errors="ignore")
        for marker in MARKERS:
            if marker in text:
                hits.add(marker)

    return bool(hits), sorted(hits), len(files), total_bytes


def has_cert(user_data_dir):
    """USER_DATA_DIR 底下任何一個子 profile（Default、Profile 1...）有沒有憑證痕跡。"""
    path = Path(user_data_dir)
    if not path.is_dir():
        return False
    for sub in path.iterdir():
        if sub.is_dir() and scan_profile_dir(sub)[0]:
            return True
    return False


def delete_profile(path):
    shutil.rmtree(path)


def launch_manual_chrome(chrome_exe, profile_path):
    """
    開一個一般模式（非自動化）的 Chrome，指到這個資料夾，用來讓 Chrome 把資料夾初始化出來。
    回傳 Popen；呼叫方負責之後用 profile_initialized() 確認、再用 chrome_pids_for_profile() 關掉它。
    """
    profile_path.mkdir(parents=True, exist_ok=True)
    args = [
        str(chrome_exe),
        f"--user-data-dir={profile_path}",
        "--no-first-run",
        "--no-default-browser-check",
        HOME_URL,
    ]
    return subprocess.Popen(args)


def profile_initialized(path):
    """Chrome 是不是真的把這個資料夾初始化過了：正常啟動過就會有 Local State 與 Default\\（見 setup-profile.ps1）。"""
    path = Path(path)
    return (path / "Local State").is_file() and (path / "Default").is_dir()


def _wmi_processes(image_name):
    """WMI 查詢：目前跑著的某支 .exe，回傳 [(pid, commandline), ...]。查不到（WMI 不可用）就回空清單。"""
    try:
        import win32com.client
        wmi = win32com.client.GetObject("winmgmts:")
        rows = wmi.ExecQuery(
            f"SELECT ProcessId, CommandLine FROM Win32_Process WHERE Name='{image_name}'"
        )
        return [(row.ProcessId, row.CommandLine or "") for row in rows]
    except Exception:
        return []


def browser_running(image_name):
    """這支瀏覽器的 .exe 現在是不是還在跑（不管開哪個 profile）。"""
    return bool(_wmi_processes(image_name))


def chrome_pids_for_profile(profile_path):
    """指到這個資料夾的 chrome.exe 行程 PID（--user-data-dir 出現在command line 裡）。"""
    target = str(Path(profile_path).resolve()).lower()
    return [pid for pid, cmdline in _wmi_processes("chrome.exe") if target in cmdline.lower()]


def kill_pids(pids):
    for pid in pids:
        subprocess.run(["taskkill", "/F", "/PID", str(pid)], capture_output=True, check=False)


# 掃描的瀏覽器來源，跟 migrate-cert.ps1 第 3 節同一份清單。
def _browsers():
    return (
        {"name": "Chrome", "exe": "chrome.exe",
         "user_data": Path(os.environ.get("LOCALAPPDATA", "")) / "Google" / "Chrome" / "User Data"},
        {"name": "Edge", "exe": "msedge.exe",
         "user_data": Path(os.environ.get("LOCALAPPDATA", "")) / "Microsoft" / "Edge" / "User Data"},
    )


def scan_cert_sources(target_profile_dir):
    """
    列出 Chrome/Edge 每一個 profile（Default、Profile 1...）的掃描結果，排除目標自己。
    回傳一串 dict：browser、exe、name、path、found、hits、files、bytes。
    """
    target = Path(target_profile_dir).resolve() if target_profile_dir else None
    candidates = []
    for browser in _browsers():
        user_data = browser["user_data"]
        if not user_data.is_dir():
            continue
        for sub in sorted(user_data.iterdir()):
            if not sub.is_dir():
                continue
            if sub.name != "Default" and not sub.name.startswith("Profile "):
                continue
            try:
                if target is not None and sub.resolve() == target:
                    continue
            except OSError:
                pass
            found, hits, files, total_bytes = scan_profile_dir(sub)
            candidates.append({
                "browser": browser["name"], "exe": browser["exe"], "name": sub.name,
                "path": sub, "found": found, "hits": hits, "files": files, "bytes": total_bytes,
            })
    return candidates


def copy_cert(source_profile_dir, target_profile_dir):
    """
    把來源 profile 的 Local Storage 整個複製到目標 profile，複製前備份目標現有的一份。
    回傳備份路徑（目標原本沒有 Local Storage 就回 None）。

    複製檔案不會經過 RA，來源那張憑證不會被作廢（見 migrate-cert.ps1 檔頭）。
    """
    source_ls = Path(source_profile_dir) / "Local Storage"
    target_ls = Path(target_profile_dir) / "Local Storage"
    if not source_ls.is_dir():
        raise FileNotFoundError(f"來源找不到 Local Storage 資料夾：{source_ls}")

    backup = None
    if target_ls.exists():
        stamp = datetime.datetime.now().strftime("%Y%m%d-%H%M%S")
        backup = target_ls.with_name(f"Local Storage.bak-{stamp}")
        shutil.move(str(target_ls), str(backup))

    shutil.copytree(source_ls, target_ls)

    # 上次執行留下的 LOCK 會讓 Chrome 以為資料夾還被佔用，複製過來就順手清掉。
    lock = target_ls / "leveldb" / "LOCK"
    if lock.exists():
        try:
            lock.unlink()
        except OSError:
            pass

    return backup
