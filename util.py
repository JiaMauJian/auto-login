"""
數字、格子名稱與 .env 設定的小轉換。沒有任何相依套件，所以每個模組都能安心引用。
"""

import os

# 金額比對用的容差。Excel 存的是浮點數，893.0 有可能變成 892.9999999999999，
# 直接用 == 比會說「不一樣」，畫面上就會出現一個其實沒變、卻被當成有變化的數字。
EPSILON = 0.005


def env_int(key, default):
    """
    從 .env 讀一個整數。沒設、留空、或填了看不懂的東西都回 default ——
    這些都是可有可無的微調，不值得為了一個打錯的數字讓程式跑不起來。
    """
    raw = os.getenv(key, "").strip()
    if not raw:
        return default
    try:
        return int(float(raw))
    except ValueError:
        return default


def to_num(value, default=0.0):
    """網頁回傳的數字幾乎都是字串，有些欄位還會是空字串或 null，統一轉成數字。"""
    try:
        return float(str(value).strip())
    except (TypeError, ValueError):
        return default


def cell_name(row, col):
    """(8, 2) -> B8。這張表只用到 A~F，不處理 AA 之後的欄。"""
    return f"{chr(64 + col)}{row}"


def same_number(left, right):
    """把兩個值當成金額比。任一邊不是數字（例如空格）就算不同。"""
    a, b = to_num(left, None), to_num(right, None)
    if a is None or b is None:
        return False
    return abs(a - b) <= EPSILON


def pad(text, width):
    """
    按終端機的顯示寬度補空白。

    不能用 str.ljust —— 它把「股數」算成 2 格，實際上中文字佔 4 格，
    只要欄位裡有中文，整張表就會歪掉。
    """
    import unicodedata

    size = sum(2 if unicodedata.east_asian_width(ch) in "WF" else 1 for ch in str(text))
    return f"{text}{' ' * max(0, width - size)}"


def values_match(left, right):
    """same_number，但兩邊都是空的算相同 —— 空格沒被動過也是一種「沒變」。"""
    if left is None and right is None:
        return True
    return same_number(left, right)


def show(value):
    """
    印給人看的數字。None 顯示成 (空)，數字加千分位、去掉多餘的零。

    為什麼不用 :g
    -------------
    :g 只留 6 位有效數字，超過就跳成 3.12966e+06。現金餘額動輒七八位數，
    畫面上全是科學記號，人根本核對不出「三百一十二萬」還是「三千一百萬」。
    改成 12 位有效數字 + 千分位，錢的位數一眼就數得出來，
    浮點數的 34.550000000000004 也還是會收成 34.55。
    """
    if value is None:
        return "(空)"
    number = to_num(value, None)
    if number is None:
        return str(value).strip()
    return f"{number:,.12g}"
