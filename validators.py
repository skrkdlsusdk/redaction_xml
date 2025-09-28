# validators.py
import re
from datetime import datetime

def _digits(s: str) -> str:
    return re.sub(r"\D", "", s or "")

# --- 주민등록번호 ---
def is_valid_rrn_date_only(rrn: str) -> bool:
    d = _digits(rrn)
    if len(d) != 13:
        return False
    birth = d[:6]
    gender = d[6]
    try:
        y = ("20" if gender in "34" else "19") + birth[:2]
        dt = datetime.strptime(y + birth[2:], "%Y%m%d")
        if dt > datetime.today():
            return False
    except ValueError:
        return False
    return True

def is_valid_rrn_checksum(rrn: str) -> bool:
    d = _digits(rrn)
    if len(d) != 13:
        return False
    weights = [2,3,4,5,6,7,8,9,2,3,4,5]
    total = sum(int(x)*w for x, w in zip(d[:-1], weights))
    chk = (11 - (total % 11)) % 10
    return chk == int(d[-1])

def is_valid_rrn(rrn: str, use_checksum: bool = False) -> bool:
    if not is_valid_rrn_date_only(rrn):
        return False
    if use_checksum and not is_valid_rrn_checksum(rrn):
        return False
    return True

# --- 카드 ---
def _luhn_ok(d: str) -> bool:
    s = 0
    alt = False
    for ch in reversed(d):
        n = ord(ch) - 48
        if alt:
            n *= 2
            if n > 9:
                n -= 9
        s += n
        alt = not alt
    return (s % 10) == 0

def is_valid_card(number: str, options: dict | None = None) -> bool:
    opts = {"luhn": True, "iin": True}
    if options:
        opts.update(options)

    d = _digits(number)
    if len(d) not in (15, 16):
        return False

    if opts["iin"]:
        if len(d) == 16:
            prefix2 = int(d[:2]) if d[:2].isdigit() else None
            prefix4 = int(d[:4]) if d[:4].isdigit() else None
            if d[0] == "4":
                pass  # Visa
            elif d[0] == "5" and 51 <= int(d[:2]) <= 55:
                pass  # Master
            elif d[0] == "2" and 2221 <= prefix4 <= 2720:
                pass  # Master 2-series
            elif d[0] == "6":
                pass  # Discover
            elif d[0] == "9":
                pass  # 내부 정책 허용 BIN
            elif prefix2 == 35:
                pass  # JCB
            else:
                return False
        else:  # 15자리
            if not (d.startswith("34") or d.startswith("37")):
                return False

    if opts["luhn"] and not _luhn_ok(d):
        return False

    return True

# --- 휴대폰 ---
def is_valid_phone_mobile(number: str, options: dict | None = None) -> bool:
    d = _digits(number)
    return d.startswith("010") and len(d) == 11

# --- 지역번호 ---
def is_valid_phone_city(number: str, options: dict | None = None) -> bool:
    d = _digits(number)
    if d.startswith("02") and 9 <= len(d) <= 10:
        return True
    if d[:2] in {f"0{x}" for x in range(31, 65)} and 10 <= len(d) <= 11:
        return True
    return False

# --- 이메일 ---
def is_valid_email(addr: str, options: dict | None = None) -> bool:
    pat = re.compile(r"^[A-Za-z0-9._%+-]+@(?:[A-Za-z0-9-]+\.)+[A-Za-z]{2,}$")
    return bool(pat.match(addr or ""))
