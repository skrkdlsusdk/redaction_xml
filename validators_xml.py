# server/validators_xml.py
import re
from datetime import datetime

# 숫자만 추출
def _digits(s: str) -> str:
    return re.sub(r"\D", "", s or "")

# Luhn 체크
def _luhn_ok(digits: str) -> bool:
    s = 0
    rev = digits[::-1]
    for i, ch in enumerate(rev):
        n = ord(ch) - 48
        if i % 2 == 1:
            n *= 2
            if n > 9:
                n -= 9
        s += n
    return s % 10 == 0

def is_valid_card(s: str) -> bool:
    d = _digits(s)
    # 길이 / IIN 간단 필터 (Visa/Master/Amex/Discover 포함)
    if not (13 <= len(d) <= 19):
        return False
    if not _luhn_ok(d):
        return False
    # IIN 대략 필터: 2-6 시작 (Master 51-55, 2221-2720 등), 4(Visa), 34/37(Amex), 6011/65/64x/622(Discover)
    if not (
        d.startswith("4") or
        d.startswith("5") or
        d.startswith("2") or
        d.startswith("34") or d.startswith("37") or
        d.startswith("6011") or d.startswith("65") or d.startswith("64") or d.startswith("622")
    ):
        return False
    return True

def is_valid_email(s: str) -> bool:
    return re.fullmatch(r"[A-Za-z0-9._%+-]+@(?:[A-Za-z0-9-]+\.)+[A-Za-z]{2,}", s or "") is not None

def is_valid_phone_mobile(s: str) -> bool:
    d = _digits(s)
    # 010/011/016/017/018/019 + 7~8자리
    return re.fullmatch(r"01[016789]\d{7,8}", d or "") is not None

def is_valid_phone_city(s: str) -> bool:
    d = _digits(s)
    # 02 + 7~8자리, 또는 0(3x~6x) + 8자리
    if re.fullmatch(r"02\d{7,8}", d or ""):
        return True
    return re.fullmatch(r"0(?:3[1-3]|4[1-4]|5[1-5]|6[1-4])\d{8}", d or "") is not None

# 생년월일 6자리(yyMMdd)
def is_valid_date6(s: str) -> bool:
    d = _digits(s)
    if len(d) != 6:
        return False
    try:
        y = int(d[:2])
        m = int(d[2:4])
        dd = int(d[4:6])
        # 00~현재 연도의 뒤 2자리까지 허용 → 세기 보정
        this_year = int(datetime.today().strftime("%y"))
        full_year = 1900 + y if y > this_year else 2000 + y
        datetime(full_year, m, dd)
        return True
    except ValueError:
        return False

# 주민등록번호(날짜+체크섬)
def is_valid_rrn_checksum(s: str) -> bool:
    d = _digits(s)
    if len(d) != 13:
        return False
    weights = [2,3,4,5,6,7,8,9,2,3,4,5]
    total = sum(int(x)*w for x, w in zip(d[:-1], weights))
    chk = (11 - (total % 11)) % 10
    return chk == int(d[-1])

def is_valid_rrn(s: str) -> bool:
    d = _digits(s)
    if len(d) != 13:
        return False
    if not is_valid_date6(d[:6]):
        return False
    return is_valid_rrn_checksum(d)

# 외국인등록번호 (간단판)
def is_valid_fgn_checksum(s: str) -> bool:
    d = _digits(s)
    if len(d) != 13:
        return False
    weights = [2,3,4,5,6,7,8,9,2,3,4,5]
    total = sum(int(x)*w for x, w in zip(d[:-1], weights))
    chk = (11 - (total % 11) + 2) % 10
    return chk == int(d[-1])

def is_valid_fgn(s: str) -> bool:
    d = _digits(s)
    if len(d) != 13:
        return False
    # 생년월일 6자리 + 구분코드(5~8중 하나)
    if not is_valid_date6(d[:6]):
        return False
    return is_valid_fgn_checksum(d)

# 운전면허 (간단 정합성)
def is_valid_driver_license(s: str, opts: dict | None = None) -> bool:
    d = _digits(s)
    return 10 <= len(d) <= 12
