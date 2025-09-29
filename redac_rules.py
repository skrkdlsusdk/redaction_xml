import re

# validators 임포트: 패키지/스크립트 실행 모두 지원
try:
    from .validators import (
        is_valid_rrn,
        is_valid_fgn,
        is_valid_phone_mobile,
        is_valid_phone_city,
        is_valid_email,
        is_valid_card,
        is_valid_driver_license,
    )
except ImportError:
    from validators import (
        is_valid_rrn,
        is_valid_fgn,
        is_valid_phone_mobile,
        is_valid_phone_city,
        is_valid_email,
        is_valid_card,
        is_valid_driver_license,
    )

# --- 정규식들 ---

# 주민등록번호 (내국인)
RRN_RE = re.compile(
    r"(?:\d{2}(?:0[1-9]|1[0-2])"      # 연월
    r"(?:0[1-9]|[12]\d|3[01]))"       # 일
    r"-?[1234]\d{6}"
)

# 외국인등록번호
FGN_RE = re.compile(
    r"(?:\d{2}(?:0[1-9]|1[0-2])"      # 연월
    r"(?:0[1-9]|[12]\d|3[01]))"       # 일
    r"-?[5678]\d{6}"                  # 구분코드(5~8) + 나머지 6자리
)

# 카드번호 (하이픈/공백 허용, 15~16자리 후보)
CARD_RE = re.compile(r"(?:\d[ -]?){15,16}")

# 이메일
EMAIL_RE = re.compile(r"[A-Za-z0-9._%+-]+@(?:[A-Za-z0-9-]+\.)+[A-Za-z]{2,}")

# 휴대폰
MOBILE_RE = re.compile(r"01[016789]-?\d{3,4}-?\d{4}")

# 지역번호 유선
CITY_RE = re.compile(r"(?:02|0(?:3[1-3]|4[1-4]|5[1-5]|6[1-4]))-?\d{3,4}-?\d{4}")

# 여권번호 (구/신)
PASSPORT_RE = re.compile(
    r"(?:"
    r"(?:[MSRODG]\d{8})"               # 구여권: M12345678
    r"|"
    r"(?:[MSRODG]\d{3}[A-Z]\d{4})"     # 신여권: M123A4567
    r")"
)

# 운전면허번호 (구/신 혼용 허용)
DRIVER_RE = re.compile(r"\d{2}-?\d{2}-?\d{6}-?\d{2}")

# --- RULES 매핑 ---
RULES = {
    "rrn": {
        "regex": RRN_RE,
        "validator": is_valid_rrn,
    },
    "fgn": {
        "regex": FGN_RE,
        "validator": is_valid_fgn,
    },
    "email": {
        "regex": EMAIL_RE,
        "validator": is_valid_email,
    },
    "phone_mobile": {
        "regex": MOBILE_RE,
        "validator": is_valid_phone_mobile,
    },
    "phone_city": {
        "regex": CITY_RE,
        "validator": is_valid_phone_city,
    },
    "card": {
        "regex": CARD_RE,
        "validator": is_valid_card,
    },
    "passport": {
        "regex": PASSPORT_RE,
        "validator": lambda v, _opts=None: True,  # 형식만 맞으면 레닥션
    },
    "driver_license": {
        "regex": DRIVER_RE,
        "validator": is_valid_driver_license,
    },
}

# API 노출용(선택)
PRESET_PATTERNS = [
    {
        "name": name,
        "regex": rule["regex"].pattern,
        "case_sensitive": False,
        "whole_word": False,
    }
    for name, rule in RULES.items()
]
