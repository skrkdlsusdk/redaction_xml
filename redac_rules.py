# redac_rules.py
import re

# --- validators 임포트: 패키지/스크립트 실행 모두 지원 (방법 B) ---
try:
    # 패키지로 실행할 때
    from .validators import (
        is_valid_rrn,
        is_valid_phone_mobile,
        is_valid_phone_city,
        is_valid_email,
        is_valid_card,
    )
except ImportError:
    # 단일 스크립트로 실행할 때
    from validators import (
        is_valid_rrn,
        is_valid_phone_mobile,
        is_valid_phone_city,
        is_valid_email,
        is_valid_card,
    )

# --- 주민등록번호(간단 월/일 범위 반영, 하이픈 선택) ---
RRN_RE = re.compile(r"(?:\d{2}(?:0[1-9]|1[0-2])(?:0[1-9]|[12]\d|3[01]))-?\d{7}")

# --- 카드번호 (숫자/하이픈/공백 허용, 15~16 digits) ---
CARD_RE = re.compile(r"(?:\d[ -]?){15,16}")

# --- 이메일 ---
EMAIL_RE = re.compile(r"[A-Za-z0-9._%+-]+@(?:[A-Za-z0-9-]+\.)+[A-Za-z]{2,}")

# --- 휴대폰 ---
MOBILE_RE = re.compile(r"01[016789]-?\d{3,4}-?\d{4}")

# --- 지역번호 ---
CITY_RE = re.compile(r"(?:02|0(?:3[1-3]|4[1-4]|5[1-5]|6[1-4]))-?\d{3,4}-?\d{4}")

# --- 여권번호 (1~2 영문 + 7~8 숫자) ---
PASSPORT_RE = re.compile(r"[A-Z]{1,2}\d{7,8}")

# --- 운전면허번호 ---
# 신포맷: NN-NN-NNNNNN-NN
# 구포맷: NN-NN-NNNNNN (마지막 -NN 생략되는 경우 허용)
DRIVER_RE = re.compile(r"\d{2}-\d{2}-\d{6}(?:-\d{2})?")

# --- 룰 정의 ---
RULES = {
    "rrn": {
        "regex": RRN_RE,
        # 요청 options에서 rrn_checksum(기본 True) 반영
        "validator": lambda v, opts=None: is_valid_rrn(
            v, use_checksum=(opts or {}).get("rrn_checksum", True)
        ),
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
        # 카드 옵션(luhn, iin 등) 그대로 전달
        "validator": lambda v, opts=None: is_valid_card(v, options=opts),
    },
    "passport": {
        "regex": PASSPORT_RE,
        "validator": lambda v, _opts=None: True,
    },
    "driver_license": {
        "regex": DRIVER_RE,
        "validator": lambda v, _opts=None: True,
    },
}

# --- 프리셋 (API로 노출) ---
PRESET_PATTERNS = [
    {"name": "rrn",            "regex": RRN_RE.pattern,        "case_sensitive": False, "whole_word": False},
    {"name": "email",          "regex": EMAIL_RE.pattern,      "case_sensitive": False, "whole_word": False},
    {"name": "phone_mobile",   "regex": MOBILE_RE.pattern,     "case_sensitive": False, "whole_word": False},
    {"name": "phone_city",     "regex": CITY_RE.pattern,       "case_sensitive": False, "whole_word": False},
    {"name": "card",           "regex": CARD_RE.pattern,       "case_sensitive": False, "whole_word": False},
    {"name": "passport",       "regex": PASSPORT_RE.pattern,   "case_sensitive": False, "whole_word": False},
    {"name": "driver_license", "regex": DRIVER_RE.pattern,     "case_sensitive": False, "whole_word": False},
]
