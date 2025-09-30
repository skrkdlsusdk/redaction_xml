# server/redac_rules_xml.py
import re

# validators 임포트: 패키지/스크립트 실행 모두 지원
try:
    from .validators_xml import (
        is_valid_rrn, is_valid_fgn, is_valid_email, is_valid_phone_mobile,
        is_valid_phone_city, is_valid_card, is_valid_driver_license,
    )
except Exception:
    from validators_xml import (  # type: ignore
        is_valid_rrn, is_valid_fgn, is_valid_email, is_valid_phone_mobile,
        is_valid_phone_city, is_valid_card, is_valid_driver_license,
    )

# 기본 패턴들 (digits-only 매칭 후 validator로 2차 필터)
RRN_RE = re.compile(r"\d{6}-?\d{7}")
FGN_RE = re.compile(r"\d{6}-?\d{7}")
EMAIL_RE = re.compile(r"[A-Za-z0-9._%+-]+@(?:[A-Za-z0-9-]+\.)+[A-Za-z]{2,}")
MOBILE_RE = re.compile(r"01[016789]-?\d{3,4}-?\d{4}")
CITY_RE = re.compile(r"(?:02|0(?:3[1-3]|4[1-4]|5[1-5]|6[1-4]))-?\d{3,4}-?\d{4}")
CARD_RE = re.compile(r"\d[\d -]{11,}\d")  # 느슨히 찾고 Luhn/IIN으로 거른다

# 여권 (구/신여권 일부 패턴)
PASSPORT_RE = re.compile(
    r"(?:"
    r"[A-Z]{2}\d{7}"             # 구여권: AB1234567
    r"|"
    r"(?:[MSRODG]\d{3}[A-Z]\d{4})"     # 신여권: M123A4567
    r")"
)

DRIVER_LICENSE_RE = re.compile(
    r"(?:\d{2}-\d{2}-\d{6}-\d{2}|\d{2}-\d{2}-\d{7}|\d{2}\d{2}\d{6}\d{2})"
)

# 마스킹 기본 설정
DEFAULT_MASK = "*"

# 패턴 테이블
RULES = [
    {"name": "rrn", "regex": RRN_RE, "validator": is_valid_rrn, "mask": "*"},
    {"name": "fgn", "regex": FGN_RE, "validator": is_valid_fgn, "mask": "*"},
    {"name": "email", "regex": EMAIL_RE, "validator": is_valid_email, "mask": "*"},
    {"name": "phone_mobile", "regex": MOBILE_RE, "validator": is_valid_phone_mobile, "mask": "*"},
    {"name": "phone_city", "regex": CITY_RE, "validator": is_valid_phone_city, "mask": "*"},
    {"name": "card", "regex": CARD_RE, "validator": is_valid_card, "mask": "*"},
    {"name": "passport", "regex": PASSPORT_RE, "validator": None, "mask": "*"},
    {"name": "driver_license", "regex": DRIVER_LICENSE_RE, "validator": is_valid_driver_license, "mask": "*"},
]
