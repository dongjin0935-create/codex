from __future__ import annotations

import re
from datetime import date, datetime


_DATE_PATTERNS = ("%Y-%m-%d", "%Y.%m.%d", "%Y/%m/%d", "%Y%m%d")


def normalize_date_text(text: str) -> str:
    value = re.sub(r"\s+", "", text)
    value = value.replace("년", "-").replace("월", "-").replace("일", "")
    value = value.replace(".", "-").replace("/", "-")
    value = re.sub(r"-+", "-", value).strip("-")
    return value


def parse_date(text: str) -> date:
    clean = normalize_date_text(text)
    for pattern in _DATE_PATTERNS:
        try:
            return datetime.strptime(clean if pattern != "%Y%m%d" else clean.replace("-", ""), pattern).date()
        except ValueError:
            continue
    raise ValueError(f"날짜 파싱 실패: {text}")


def date_to_compact(value: date) -> str:
    return value.strftime("%Y%m%d")
