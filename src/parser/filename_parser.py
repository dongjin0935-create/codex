from __future__ import annotations

import re
from pathlib import Path

from config.constants import ORDER_NO_PADDING


FILENAME_PATTERN = re.compile(
    r"(?P<issue_date>\d{8})-(?P<project_no>\d+)-(?P<order_no>\d+)-(?P<supplier>.+)\.pdf$",
    re.IGNORECASE,
)


def parse_from_filename(pdf_path: Path) -> dict[str, str] | None:
    match = FILENAME_PATTERN.search(pdf_path.name)
    if not match:
        return None
    data = match.groupdict()
    data["order_no"] = data["order_no"].zfill(ORDER_NO_PADDING)
    return data
