from __future__ import annotations

import logging
import re
from datetime import date
from decimal import Decimal
from pathlib import Path

from config.constants import ORDER_NO_PADDING
from src.models import ItemData, OrderData
from src.parser.filename_parser import parse_from_filename
from src.utils.date_utils import date_to_compact, parse_date
from src.utils.number_utils import to_decimal

logger = logging.getLogger(__name__)


class PdfParseError(RuntimeError):
    pass


def _extract_text(pdf_path: Path) -> str:
    try:
        import pdfplumber  # type: ignore
    except Exception as exc:
        raise PdfParseError("pdfplumber 패키지가 필요합니다.") from exc

    pages: list[str] = []
    with pdfplumber.open(str(pdf_path)) as pdf:
        for page in pdf.pages:
            pages.append(page.extract_text() or "")
    return "\n".join(pages)


def _find_first(patterns: list[re.Pattern[str]], text: str) -> str | None:
    for pattern in patterns:
        m = pattern.search(text)
        if m:
            return m.group(1)
    return None


def _parse_items(lines: list[str]) -> list[ItemData]:
    start = None
    for idx, line in enumerate(lines):
        if "POR" in line and "자재번호" in line and "품명" in line:
            start = idx + 1
            break
    if start is None:
        start = 0

    item_pattern = re.compile(
        r"^(?P<seq>\d+)\s+(?P<por>\S+)\s+(?P<mat>\S+)\s+(?P<desc>.*?)\s+(?P<unit>EA|SET|M|KG|L|BOX|대|개)\s+(?P<qty>[\d,\.]+)\s+(?P<price>[\d,\.]+)(?:\s+(?P<amount>[\d,\.]+))?(?:\s+(?P<due>\d{4}[./-]\d{1,2}[./-]\d{1,2}))?$"
    )

    items: list[ItemData] = []
    current: ItemData | None = None
    for line in lines[start:]:
        raw = " ".join(line.split())
        if not raw:
            continue
        matched = item_pattern.match(raw)
        if matched:
            data = matched.groupdict()
            qty = to_decimal(data["qty"])
            price = to_decimal(data["price"])
            amount = to_decimal(data.get("amount")) if data.get("amount") else qty * price
            current = ItemData(
                seq=int(data["seq"]),
                por_no=data["por"],
                material_no=data["mat"],
                description=data["desc"].strip(),
                unit=data["unit"],
                qty=qty,
                unit_price=price,
                amount=amount,
                due_date_text=data.get("due"),
            )
            items.append(current)
        elif current and not raw.startswith(("소계", "합계", "비고")):
            current.description = f"{current.description} {raw}".strip()
    return items


def parse_order_pdf(pdf_path: Path) -> OrderData:
    text = _extract_text(pdf_path)
    lines = [line.strip() for line in text.splitlines()]

    issue_date_text = _find_first([
        re.compile(r"발행일자\s*[:：]?\s*([0-9./-]+)"),
        re.compile(r"작성일자\s*[:：]?\s*([0-9./-]+)"),
    ], text)
    project_no = _find_first([re.compile(r"공사번호\s*[:：]?\s*(\d+)")], text)
    order_no = _find_first([re.compile(r"발주번호\s*[:：]?\s*(\d+)")], text)
    due_date_text = _find_first([re.compile(r"납기일\s*[:：]?\s*([0-9./-]+)")], text)

    filename_data = parse_from_filename(pdf_path)
    if filename_data:
        issue_date_text = issue_date_text or filename_data["issue_date"]
        project_no = project_no or filename_data["project_no"]
        order_no = order_no or filename_data["order_no"]

    if not issue_date_text:
        raise PdfParseError("PDF에서 발행일자를 찾지 못했습니다.")
    if not project_no:
        raise PdfParseError("PDF에서 공사번호를 찾지 못했습니다.")
    if not order_no:
        raise PdfParseError("PDF에서 발주번호를 찾지 못했습니다.")

    order_no = str(order_no).zfill(ORDER_NO_PADDING)
    issue_date: date = parse_date(issue_date_text)
    due_date = parse_date(due_date_text) if due_date_text else None

    items = _parse_items(lines)
    if not items:
        snippet = "\n".join(lines[:40])
        logger.error("품목 추출 실패. 원문 일부:\n%s", snippet)
        raise PdfParseError("PDF에서 품목을 1건도 찾지 못했습니다.")

    project_order_key = f"{project_no}-{order_no}"
    issue_compact = date_to_compact(issue_date)
    return OrderData(
        source_pdf_path=str(pdf_path),
        issue_date=issue_date,
        issue_date_text=issue_date_text,
        issue_date_compact=issue_compact,
        project_no=project_no,
        order_no=order_no,
        project_order_key=project_order_key,
        memo_key=f"{issue_compact}-{project_no}-{order_no}",
        due_date=due_date,
        due_date_text=due_date_text,
        supplier_name=filename_data.get("supplier") if filename_data else None,
        customer_name=None,
        items=items,
    )
