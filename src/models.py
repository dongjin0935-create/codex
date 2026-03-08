from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date
from decimal import Decimal


@dataclass(slots=True)
class ItemData:
    seq: int
    por_no: str
    material_no: str
    description: str
    unit: str
    qty: Decimal
    unit_price: Decimal
    amount: Decimal
    due_date_text: str | None = None


@dataclass(slots=True)
class OrderData:
    source_pdf_path: str
    issue_date: date
    issue_date_text: str
    issue_date_compact: str
    project_no: str
    order_no: str
    project_order_key: str
    memo_key: str
    due_date: date | None
    due_date_text: str | None
    supplier_name: str | None
    customer_name: str | None
    items: list[ItemData] = field(default_factory=list)
