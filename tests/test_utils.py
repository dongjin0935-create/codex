from decimal import Decimal

from src.utils.date_utils import parse_date
from src.utils.number_utils import to_decimal


def test_date_normalization() -> None:
    assert parse_date("2026년 01월 29일").strftime("%Y-%m-%d") == "2026-01-29"


def test_number_to_decimal() -> None:
    assert to_decimal("1,234.50") == Decimal("1234.50")
