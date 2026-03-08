from pathlib import Path

from src.parser.filename_parser import parse_from_filename


def test_parse_from_filename() -> None:
    data = parse_from_filename(Path("20260129-224001-154-동진산업.pdf"))
    assert data is not None
    assert data["issue_date"] == "20260129"
    assert data["project_no"] == "224001"
    assert data["order_no"] == "0154"
