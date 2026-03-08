from pathlib import Path

import pytest

from src.analyzer.erp_analyzer import analyze_erp_template

openpyxl = pytest.importorskip("openpyxl")
Workbook = openpyxl.Workbook


def test_erp_analyzer_header_map(tmp_path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "주문서입력"
    headers = ["일자", "순번", "거래처코드", "거래처명", "납기일자", "프로젝트명", "품목명", "규격", "수량", "단가", "공급가액", "부가세", "적요"]
    for i, h in enumerate(headers, start=1):
        ws.cell(row=3, column=i, value=h)
        ws.cell(row=4, column=i, value="SAMPLE")
    path = tmp_path / "erp.xlsx"
    wb.save(path)

    result = analyze_erp_template(path)
    assert result.header_row == 3
    assert "일자" in result.header_map


def test_generated_mapping_json_creation(tmp_path: Path) -> None:
    config = tmp_path / "generated_mapping.json"
    config.write_text('{"ok": true}', encoding="utf-8")
    assert config.exists()
