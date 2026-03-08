from pathlib import Path

import src.parser.pdf_parser as pdf_parser


def test_pdf_parser_extract_items(monkeypatch) -> None:
    sample_text = """
    발행일자: 2026-01-29
    공사번호: 224001
    발주번호: 154
    납기일: 2026-02-10
    순번 POR NO 자재번호 품명 및 규격 단위 수량 단가 금액 납기일
    1 POR01 MAT01 볼트 M10 EA 10 1000 10000 2026-02-10
    """

    monkeypatch.setattr(pdf_parser, "_extract_text", lambda _: sample_text)
    order = pdf_parser.parse_order_pdf(Path("20260129-224001-0154-동진산업.pdf"))
    assert order.project_no == "224001"
    assert order.order_no == "0154"
    assert len(order.items) >= 1
