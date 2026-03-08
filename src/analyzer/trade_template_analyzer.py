from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from config.constants import TRADE_ITEM_HEADERS, TRADE_META_LABELS

try:
    import win32com.client  # type: ignore
except Exception:  # pragma: no cover
    win32com = None


@dataclass(slots=True)
class TradeSheetCapacity:
    sheet_name: str
    start_row: int
    end_row: int
    max_rows: int
    item_columns: dict[str, int]


@dataclass(slots=True)
class TradeAnalysis:
    sheet_names: list[str]
    role_sheets: dict[str, list[str]]
    meta_cells: dict[str, str]
    item_sheets: list[TradeSheetCapacity]


def _is_footer(text: str) -> bool:
    return any(key in text for key in ("소계", "합계", "비고"))


def analyze_trade_template(template_path: Path) -> TradeAnalysis:
    if win32com is None:
        raise RuntimeError("pywin32/Excel COM 사용 불가: 거래명세서 템플릿 분석 실패")

    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    wb = excel.Workbooks.Open(str(template_path))
    try:
        sheet_names = [ws.Name for ws in wb.Worksheets]
        role_sheets = {"갑지": [], "을지": [], "공급자정보": []}
        meta_cells: dict[str, str] = {}
        item_sheets: list[TradeSheetCapacity] = []

        for ws in wb.Worksheets:
            used = ws.UsedRange
            values = used.Value
            texts = []
            if isinstance(values, tuple):
                for row in values:
                    if isinstance(row, tuple):
                        for c in row:
                            if c:
                                texts.append(str(c))

            joined = " ".join(texts)
            for role in role_sheets:
                if role in ws.Name or role in joined:
                    role_sheets[role].append(ws.Name)

            for key, labels in TRADE_META_LABELS.items():
                if key in meta_cells:
                    continue
                found = ws.Cells.Find(What=labels[0], LookAt=1)
                if found is not None:
                    target = ws.Cells(found.Row, found.Column + 1)
                    meta_cells[key] = f"{ws.Name}!R{target.Row}C{target.Column}"

            header_row = None
            item_columns: dict[str, int] = {}
            for r in range(1, min(200, used.Rows.Count + 1)):
                hits = {}
                for c in range(1, min(80, used.Columns.Count + 1)):
                    text = str(ws.Cells(r, c).Value or "").strip()
                    if text in TRADE_ITEM_HEADERS:
                        hits[text] = c
                if all(h in hits for h in TRADE_ITEM_HEADERS):
                    header_row = r
                    item_columns = hits
                    break

            if header_row:
                start_row = header_row + 1
                end_row = start_row
                for r in range(start_row, start_row + 400):
                    row_text = " ".join(str(ws.Cells(r, c).Value or "") for c in range(1, min(30, used.Columns.Count + 1)))
                    if _is_footer(row_text):
                        end_row = r - 1
                        break
                    end_row = r
                item_sheets.append(
                    TradeSheetCapacity(
                        sheet_name=ws.Name,
                        start_row=start_row,
                        end_row=end_row,
                        max_rows=max(0, end_row - start_row + 1),
                        item_columns=item_columns,
                    )
                )

        if not role_sheets["을지"] and item_sheets:
            role_sheets["을지"] = [s.sheet_name for s in item_sheets]

        if not item_sheets:
            raise RuntimeError("거래명세서 템플릿에서 품목 헤더를 찾지 못했습니다.")

        return TradeAnalysis(
            sheet_names=sheet_names,
            role_sheets=role_sheets,
            meta_cells=meta_cells,
            item_sheets=item_sheets,
        )
    finally:
        wb.Close(False)
        excel.Quit()
