from __future__ import annotations

from pathlib import Path

from src.models import OrderData
from src.utils.file_utils import resolve_unique_path
from src.writers.trade_writer_base import TradeWriterBase

try:
    import win32com.client  # type: ignore
except Exception:  # pragma: no cover
    win32com = None


class TradeWriterWin32(TradeWriterBase):
    def write(self, order: OrderData, mapping: dict, template_path: Path, output_dir: Path, ext: str) -> Path:
        if win32com is None:
            raise RuntimeError("pywin32/Excel COM 사용 불가")

        output_name = f"강남거래명세서_{order.issue_date_compact}_{order.project_no}_{order.order_no}.{ext}"
        out_path = resolve_unique_path(output_dir / output_name)

        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        wb = excel.Workbooks.Open(str(template_path))
        try:
            file_format = 56 if ext.lower() == "xls" else 51
            wb.SaveAs(str(out_path), FileFormat=file_format)

            trade_map = mapping["trade"]
            for key, rc in trade_map.get("meta_cells", {}).items():
                sheet, pos = rc.split("!", 1)
                ws = wb.Worksheets(sheet)
                r = int(pos.split("C")[0][1:])
                c = int(pos.split("C")[1])
                if key == "project_no":
                    ws.Cells(r, c).Value = order.project_no
                elif key == "order_no":
                    ws.Cells(r, c).Value = order.order_no
                elif key == "issue_date":
                    ws.Cells(r, c).Value = order.issue_date.strftime("%Y-%m-%d")
                elif key == "due_date":
                    ws.Cells(r, c).Value = order.due_date.strftime("%Y-%m-%d") if order.due_date else ""
                elif key == "customer_name":
                    ws.Cells(r, c).Value = order.customer_name or "강남조선"

            item_idx = 0
            item_sheets = trade_map["item_sheets"]
            for spec in item_sheets:
                ws = wb.Worksheets(spec["sheet_name"])
                start = spec["start_row"]
                end = spec["end_row"]
                columns = spec["item_columns"]
                for row in range(start, end + 1):
                    if item_idx >= len(order.items):
                        break
                    item = order.items[item_idx]
                    ws.Cells(row, columns["순번"]).Value = item.seq
                    ws.Cells(row, columns["POR NO"]).Value = item.por_no
                    ws.Cells(row, columns["자재번호"]).Value = item.material_no
                    ws.Cells(row, columns["품명 및 규격"]).Value = item.description
                    ws.Cells(row, columns["단위"]).Value = item.unit
                    ws.Cells(row, columns["발주수량"]).Value = float(item.qty)
                    ws.Cells(row, columns["납품수량"]).Value = float(item.qty)
                    ws.Cells(row, columns["합격수량"]).Value = float(item.qty)
                    ws.Cells(row, columns["단가"]).Value = float(item.unit_price)
                    ws.Cells(row, columns["금액"]).Value = float(item.amount)
                    ws.Cells(row, columns["저장위치"]).Value = ""
                    item_idx += 1
            if item_idx < len(order.items):
                raise RuntimeError("을지 시트 용량 부족")

            wb.Save()
            return out_path
        finally:
            wb.Close(SaveChanges=True)
            excel.Quit()
