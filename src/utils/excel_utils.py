from __future__ import annotations

from collections import defaultdict
from copy import copy
from typing import Iterable

from typing import Any

Worksheet = Any


def normalize_header(value: object) -> str:
    return str(value or "").strip()


def find_header_row(ws: Worksheet, required_headers: Iterable[str], max_row: int = 20) -> int:
    target = set(required_headers)
    for row in range(1, max_row + 1):
        row_values = {normalize_header(ws.cell(row=row, column=col).value) for col in range(1, ws.max_column + 1)}
        if target.issubset(row_values):
            return row
    raise ValueError("헤더 행을 찾을 수 없습니다.")


def build_header_map(ws: Worksheet, header_row: int) -> dict[str, list[int]]:
    mapping: dict[str, list[int]] = defaultdict(list)
    for col in range(1, ws.max_column + 1):
        label = normalize_header(ws.cell(row=header_row, column=col).value)
        if label:
            mapping[label].append(col)
    return dict(mapping)


def find_first_data_row(ws: Worksheet, header_row: int) -> int:
    for row in range(header_row + 1, ws.max_row + 1):
        if any(ws.cell(row=row, column=col).value not in (None, "") for col in range(1, ws.max_column + 1)):
            return row
    return header_row + 1


def copy_row_style(ws: Worksheet, from_row: int, to_row: int) -> None:
    for col in range(1, ws.max_column + 1):
        src = ws.cell(from_row, col)
        dst = ws.cell(to_row, col)
        if src.has_style:
            dst._style = copy(src._style)
        dst.number_format = src.number_format
        dst.alignment = copy(src.alignment)
        dst.font = copy(src.font)
        dst.fill = copy(src.fill)
        dst.border = copy(src.border)
        dst.protection = copy(src.protection)
