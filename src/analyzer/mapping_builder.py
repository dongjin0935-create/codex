from __future__ import annotations

import json
from dataclasses import asdict
from pathlib import Path

from src.analyzer.erp_analyzer import analyze_erp_template
from src.analyzer.trade_template_analyzer import analyze_trade_template


def build_and_save_mapping(erp_template: Path, trade_template: Path, output_config: Path) -> dict:
    erp = analyze_erp_template(erp_template)
    trade = analyze_trade_template(trade_template)

    mapping = {
        "erp": asdict(erp),
        "trade": {
            "sheet_names": trade.sheet_names,
            "role_sheets": trade.role_sheets,
            "meta_cells": trade.meta_cells,
            "item_sheets": [asdict(s) for s in trade.item_sheets],
        },
        "file_naming": {
            "trade": "강남거래명세서_{yyyymmdd}_{project_no}_{order_no}",
            "erp": "ERP업로드_{yyyymmdd}_{project_no}_{order_no}.xlsx",
        },
    }

    output_config.parent.mkdir(parents=True, exist_ok=True)
    output_config.write_text(json.dumps(mapping, ensure_ascii=False, indent=2), encoding="utf-8")
    return mapping


def load_mapping(config_path: Path) -> dict:
    return json.loads(config_path.read_text(encoding="utf-8"))
