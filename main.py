from __future__ import annotations

import argparse
import json
import logging
import traceback
from pathlib import Path

from src.analyzer.mapping_builder import build_and_save_mapping, load_mapping
from src.parser.pdf_parser import parse_order_pdf
from src.utils.file_utils import ensure_dir
from src.utils.logging_utils import setup_logging
from src.writers.erp_writer import write_erp_file
from src.writers.trade_writer_openpyxl import TradeWriterOpenpyxl
from src.writers.trade_writer_win32 import TradeWriterWin32

logger = logging.getLogger(__name__)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="강남조선 발주서 자동 변환 프로그램")
    parser.add_argument("--pdf", required=True)
    parser.add_argument("--erp-template", required=True)
    parser.add_argument("--trade-template", required=True)
    parser.add_argument("--output-dir", required=True)
    parser.add_argument("--config", default="config/generated_mapping.json")
    parser.add_argument("--trade-ext", choices=["xls", "xlsx"], default="xls")
    parser.add_argument("--analyze-only", action="store_true")
    parser.add_argument("--debug", action="store_true")
    return parser.parse_args()


def run() -> int:
    args = parse_args()

    pdf_path = Path(args.pdf)
    erp_template = Path(args.erp_template)
    trade_template = Path(args.trade_template)
    output_dir = ensure_dir(Path(args.output_dir))
    config_path = Path(args.config)

    setup_logging(output_dir, debug=args.debug)

    try:
        if not config_path.exists() or args.analyze_only:
            logger.info("샘플 파일 분석 시작")
            mapping = build_and_save_mapping(erp_template, trade_template, config_path)
            logger.info("매핑 생성 완료: %s", config_path)
            if args.debug:
                logger.debug("매핑 상세:\n%s", json.dumps(mapping, ensure_ascii=False, indent=2))
            if args.analyze_only:
                return 0

        mapping = load_mapping(config_path)
        order = parse_order_pdf(pdf_path)

        erp_out = write_erp_file(order, mapping, erp_template, output_dir)
        logger.info("ERP 파일 생성 완료: %s", erp_out)

        try:
            trade_writer = TradeWriterWin32()
            trade_out = trade_writer.write(order, mapping, trade_template, output_dir, args.trade_ext)
        except Exception as exc:
            logger.warning("Win32 작성 실패, openpyxl 백엔드 시도: %s", exc)
            trade_writer = TradeWriterOpenpyxl()
            trade_out = trade_writer.write(order, mapping, trade_template, output_dir, "xlsx")

        logger.info("거래명세서 생성 완료: %s", trade_out)
        logger.info("모든 작업이 완료되었습니다.")
        return 0
    except Exception as exc:
        logger.error("실행 실패: %s", exc)
        if args.debug:
            traceback.print_exc()
        return 1


if __name__ == "__main__":
    raise SystemExit(run())
