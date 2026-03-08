from __future__ import annotations

import logging
from pathlib import Path

from src.utils.file_utils import ensure_dir


def setup_logging(output_dir: Path, debug: bool = False) -> None:
    logs_dir = ensure_dir(output_dir / "logs")
    level = logging.DEBUG if debug else logging.INFO
    handlers: list[logging.Handler] = [logging.StreamHandler()]
    handlers.append(logging.FileHandler(logs_dir / "run.log", encoding="utf-8"))
    logging.basicConfig(
        level=level,
        format="%(asctime)s [%(levelname)s] %(name)s - %(message)s",
        handlers=handlers,
        force=True,
    )
