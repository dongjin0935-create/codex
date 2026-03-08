from __future__ import annotations

from abc import ABC, abstractmethod
from pathlib import Path

from src.models import OrderData


class TradeWriterBase(ABC):
    @abstractmethod
    def write(self, order: OrderData, mapping: dict, template_path: Path, output_dir: Path, ext: str) -> Path:
        raise NotImplementedError
