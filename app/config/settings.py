from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
import os


@dataclass(frozen=True)
class OCRConfig:
    dpi: int = 220
    first_page: int = 1
    last_page: int = 3
    lang: str = "por"
    force_tesseract_cmd: str = ""
    force_poppler_path: str = ""


@dataclass(frozen=True)
class PipelineConfig:
    max_workers: int = max(1, (os.cpu_count() or 2) - 1)
    enable_sqlite_cache: bool = True
    enable_pdf_text_cache: bool = True


@dataclass(frozen=True)
class AppConfig:
    ocr: OCRConfig = OCRConfig()
    pipeline: PipelineConfig = PipelineConfig()
    app_name: str = "Organizador de Comprovantes"


def load_app_config() -> AppConfig:
    dpi = int(os.getenv("ORG_OCR_DPI", "220"))
    last_page = int(os.getenv("ORG_OCR_LAST_PAGE", "3"))
    force_tess = os.getenv("ORG_TESSERACT_CMD", "").strip()
    force_poppler = os.getenv("ORG_POPPLER_PATH", "").strip()
    max_workers = int(os.getenv("ORG_MAX_WORKERS", str(max(1, (os.cpu_count() or 2) - 1))))
    return AppConfig(
        ocr=OCRConfig(dpi=dpi, first_page=1, last_page=max(1, last_page), force_tesseract_cmd=force_tess, force_poppler_path=force_poppler),
        pipeline=PipelineConfig(max_workers=max_workers),
    )
