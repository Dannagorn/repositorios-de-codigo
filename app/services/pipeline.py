from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable, Any

from app.infrastructure.timing import StageTiming, timed_stage


@dataclass
class PipelineRunResult:
    pdf_path: Path
    payload: dict[str, Any]
    timings: list[StageTiming] = field(default_factory=list)


@dataclass
class PipelineDeps:
    read_pdf_text: Callable[[Path], tuple[str, str, bool]]
    classify_and_extract: Callable[[str], dict[str, str]]


class ReceiptPipeline:
    def __init__(self, deps: PipelineDeps):
        self.deps = deps

    def process_file(self, pdf_path: Path) -> PipelineRunResult:
        timings: list[StageTiming] = []
        collector = timings.append

        with timed_stage("read_pdf_text", collector):
            texto, estrategia_leitura, usou_ocr = self.deps.read_pdf_text(pdf_path)

        with timed_stage("classify_extract", collector):
            dados = self.deps.classify_and_extract(texto)

        dados["_texto_fonte"] = texto
        dados["estrategia_leitura"] = estrategia_leitura
        dados["usou_ocr"] = str(bool(usou_ocr))
        return PipelineRunResult(pdf_path=pdf_path, payload=dados, timings=timings)
