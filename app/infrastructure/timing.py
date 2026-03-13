from __future__ import annotations

from dataclasses import dataclass
from time import perf_counter
from contextlib import contextmanager
from typing import Iterator, Callable


@dataclass
class StageTiming:
    stage: str
    elapsed_ms: float


@contextmanager
def timed_stage(stage: str, collector: Callable[[StageTiming], None]) -> Iterator[None]:
    t0 = perf_counter()
    try:
        yield
    finally:
        collector(StageTiming(stage=stage, elapsed_ms=round((perf_counter() - t0) * 1000.0, 2)))
