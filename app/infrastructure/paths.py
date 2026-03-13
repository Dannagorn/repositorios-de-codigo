from __future__ import annotations

from pathlib import Path
import sys


def is_frozen() -> bool:
    return bool(getattr(sys, "frozen", False))


def app_root() -> Path:
    if is_frozen():
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parents[2]


def resource_path(relative: str) -> Path:
    base = Path(getattr(sys, "_MEIPASS", app_root()))
    return (base / relative).resolve()


def data_path() -> Path:
    p = Path.home() / ".organizador_comprovantes"
    p.mkdir(parents=True, exist_ok=True)
    return p


def logs_path() -> Path:
    p = data_path() / "logs"
    p.mkdir(parents=True, exist_ok=True)
    return p
