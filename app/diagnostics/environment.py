from __future__ import annotations

from dataclasses import dataclass, asdict
from pathlib import Path
import importlib.util
import shutil
import sys
import tempfile

from app.infrastructure.paths import is_frozen, app_root, data_path


@dataclass
class EnvironmentDiagnostics:
    frozen: bool
    python: str
    app_root: str
    data_path: str
    tesseract: str
    pdftoppm: str
    writable_data_dir: bool
    missing_python_modules: list[str]


def _check_writable(path: Path) -> bool:
    try:
        path.mkdir(parents=True, exist_ok=True)
        with tempfile.NamedTemporaryFile(dir=path, delete=True):
            pass
        return True
    except Exception:
        return False


def run_environment_diagnostics() -> dict[str, object]:
    req = ["pypdf", "PyPDF2", "pytesseract", "pdf2image", "openpyxl", "tkinter"]
    missing = [m for m in req if importlib.util.find_spec(m) is None]
    diag = EnvironmentDiagnostics(
        frozen=is_frozen(),
        python=sys.executable,
        app_root=str(app_root()),
        data_path=str(data_path()),
        tesseract=shutil.which("tesseract") or "",
        pdftoppm=shutil.which("pdftoppm") or "",
        writable_data_dir=_check_writable(data_path()),
        missing_python_modules=missing,
    )
    return asdict(diag)
