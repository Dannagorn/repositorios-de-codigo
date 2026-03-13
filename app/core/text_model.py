from __future__ import annotations

from dataclasses import dataclass
import re
import unicodedata


@dataclass(frozen=True)
class TextoProcessado:
    original: str
    normalizado: str
    uppercase: str
    linhas: tuple[str, ...]

    @staticmethod
    def _normalizar(texto: str) -> str:
        t = unicodedata.normalize("NFKD", texto or "").encode("ASCII", "ignore").decode("ASCII")
        linhas = [re.sub(r"\s+", " ", ln).strip() for ln in t.splitlines()]
        return "\n".join(linhas).strip()

    @classmethod
    def from_texto(cls, texto: str) -> "TextoProcessado":
        n = cls._normalizar(texto)
        return cls(original=texto or "", normalizado=n, uppercase=n.upper(), linhas=tuple(n.splitlines()))
