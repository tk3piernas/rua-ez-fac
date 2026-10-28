"""Parser registry. Each module in this package exposing PARSER is registered automatically:
a new invoice format = a new file here + its test, no changes to the core."""
from __future__ import annotations

import importlib
import pkgutil

from .base import Parser

_REGISTRO: list[Parser] = []


def registrados() -> list[Parser]:
    if not _REGISTRO:
        for m in pkgutil.iter_modules(__path__):
            if m.name.startswith("_") or m.name == "base":
                continue
            mod = importlib.import_module(f"{__name__}.{m.name}")
            p = getattr(mod, "PARSER", None)
            if isinstance(p, Parser):
                _REGISTRO.append(p)
    return _REGISTRO


def elegir(doc, nombre_archivo: str) -> tuple[Parser | None, float]:
    mejor, puntaje = None, 0.0
    for p in registrados():
        try:
            s = p.detectar(doc, nombre_archivo)
        except Exception:
            s = 0.0
        if s > puntaje:
            mejor, puntaje = p, s
    return mejor, puntaje
