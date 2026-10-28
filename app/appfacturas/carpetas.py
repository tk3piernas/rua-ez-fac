"""Index of the invoices already saved in Vehicular\\* and Hogar\\* (PS1 Indexar-Carpetas)."""
from __future__ import annotations

import re
from collections import defaultdict
from pathlib import Path

from .parsers.base import nombre_afip, tipo_de_codigo
from .util import sha256


def clave(cuit: str, tipo: str, pv, nro) -> str:
    return f"{cuit}|{tipo}|{int(pv)}|{int(nro)}"


def clave_nf(prov: str, tipo: str, numero: str) -> str:
    """Key of a non-fiscal receipt ("FAC0297"): provider key + its own number."""
    return f"{prov}|{tipo}|{numero.upper()}"


def _tipo_nombre(t: str) -> str:
    return "Nota de Crédito" if t.startswith("Nota de Cr") else "Nota de Débito" if t.startswith("Nota de D") else "Factura"


class Indice:
    def __init__(self, base: Path, areas, cuit_por_carpeta: dict[str, str]):
        self.base = base
        self.por_clave: dict[str, list[Path]] = defaultdict(list)
        self.por_tamano: dict[int, list[Path]] = defaultdict(list)
        self.sin_clave: list[Path] = []
        self._hash: dict[Path, str] = {}
        self.cuit_por_carpeta = cuit_por_carpeta
        self.total = 0
        for a in areas:
            d = base / a
            if not d.is_dir():
                continue
            for f in d.rglob("*"):
                if not f.is_file() or f.suffix.lower() != ".pdf":
                    continue
                self.total += 1
                k = self.clave_de_archivo(f)
                if k:
                    self.por_clave[k].append(f)
                else:
                    self.sin_clave.append(f)
                self.por_tamano[f.stat().st_size].append(f)

    def clave_de_archivo(self, f: Path) -> str | None:
        afip = nombre_afip(f.name)
        if afip:
            cu, cod, pv, nro = afip
            return clave(cu, tipo_de_codigo(cod)[1], pv, nro)
        cu = self.cuit_por_carpeta.get(f.parent.name.lower())
        m = re.match(r"^(Factura|Nota de Cr\S+|Nota de D\S+) (\d{4,5})-(\d{8})", f.stem)
        if m and cu:
            return clave(cu, _tipo_nombre(m.group(1)), m.group(2), m.group(3))
        m = re.match(r"^(Factura|Nota de Cr\S+|Nota de D\S+) ([A-Za-z]{1,5}-?\d{1,10})(?: \(\d+\))?$", f.stem)
        if m and cu:
            return clave_nf(cu, _tipo_nombre(m.group(1)), m.group(2))
        return None

    def hash_de(self, p: Path) -> str:
        if p not in self._hash:
            self._hash[p] = sha256(p)
        return self._hash[p]

    def agregar(self, k: str | None, p: Path, tamano: int, h: str) -> None:
        """A file this batch will save (so later files in the batch see it)."""
        if k:
            self.por_clave[k].append(p)
        self.por_tamano[tamano].append(p)
        self._hash[p] = h

    def buscar(self, k: str | None, tamano: int, h: str, excluir: Path | None = None) -> Path | None:
        if k and self.por_clave.get(k):
            return self.por_clave[k][0]
        for x in self.por_tamano.get(tamano, []):
            if x != excluir and self.hash_de(x) == h:
                return x
        return None

    def rel(self, p: Path) -> str:
        try:
            return str(p.relative_to(self.base))
        except ValueError:
            return str(p)
