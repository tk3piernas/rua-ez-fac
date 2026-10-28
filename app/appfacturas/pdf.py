"""PDF reading with PyMuPDF. Opens ARCA invoices (RC4, empty password, copy disabled)."""
from __future__ import annotations

import re
import zipfile
from dataclasses import dataclass, field
from pathlib import Path

import pymupdf

_COPIA_RE = re.compile(r"\b(D\s*U\s*P\s*L\s*I\s*C\s*A\s*D\s*O|T\s*R\s*I\s*P\s*L\s*I\s*C\s*A\s*D\s*O|CUADRUPLICADO)\b")


@dataclass
class Span:
    x0: float
    y0: float
    x1: float
    y1: float
    texto: str
    size: float = 0.0

    @property
    def yc(self) -> float:
        return (self.y0 + self.y1) / 2


@dataclass
class Pagina:
    numero: int
    texto: str
    spans: list[Span] = field(default_factory=list)
    ancho: float = 0.0
    alto: float = 0.0


@dataclass
class DocPdf:
    ruta: Path
    paginas: list[Pagina]            # only the ORIGINAL copy
    n_paginas: int = 0
    error: str = ""

    @property
    def texto(self) -> str:
        return "\n".join(p.texto for p in self.paginas)

    @property
    def tiene_texto(self) -> bool:
        return len(re.sub(r"\s", "", self.texto)) > 40


def _spans(page) -> list[Span]:
    out = []
    for b in page.get_text("dict")["blocks"]:
        for line in b.get("lines", []):
            for s in line["spans"]:
                t = s["text"].strip()
                if t:
                    x0, y0, x1, y1 = s["bbox"]
                    out.append(Span(x0, y0, x1, y1, t, s.get("size", 0)))
    out.sort(key=lambda s: (round(s.yc), s.x0))
    return out


def leer(ruta: Path) -> DocPdf:
    """Pages of the ORIGINAL copy: stops at the first page whose header says DUPLICADO/TRIPLICADO."""
    try:
        doc = pymupdf.open(ruta)
    except Exception as e:  # damaged or not a PDF
        return DocPdf(ruta, [], 0, f"no se pudo abrir: {e}")
    try:
        if doc.needs_pass and not doc.authenticate(""):
            return DocPdf(ruta, [], doc.page_count, "PDF con contraseña")
        paginas = []
        for i, page in enumerate(doc):
            texto = page.get_text()
            spans = _spans(page)
            # the copy mark is printed at the top (ARCA) or at the bottom (some billing software, "D U P L I C A D O")
            borde = " ".join(s.texto for s in spans if s.y0 < page.rect.height * 0.12 or s.y0 > page.rect.height * 0.88)
            if i > 0 and _COPIA_RE.search(borde):
                break
            paginas.append(Pagina(i + 1, texto, spans, page.rect.width, page.rect.height))
        return DocPdf(ruta, paginas, doc.page_count)
    finally:
        doc.close()


def imagen_pagina(ruta: Path, pagina: int = 0, dpi: int = 150) -> bytes:
    """PNG of one page (for scanned invoices sent to Claude)."""
    with pymupdf.open(ruta) as doc:
        return doc[pagina].get_pixmap(dpi=dpi).tobytes("png")


def expandir_zip(zip_path: Path, destino: Path) -> list[Path]:
    """Extract the PDFs of a zip (recursively, safely) and return them sorted."""
    destino.mkdir(parents=True, exist_ok=True)
    pdfs = []
    with zipfile.ZipFile(zip_path) as z:
        for info in z.infolist():
            if info.is_dir() or not info.filename.lower().endswith(".pdf"):
                continue
            nombre = Path(info.filename).name          # no paths from inside the zip
            if not nombre:
                continue
            out = destino / nombre
            i = 2
            while out.exists():
                out = destino / f"{Path(nombre).stem} ({i}).pdf"
                i += 1
            with z.open(info) as src, open(out, "wb") as dst:
                dst.write(src.read())
            pdfs.append(out)
    return sorted(pdfs, key=lambda p: p.name.lower())
