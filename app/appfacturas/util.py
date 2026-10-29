"""Small helpers shared by every module: amounts, dates, CUIT/CBU checks, names."""
from __future__ import annotations

import hashlib
import json
import os
import re
import unicodedata
from datetime import date, datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path
from urllib.parse import quote

# ---------------------------------------------------------------- amounts

_AMOUNT_RE = re.compile(r"-?\$?\s*-?\d[\d.,]*")


def parse_monto(s: str | None) -> Decimal | None:
    """Parse an Argentine or US formatted amount ("130.000,50", "130000,50", "$95,000.00")."""
    if s is None:
        return None
    t = str(s).strip().replace("$", "").replace(" ", "").replace(" ", "")
    if not t:
        return None
    neg = t.startswith("-") or (t.startswith("(") and t.endswith(")"))
    t = t.strip("-()")
    if not re.fullmatch(r"[\d.,]+", t) or not re.search(r"\d", t):
        return None
    last_c, last_p = t.rfind(","), t.rfind(".")
    if last_c >= 0 and last_p >= 0:
        dec = "," if last_c > last_p else "."          # the last separator is the decimal one
    elif last_c >= 0 or last_p >= 0:
        sep = "," if last_c >= 0 else "."
        tail = len(t) - max(last_c, last_p) - 1
        # one separator followed by 1-2 digits = decimal ("130000,00", "1.5"); 3 digits = thousands
        dec = sep if t.count(sep) == 1 and tail in (1, 2) else ""
    else:
        dec = ""
    if dec == ",":
        t = t.replace(".", "").replace(",", ".")
    elif dec == ".":
        t = t.replace(",", "")
    else:
        t = t.replace(",", "").replace(".", "")
    try:
        v = Decimal(t)
    except InvalidOperation:
        return None
    return -v if neg else v


def fmt_monto(v: Decimal | None) -> str:
    """$ 1.234,56 (Argentine format) for reports."""
    if v is None:
        return "sin monto"
    s = f"{abs(v):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return f"{'-' if v < 0 else ''}$ {s}"


# ---------------------------------------------------------------- dates

_MESES = {m: i + 1 for i, ms in enumerate([
    ("ene", "jan"), ("feb",), ("mar",), ("abr", "apr"), ("may",), ("jun",), ("jul",), ("ago", "aug"),
    ("sep", "set"), ("oct",), ("nov",), ("dic", "dec")]) for m in ms}


def parse_fecha(s: str | None) -> date | None:
    """dd/mm/yyyy (also - or .), yyyy-mm-dd, "jun 13, 2026", "13 de junio de 2026"."""
    if not s:
        return None
    try:
        m = re.search(r"(?<!\d)(\d{4})-(\d{2})-(\d{2})(?!\d)", s)                         # ISO 2026-10-22
        if m:
            return date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
        m = re.search(r"(\d{1,2})[/.-](\d{1,2})[/.-](\d{4}|\d{2})(?!\d)", s)
        if m:
            y = int(m.group(3))
            return date(y + 2000 if y < 100 else y, int(m.group(2)), int(m.group(1)))
        m = re.search(r"(?i)\b([a-z]{3})[a-z]*\.?\s+(\d{1,2}),?\s+(\d{4})", s)          # jun 13, 2026
        if m and m.group(1).lower() in _MESES:
            return date(int(m.group(3)), _MESES[m.group(1).lower()], int(m.group(2)))
        m = re.search(r"(?i)(\d{1,2})\s+de\s+([a-z]{3})[a-z]*\s+(?:de\s+)?(\d{4})", s)   # 13 de junio de 2026
        if m and m.group(2).lower() in _MESES:
            return date(int(m.group(3)), _MESES[m.group(2).lower()], int(m.group(1)))
    except ValueError:
        return None
    return None


def fmt_fecha(d: date | None) -> str:
    return d.strftime("%d/%m/%Y") if d else ""


# ---------------------------------------------------------------- CUIT / CBU

def solo_digitos(s: str | None) -> str:
    return re.sub(r"\D", "", s or "")


def cuit_valido(c: str | None) -> bool:
    c = solo_digitos(c)
    if len(c) != 11:
        return False
    w = (5, 4, 3, 2, 7, 6, 5, 4, 3, 2)
    r = 11 - sum(int(a) * b for a, b in zip(c[:10], w)) % 11
    r = {11: 0, 10: 9}.get(r, r)
    return r == int(c[10])


def cbu_valido(c: str | None) -> bool:
    c = solo_digitos(c)
    if len(c) != 22:
        return False
    d = [int(x) for x in c]
    w1 = (7, 1, 3, 9, 7, 1, 3)
    w2 = (3, 9, 7, 1, 3, 9, 7, 1, 3, 9, 7, 1, 3)
    if (10 - sum(a * b for a, b in zip(d[:7], w1)) % 10) % 10 != d[7]:
        return False
    return (10 - sum(a * b for a, b in zip(d[8:21], w2)) % 10) % 10 == d[21]


# ---------------------------------------------------------------- names

def norm(s: str | None) -> str:
    """Lowercase, no accents, only [a-z0-9] separated by single spaces (PS1 Norm)."""
    if not s:
        return ""
    d = unicodedata.normalize("NFD", s)
    d = "".join(ch for ch in d if unicodedata.category(ch) != "Mn")
    return re.sub(r"[^a-z0-9]+", " ", d.lower()).strip()


def limpiar_nombre(s: str) -> str:
    """Safe folder/file name (PS1 Limpiar-Nombre)."""
    return re.sub(r'[\\/:*?"<>|]', "-", s).strip().rstrip(".")


def ruta_libre(p: Path, ocupadas: set[str] | None = None) -> Path:
    """p, or "name (2).ext", "name (3).ext"... if it exists (on disk or in `ocupadas`)."""
    ocupadas = ocupadas or set()

    def libre(x: Path) -> bool:
        return not x.exists() and str(x).lower() not in ocupadas

    if libre(p):
        return p
    i = 2
    while True:
        q = p.with_name(f"{p.stem} ({i}){p.suffix}")
        if libre(q):
            return q
        i += 1


def url_sharepoint(base: str, *partes: str) -> str:
    """Same as the PS1: base with spaces as %20 + each part with EscapeDataString."""
    return base.replace(" ", "%20").rstrip("/") + "/" + "/".join(quote(p, safe="") for p in partes)


def guardar_json(ruta: Path, datos, default=None) -> None:
    """Writes a record so that a cut (app closed, PC off) never leaves it half written: a temporary file,
    flushed to disk, then swapped in. The journal and the states depend on it to undo."""
    ruta = Path(ruta)
    ruta.parent.mkdir(parents=True, exist_ok=True)
    tmp = ruta.with_name(ruta.name + ".tmp")
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(datos, f, ensure_ascii=False, indent=1, default=default)
        f.flush()
        os.fsync(f.fileno())
    os.replace(tmp, ruta)


def sha256(p: Path) -> str:
    h = hashlib.sha256()
    with open(p, "rb") as f:
        for b in iter(lambda: f.read(1 << 20), b""):
            h.update(b)
    return h.hexdigest().upper()


def ahora_id() -> str:
    return datetime.now().strftime("%Y-%m-%d_%H%M%S")


def relativa(p, base) -> str:
    """`p` relative to `base` for messages ("Vehicular\\Prov\\Factura ..."), or the full path if not inside."""
    try:
        return str(Path(p).relative_to(base))
    except (ValueError, TypeError):
        return str(p)
