"""The .bat files of the repo: every 'goto :x' / 'call :x' has its label, and each label is defined once
(a missing label stopped the installer on a PC without Python; a repeated one made it loop forever)."""
import re
from collections import Counter
from pathlib import Path

import pytest

RAIZ = Path(__file__).resolve().parents[2]
BATS = sorted(RAIZ.glob("*.bat")) + sorted((RAIZ / "docs").rglob("*.bat"))


@pytest.mark.parametrize("bat", BATS, ids=lambda p: p.name)
def test_etiquetas(bat):
    texto = bat.read_bytes().decode("utf-8", "replace")
    etiquetas = Counter(m.group(1).lower() for m in re.finditer(r"(?m)^\s*:([A-Za-z_][\w]*)\s*$", texto))
    saltos = {m.group(1).lower() for m in re.finditer(r"(?i)\b(?:goto|call)\s+:([A-Za-z_]\w*)", texto)} - {"eof"}
    assert not [e for e, n in etiquetas.items() if n > 1], f"etiquetas repetidas en {bat.name}"
    assert not saltos - set(etiquetas), f"{bat.name} salta a etiquetas que no existen: {saltos - set(etiquetas)}"


@pytest.mark.parametrize("bat", BATS, ids=lambda p: p.name)
def test_sin_bom_y_con_crlf(bat):
    b = bat.read_bytes()
    assert not b.startswith(b"\xef\xbb\xbf"), "un BOM rompe la primera línea del .bat"
    assert b.count(b"\n") == b.count(b"\r\n"), "cmd.exe necesita CRLF para leer bien las etiquetas"
