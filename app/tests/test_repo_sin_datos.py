"""The repo is public: no tracked file may contain a valid CBU or a valid CUIT other than RUA's."""
import re
import subprocess
from pathlib import Path

import pytest

from appfacturas.util import cbu_valido, cuit_valido

RAIZ = Path(__file__).resolve().parents[2]
CUIT_RUA = "30717419428"
# invented on purpose to test the check-digit functions (test_util.py)
INVENTADOS = {"0110000612345678901233", "0000003110000000000014", "0720000798765432109877",
              "3840000955555555555555", "20111111112", "23123456785"}
BINARIOS = (".png", ".ico", ".jpg", ".pdf", ".xlsx", ".zip")


def _archivos():
    try:
        r = subprocess.run(["git", "ls-files", "--cached", "--others", "--exclude-standard"], cwd=RAIZ,
                           capture_output=True, text=True, check=True)
    except (OSError, subprocess.CalledProcessError):
        pytest.skip("no hay git")
    return [RAIZ / f for f in r.stdout.splitlines() if f and not f.lower().endswith(BINARIOS)]


def test_sin_cbu_ni_cuit_reales():
    hallados = []
    for f in _archivos():
        try:
            texto = f.read_text(encoding="utf-8", errors="ignore")
        except OSError:
            continue
        for m in re.finditer(r"(?<![\d-])(\d{22})(?!\d)", texto):
            if cbu_valido(m.group(1)) and m.group(1) not in INVENTADOS:
                hallados.append(f"{f.relative_to(RAIZ)}: CBU {m.group(1)}")
        for m in re.finditer(r"(?<![\d-])(\d{2})-?(\d{8})-?(\d)(?![\d-])", texto):
            c = "".join(m.groups())
            if cuit_valido(c) and c != CUIT_RUA and c not in INVENTADOS:
                hallados.append(f"{f.relative_to(RAIZ)}: CUIT {m.group(0)}")
    assert not hallados, "datos que no pueden ir al repo público:\n" + "\n".join(hallados)
