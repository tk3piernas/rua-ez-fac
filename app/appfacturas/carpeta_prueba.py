"""The test folder (carpeta de prueba): a copy of the real 'Facturas pendientes' and workbook, plus a test batch.
Used by Configuración → "Practicar en la carpeta de prueba" and "Rearmar la carpeta de prueba"."""
from __future__ import annotations

import shutil
from pathlib import Path

from .config import carpeta_app, cargar

# real invoices used as fixtures: local only (never in the repo)
FIXTURES = next((p for p in (Path(__file__).resolve().parents[1] / "tests_privados" / "fixtures",
                             Path(__file__).resolve().parents[1] / "tests" / "fixtures") if p.is_dir()),
                Path(__file__).resolve().parents[1] / "tests_privados" / "fixtures")

# test batch (area; fixture; name in _Entrada). Lives with the fixtures: they are real invoices and stay off the repo.
def lote_prueba() -> list[tuple[str, str, str]]:
    f = FIXTURES / "lote_prueba.csv"
    if not f.is_file():
        return []
    filas = [l.split(";") for l in f.read_text(encoding="utf-8").splitlines()[1:] if l.strip()]
    return [(a, b, c) for a, b, c in filas if (FIXTURES / b).is_file()]


def carpeta_prueba() -> Path:
    return carpeta_app() / "prueba" / "OneDrive - Rua Asistencia"


def preparar() -> Path:
    """Rearmar la carpeta de prueba."""
    return rearmar()


def rearmar() -> Path:
    """Copy of the real Facturas pendientes (saved invoices + csv) and of the workbook, plus the test batch."""
    real = cargar()
    if not real.base:
        raise RuntimeError('No encuentro la carpeta "Facturas pendientes" del OneDrive de RUA.')
    dest = carpeta_prueba()
    if dest.exists():
        shutil.rmtree(dest)
    base = dest / "Facturas pendientes"
    for a in real.areas:
        if (real.base / a).is_dir():
            shutil.copytree(real.base / a, base / a)
        (base / "_Entrada" / a).mkdir(parents=True, exist_ok=True)
    (base / "_Script").mkdir(parents=True, exist_ok=True)
    if real.csv_proveedores.is_file():
        shutil.copy2(real.csv_proveedores, base / "_Script" / "proveedores.csv")
    if real.planilla and real.planilla.exists():
        shutil.copy2(real.planilla, dest / real.planilla.name)
    lote = lote_prueba()
    for area, fx, nombre in lote:
        shutil.copy2(FIXTURES / fx, base / "_Entrada" / area / nombre)
    print(f"Carpeta de prueba lista: {base}")
    print(f"Lote de prueba: {len(lote)} archivos en _Entrada")
    return base
