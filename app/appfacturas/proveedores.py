"""Provider registry: _Script\\proveedores.csv (shared with the PS1) + Proveedores sheet.

Key = CUIT. Providers without CUIT (non-fiscal receipts) use "SC:<normalised name>";
the PS1 ignores those lines because their CUIT column is empty."""
from __future__ import annotations

import csv
import re
from dataclasses import dataclass
from pathlib import Path

from .util import limpiar_nombre, norm, solo_digitos


@dataclass
class Proveedor:
    clave: str            # CUIT or "SC:<name>"
    area: str
    nombre: str
    carpeta: str

    @property
    def cuit(self) -> str:
        return "" if self.clave.startswith("SC:") else self.clave


def clave_sin_cuit(nombre: str) -> str:
    return "SC:" + norm(nombre)


_SIGLAS = {"srl", "sa", "sas", "sh", "s.a.", "s.r.l.", "s.a.s.", "ii", "iii", "mg", "mmt", "stm", "sos", "arc", "mdq", "abc"}


def nombre_prolijo(s: str) -> str:
    """'AUXILIO REMOLQUES PINAMAR' -> 'Auxilio Remolques Pinamar' (keeps SRL, SA...)."""
    s = re.sub(r"\s+", " ", s or "").strip()
    if not s or not s.isupper():
        return s
    return " ".join(w.upper() if w.lower() in _SIGLAS else w.capitalize() for w in s.split(" "))


class Registro:
    def __init__(self, ruta_csv: Path | None):
        self.ruta = ruta_csv
        self.por_clave: dict[str, Proveedor] = {}
        self.lineas_crudas: list[dict] = []
        if ruta_csv and ruta_csv.exists():
            with open(ruta_csv, encoding="utf-8-sig", newline="") as f:
                for r in csv.DictReader(f, delimiter=";"):
                    self.lineas_crudas.append(r)
                    c = solo_digitos(r.get("CUIT"))
                    nombre = (r.get("Proveedor") or "").strip()
                    clave = c if len(c) == 11 else (clave_sin_cuit(nombre) if nombre else "")
                    if clave:
                        self.por_clave[clave] = Proveedor(clave, (r.get("Area") or "").strip(), nombre,
                                                          (r.get("Carpeta") or "").strip())
        self.nuevos: dict[str, Proveedor] = {}

    @property
    def por_cuit(self) -> dict[str, Proveedor]:
        return {k: v for k, v in self.por_clave.items() if not k.startswith("SC:")}

    def buscar(self, clave: str) -> Proveedor | None:
        return self.por_clave.get(clave) or self.nuevos.get(clave)

    def buscar_nombre(self, nombre: str) -> Proveedor | None:
        n = norm(nombre)
        if not n:
            return None
        for p in list(self.por_clave.values()) + list(self.nuevos.values()):
            if n in (norm(p.nombre), norm(p.carpeta)):
                return p
        return None

    def proponer(self, clave: str, area: str, nombre: str) -> Proveedor:
        """Provider not registered yet: it will be added on commit (csv + folder)."""
        p = Proveedor(clave, area, nombre, limpiar_nombre(nombre))
        self.nuevos[clave] = p
        return p

    def cuit_por_carpeta(self) -> dict[str, str]:
        return {p.carpeta.lower(): c for c, p in self.por_clave.items() if p.carpeta}

    def nombre_a_cuit(self, hoja_proveedores=()) -> dict[str, str]:
        """norm(name) -> key, from the csv (name and folder) and the Proveedores sheet."""
        m = {}
        for c, p in list(self.por_clave.items()) + list(self.nuevos.items()):
            for n in (p.nombre, p.carpeta):
                if n:
                    m[norm(n)] = c
        for r in hoja_proveedores:
            if r.get("cuit") and r.get("nombre") and norm(r["nombre"]) not in m:
                m[norm(r["nombre"])] = r["cuit"]
        return m
