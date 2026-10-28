"""Payment / contact details that providers send as .txt."""
from __future__ import annotations

import re
from dataclasses import dataclass, field
from pathlib import Path

from .bancos import banco_de, coincide
from .util import cbu_valido, cuit_valido, solo_digitos


@dataclass
class Cuenta:
    cbu: str
    alias: str = ""
    banco_escrito: str = ""
    tipo: str = ""
    nro: str = ""
    titular: str = ""

    @property
    def banco(self) -> str:
        return banco_de(self.cbu)

    @property
    def valido(self) -> bool:
        return cbu_valido(self.cbu)


@dataclass
class DatosTxt:
    archivo: Path
    nombre: str = ""          # "Proveedor:" / "Nombre de fantasía:" / "Razón social:"
    cuit: str = ""
    condicion_iva: str = ""
    domicilio: str = ""
    contacto: str = ""
    telefono: str = ""
    email: str = ""
    cuentas: list[Cuenta] = field(default_factory=list)
    titular: str = ""
    avisos: list[str] = field(default_factory=list)


def leer_texto(p: Path) -> str:
    b = p.read_bytes()
    for enc in ("utf-8-sig", "cp1252", "latin-1"):
        try:
            return b.decode(enc)
        except UnicodeDecodeError:
            continue
    return b.decode("utf-8", "replace")


def _campo(t: str, patron: str) -> str:
    m = re.search(rf"(?im)^\s*(?:{patron})\s*[:\-]\s*(.+?)\s*$", t)
    return m.group(1).strip() if m else ""


def _tipo_cuenta(s: str) -> str:
    s2 = s.lower()
    if "caja de ahorro" in s2:
        return "Caja de ahorro"
    if "corriente" in s2:
        return "Cuenta corriente"
    return s.strip()


def leer(p: Path) -> DatosTxt:
    t = leer_texto(p)
    d = DatosTxt(p)
    d.nombre = _campo(t, r"Proveedor|Nombre\s+de\s+fantas[ií]a|Raz[oó]n\s+social|Empresa")
    m = re.search(r"(?i)\b(?:CUIT|CUIL)\b\s*[:\-]?\s*(\d{2}[-\s.]?\d{8}[-\s.]?\d)", t)
    if m:
        d.cuit = solo_digitos(m.group(1))
    else:
        for mm in re.finditer(r"(?<!\d)(\d{2})[-\s]?(\d{8})[-\s]?(\d)(?!\d)", t):
            c = "".join(mm.groups())
            if cuit_valido(c):
                d.cuit = c
                break
    if d.cuit and not cuit_valido(d.cuit):
        d.avisos.append(f"El CUIT {d.cuit} tiene el dígito verificador mal")
    tl = t.lower()
    d.condicion_iva = ("IVA Responsable Inscripto" if "responsable inscripto" in tl else
                       "Responsable Monotributo" if "monotribut" in tl else "IVA Exento" if "exento" in tl else "")
    d.domicilio = _campo(t, r"Domicilio(?:\s+comercial)?|Direcci[oó]n")
    d.contacto = _campo(t, r"Contacto|Nombre(?:\s+y\s+apellido)?|Responsable")
    d.telefono = re.sub(r"[^\d+]", "", _campo(t, r"Tel[eé]fono|Tel\.?|Cel(?:ular)?|Whats\s*App|M[oó]vil"))
    m = re.search(r"[\w.+-]+@[\w-]+(\.[\w-]+)+", t)
    d.email = m.group(0) if m else ""
    d.titular = _campo(t, r"Titular(?:\s+de\s+la\s+cuenta)?")

    # accounts: one per CBU/CVU; alias, bank and type are taken from the same block (blank-line separated)
    bloques = re.split(r"\n\s*\n", t.replace("\r\n", "\n"))
    for b in bloques:
        for mm in re.finditer(r"(?<!\d)((?:\d[\s-]?){21}\d)(?!\d)", b):
            cbu = solo_digitos(mm.group(1))
            c = Cuenta(cbu)
            c.alias = (re.search(r"(?i)alias\s*[:\-]?\s*([A-Za-z0-9.\-]{6,20})", b) or [None, ""])[1]
            c.banco_escrito = _campo(b, r"Banco|Entidad")
            if not c.banco_escrito:
                m2 = re.search(r"(?i)(mercado\s*pago|ual[aá]|brubank|naranja\s*x)", b)
                c.banco_escrito = m2.group(1) if m2 else ""
            c.tipo = _tipo_cuenta(_campo(b, r"Tipo\s+de\s+cuenta|Cuenta"))
            if cbu.startswith("000"):
                c.tipo = "CVU"
            c.nro = _campo(b, r"N[°ºro.]*\s*de\s*cuenta|Cuenta\s*N[°ºro.]*")
            c.titular = _campo(b, r"Titular(?:\s+de\s+la\s+cuenta)?") or d.titular
            d.cuentas.append(c)
            if not c.valido:
                d.avisos.append(f"CBU/CVU {cbu} con dígito verificador inválido: revisar")
            elif not coincide(cbu, c.banco_escrito):
                d.avisos.append(f"El .txt dice '{c.banco_escrito}' pero el CBU {cbu[:3]}… es de {c.banco}")
    if not d.cuentas:
        d.avisos.append("Sin CBU/CVU")
    return d
