"""config.toml: created with today's values the first time, never inside OneDrive."""
from __future__ import annotations

import os
import tomllib
from dataclasses import dataclass, field
from pathlib import Path

# invoice areas: one input folder, one provider folder and one workbook sheet each
AREAS = ("Vehicular", "Hogar", "Persona", "ServiciosYa")

PLANTILLA = '''# Configuración de RUA EZ-Fac. Se puede editar con el Bloc de notas.
# Si una ruta queda vacía, la app la busca sola (igual que el script viejo).

[rutas]
# Carpeta "Facturas pendientes" dentro del OneDrive de RUA sincronizado
facturas_pendientes = ""
# Planilla. Vacía = "DEUDA A PAGAR ARG.xlsx" en la raíz del OneDrive de RUA
planilla = ""
# Base de los links de SharePoint (columna J). Vacía = se arma sola con tu cuenta de OneDrive de RUA
url_base = ""

[reglas]
cuit_rua = "30717419428"
# N° de tarea: 6 cifras dentro de este rango, solo en el detalle de la factura
tarea_min = 300000
tarea_max = 999999
# "item" = una fila por ítem de la factura; "factura" = una fila por factura (como el script viejo)
filas = "item"
# Días para el aviso "vence pronto"
dias_vence_pronto = 7
# Diferencia máxima (en $) entre el IVA real y la fórmula de la planilla antes de pisar la columna H
tolerancia_iva = 1.0

[planilla]
# true = escribe en C el texto del servicio que trae la factura; false = C queda vacía para completarla a mano
escribir_servicio = false
# Valores por defecto para C (Servicio) y D (Empresa). Vacío = se completan a mano.
servicio_vehicular = ""
empresa_vehicular = ""
servicio_hogar = ""
empresa_hogar = ""

[app]
puerto = 8765
# false = la app solo simula sobre las carpetas y la planilla reales (las escrituras van a la carpeta de prueba).
# Lo cambia Matheo desde Configuración cuando aprueba usar la planilla real.
modo_real = false
'''


@dataclass
class Config:
    base: Path | None = None              # Facturas pendientes
    planilla: Path | None = None
    url_base: str = ""
    cuit_rua: str = "30717419428"
    tarea_min: int = 300000
    tarea_max: int = 999999
    filas: str = "item"
    dias_vence_pronto: int = 7
    tolerancia_iva: float = 1.0
    defaults: dict = field(default_factory=dict)
    escribir_servicio: bool = False
    puerto: int = 8765
    modo_real: bool = False
    excel_motor: str = "com"          # "openpyxl" only in tests (there is no Excel in the cloud)
    archivo: Path | None = None
    areas: tuple = AREAS

    @property
    def entrada(self) -> Path:
        return self.base / "_Entrada"

    @property
    def csv_proveedores(self) -> Path:
        return self.base / "_Script" / "proveedores.csv"


def carpeta_app() -> Path:
    """The app folder in Documents (or APPFACTURAS_HOME for tests)."""
    if os.environ.get("APPFACTURAS_HOME"):
        return Path(os.environ["APPFACTURAS_HOME"])
    return Path(__file__).resolve().parents[2]


def buscar_base() -> Path | None:
    """Same search as the PS1: OneDriveCommercial, then %USERPROFILE%\\OneDrive - *."""
    cands = []
    if os.environ.get("OneDriveCommercial"):
        cands.append(Path(os.environ["OneDriveCommercial"]) / "Facturas pendientes")
    home = Path(os.environ.get("USERPROFILE") or Path.home())
    if home.exists():
        cands += sorted(p / "Facturas pendientes" for p in home.glob("OneDrive - *") if p.is_dir())
    return next((c for c in cands if c.is_dir()), None)


TENANT = "https://ruaasistencia-my.sharepoint.com"


def url_desde_onedrive(base: Path | None) -> str:
    """SharePoint URL of 'Facturas pendientes' for the signed-in OneDrive for Business account (Windows registry)."""
    if base is None or os.name != "nt":
        return ""
    try:
        import winreg
    except ImportError:
        return ""
    raiz = r"Software\Microsoft\OneDrive\Accounts"
    cuentas = []
    for i in range(1, 10):
        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, rf"{raiz}\Business{i}") as k:
                v = {}
                for n in ("UserFolder", "UserEmail", "ServiceEndpointUri"):
                    try:
                        v[n] = str(winreg.QueryValueEx(k, n)[0])
                    except OSError:
                        v[n] = ""
                cuentas.append(v)
        except OSError:
            continue
    padre = str(base.parent).rstrip("\\").lower()
    cuentas.sort(key=lambda v: v["UserFolder"].rstrip("\\").lower() != padre)   # the account that syncs this folder first
    for v in cuentas:
        ep = v["ServiceEndpointUri"]
        if "/personal/" in ep:
            return ep.split("/_api")[0].rstrip("/") + "/Documents/" + base.name
        if v["UserEmail"]:
            usuario = v["UserEmail"].lower().replace("@", "_").replace(".", "_")
            return f"{TENANT}/personal/{usuario}/Documents/{base.name}"
    return ""


def cargar(ruta: Path | None = None, base: Path | None = None) -> Config:
    ruta = ruta or carpeta_app() / "config.toml"
    if not ruta.exists():
        ruta.parent.mkdir(parents=True, exist_ok=True)
        ruta.write_text(PLANTILLA, encoding="utf-8")
    d = tomllib.loads(ruta.read_text(encoding="utf-8"))
    r, g, pl, ap = d.get("rutas", {}), d.get("reglas", {}), d.get("planilla", {}), d.get("app", {})
    c = Config(archivo=ruta)
    c.base = base or (Path(r["facturas_pendientes"]) if r.get("facturas_pendientes") else buscar_base())
    if r.get("planilla"):
        c.planilla = Path(r["planilla"])
    elif c.base:
        c.planilla = c.base.parent / "DEUDA A PAGAR ARG.xlsx"
    c.url_base = r.get("url_base", "") or url_desde_onedrive(c.base)
    c.cuit_rua = str(g.get("cuit_rua", c.cuit_rua))
    c.tarea_min = int(g.get("tarea_min", c.tarea_min))
    c.tarea_max = int(g.get("tarea_max", c.tarea_max))
    c.filas = g.get("filas", c.filas) if g.get("filas") in ("item", "factura") else "item"
    c.dias_vence_pronto = int(g.get("dias_vence_pronto", c.dias_vence_pronto))
    c.tolerancia_iva = float(g.get("tolerancia_iva", c.tolerancia_iva))
    c.defaults = {k: v for k, v in pl.items() if v and isinstance(v, str)}
    c.escribir_servicio = bool(pl.get("escribir_servicio", False))
    c.puerto = int(ap.get("puerto", c.puerto))
    c.modo_real = bool(ap.get("modo_real", False))
    return c


def guardar_opcion(ruta: Path, seccion: str, clave: str, valor) -> None:
    """Changes one 'key = value' line of config.toml keeping the comments (adds it if missing)."""
    import re
    texto = ruta.read_text(encoding="utf-8")
    v = ("true" if valor else "false") if isinstance(valor, bool) else (str(valor) if isinstance(valor, (int, float)) else f'"{valor}"')
    m = re.search(rf"(?ms)^\[{seccion}\]\s*$(.*?)(?=^\[|\Z)", texto)
    if not m:
        texto = texto.rstrip() + f"\n\n[{seccion}]\n{clave} = {v}\n"
    else:
        cuerpo = m.group(1)
        if re.search(rf"(?m)^{clave}\s*=", cuerpo):
            cuerpo = re.sub(rf"(?m)^{clave}\s*=.*$", f"{clave} = {v}", cuerpo)
        else:
            cuerpo = cuerpo.rstrip("\n") + f"\n{clave} = {v}\n\n"
        texto = texto[:m.start(1)] + cuerpo + texto[m.end(1):]
    ruta.write_text(texto, encoding="utf-8")
