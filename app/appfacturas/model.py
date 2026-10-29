"""Data model of a batch (lote). Everything serialises to JSON for lote.json."""
from __future__ import annotations

from dataclasses import asdict, dataclass, field, is_dataclass
from datetime import date
from decimal import Decimal
from enum import Enum
from pathlib import Path


class Nivel(str, Enum):
    BLOQUEA = "bloquea"      # the invoice/row is not loaded
    REVISAR = "revisar"      # loaded, but look at it (yellow, Factura A, due soon...)
    INFO = "info"


class Estado(str, Enum):
    LEIDA = "leída"
    REVISAR = "revisar"
    LISTA = "lista"
    DUPLICADA = "duplicada"
    EXCLUIDA = "excluida"
    PENDIENTE_CLAUDE = "pendiente Claude"
    NO_LEIDA = "no leída"
    CONFIRMADA = "confirmada"
    HECHA = "hecha"


@dataclass
class Aviso:
    nivel: Nivel
    codigo: str
    texto: str
    ref: str = ""            # factura id or "" for the whole batch


@dataclass
class Archivo:
    ruta: Path                       # where it is now (temp copy for zips)
    nombre: str                      # name shown to the user ("x.zip > a.pdf")
    area: str = ""
    area_origen: str = ""            # carpeta / proveedor / elegida
    zip: Path | None = None          # zip it came from
    sha256: str = ""
    tamano: int = 0


@dataclass
class Item:
    descripcion: str
    monto: Decimal | None            # net for A/M, total for B/C
    tareas: list[str] = field(default_factory=list)
    codigo: str = ""
    iva: Decimal | None = None       # only when the invoice details it per item (A/M)


@dataclass
class Extraccion:
    """What a parser (or Claude) returns. Same fields for every source."""
    fuente: str = ""
    cuit: str = ""
    cod_comprobante: int | None = None
    letra: str = ""
    tipo: str = "Factura"            # Factura / Nota de Crédito / Nota de Débito / Recibo / Nota de Venta
    pv: int | None = None
    nro: int | None = None
    razon_social: str = ""
    nombre_fantasia: str = ""
    domicilio: str = ""
    telefono: str = ""               # the issuer's, only from its part of the invoice (parsers.base.contacto_emisor)
    email: str = ""
    condicion_emisor: str = ""
    condicion_receptor: str = ""
    cuit_receptor: str = ""
    emision: date | None = None
    desde: date | None = None
    vto: date | None = None
    neto: Decimal | None = None
    iva: Decimal | None = None
    total: Decimal | None = None
    otros_tributos: Decimal | None = None
    no_gravado: Decimal | None = None    # "Importe Neto No Gravado" + "Importe Exento": paid, without IVA
    items: list[Item] = field(default_factory=list)
    fiscal: bool = True              # False = no CUIT/CAE (e.g. "FAC0001" receipts)
    numero_txt: str = ""             # number of a non-fiscal receipt ("FAC0001")
    cbu: str = ""                    # CBU printed on the invoice, if any
    tareas_archivo: list[str] = field(default_factory=list)   # task numbers in the file name (hint only)
    confianza: float = 0.0
    faltantes: list[str] = field(default_factory=list)
    notas: list[str] = field(default_factory=list)

    @property
    def es_nc(self) -> bool:
        return self.tipo == "Nota de Crédito"

    @property
    def numero(self) -> str:
        if self.pv is None or self.nro is None:
            return self.numero_txt or "?"
        return f"{self.pv:05d}-{self.nro:08d}" if self.pv > 9999 else f"{self.pv:04d}-{self.nro:08d}"


@dataclass
class Factura:
    id: str
    archivo: Archivo
    ext: Extraccion | None = None
    estado: Estado = Estado.LEIDA
    proveedor: str = ""               # name as in the registry / sheet
    proveedor_clave: str = ""         # CUIT or "SC:<name>"
    carpeta: str = ""                 # provider folder
    proveedor_nuevo: bool = False
    clave: str = ""                   # CUIT|tipo|PV|nro
    nombre_destino: str = ""          # "Factura 0002-00000189.pdf"
    destino: Path | None = None       # where the PDF would go (None = already saved / not moved)
    a_duplicados: bool = False        # the incoming copy goes to _Entrada\_Duplicados
    existente: Path | None = None     # already saved copy (link points there)
    url: str = ""
    texto_link: str = ""              # "Factura 0002-00000189"
    monto: Decimal | None = None      # G for the whole invoice (negative for NC)
    iva: Decimal | None = None        # H for the whole invoice (None = leave formula)
    dup: str = ""                     # certain duplicate: why it is not loaded
    vto_supuesto: bool = False        # no due date in the PDF: F got the issue date
    posibles: list[str] = field(default_factory=list)   # possible duplicates (yellow)


@dataclass
class Fila:
    factura_id: str
    hoja: str
    proveedor: str
    tarea: str = ""
    servicio: str = ""
    empresa: str = ""
    fecha: date | None = None
    vto: date | None = None
    monto: Decimal | None = None
    iva: Decimal | None = None        # None = keep the sheet formula
    url: str = ""
    texto_link: str = ""
    resaltar: bool = False
    comentario: str = ""
    excluida: bool = False
    editada: bool = False
    bloqueada: bool = False           # its invoice has a BLOQUEA warning: never loaded, cannot be ticked


@dataclass
class Lote:
    id: str
    simular: bool = True
    modo_filas: str = "item"
    facturas: list[Factura] = field(default_factory=list)
    filas: list[Fila] = field(default_factory=list)
    avisos: list[Aviso] = field(default_factory=list)
    pendientes: list[str] = field(default_factory=list)
    proveedores: list[dict] = field(default_factory=list)   # rows A:O for the Proveedores sheet (+ meta)
    txts: list[dict] = field(default_factory=list)          # what each .txt said and where it was assigned
    plan: list[dict] = field(default_factory=list)
    diario: list[dict] = field(default_factory=list)
    contexto: dict = field(default_factory=dict)

    def aviso(self, nivel: Nivel, codigo: str, texto: str, ref: str = "") -> None:
        self.avisos.append(Aviso(nivel, codigo, texto, ref))

    @property
    def total(self) -> Decimal:
        return sum((f.monto or Decimal(0) for f in self.filas if not f.excluida), Decimal(0))


def a_json(o):
    """json.dump default= for the model."""
    if is_dataclass(o):
        return asdict(o)
    if isinstance(o, (Decimal,)):
        return str(o)
    if isinstance(o, (date,)):
        return o.isoformat()
    if isinstance(o, Path):
        return str(o)
    if isinstance(o, Enum):
        return o.value
    raise TypeError(type(o))


def cargar_lote(ruta: Path) -> "Lote":
    """Rebuild the parts of a saved batch needed to retry it: rows, provider rows and the file plan."""
    import json
    from datetime import date as _d
    d = json.loads(Path(ruta).read_text(encoding="utf-8"))
    dec = lambda v: Decimal(v) if v not in (None, "") else None
    fec = lambda v: _d.fromisoformat(v) if v else None
    lote = Lote(id=d["id"], simular=d.get("simular", False), modo_filas=d.get("modo_filas", "item"))
    for r in d.get("filas", []):
        lote.filas.append(Fila(**{**r, "fecha": fec(r.get("fecha")), "vto": fec(r.get("vto")),
                                  "monto": dec(r.get("monto")), "iva": dec(r.get("iva"))}))
    lote.proveedores = d.get("proveedores", [])
    lote.plan = d.get("plan", [])
    return lote
