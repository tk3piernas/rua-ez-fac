"""Guided tour for beginners: shown on the first start and with the "?" button in the header.

The steps point at elements marked with data-tour="..." in ui.py; static/recorrido.js draws the highlight and
the card. A step whose element is not on screen (the grid before reading a batch, "N para leer" with nothing
waiting) shows its card in the middle with the "falta" note."""
from __future__ import annotations

import json
from datetime import datetime
from pathlib import Path

PASOS = [
    {"titulo": "Bienvenido a RUA EZ-Fac",
     "texto": "Te muestro cada parte de la app en un minuto. Podés salir cuando quieras con <b>Saltar</b> o Esc, "
              "y volver a verlo con el botón <b>?</b> de arriba."},
    {"sel": "modo", "titulo": "Modo de trabajo",
     "texto": "Siempre dice sobre qué trabajás: <b>solo simulando</b> (no se toca nada), la <b>carpeta de prueba</b> "
              "(una copia para practicar) o lo <b>real</b>. Se cambia en Configuración."},
    {"sel": "para-leer", "titulo": "Facturas esperando",
     "texto": "Cuando hay facturas en <i>_Entrada</i> o subidas acá, aparece <b>N para leer</b> (y cuántas son "
              "nuevas). Un clic las lee desde cualquier pestaña.",
     "falta": "Ahora no hay nada esperando, por eso no se ve."},
    {"sel": "planilla", "titulo": "La planilla",
     "texto": "<b>Abrir planilla</b> la abre en Excel; <b>Exportar planilla</b> guarda una copia con el nombre que "
              "elijas, sin cambiar la original."},
    {"sel": "salir", "titulo": "Salir",
     "texto": "Cierra la ventana y la app. Para abrirla de nuevo, el ícono RUA EZ-Fac (lo podés anclar a la barra de "
              "tareas con clic derecho)."},
    {"tab": "tab-cargar", "sel": "area", "titulo": "1. Cargar: el área",
     "texto": "Elegí de qué área son los archivos que subís (Vehicular, Hogar, Persona, ServiciosYa) o dejá "
              "<i>Detectar por proveedor</i>. Los .txt necesitan el área elegida."},
    {"tab": "tab-cargar", "sel": "subir", "titulo": "Subir facturas",
     "texto": "Arrastrá acá los PDF, zip o .txt de los proveedores (o hacé clic). Se copian a la app: los originales "
              "quedan donde estaban. También podés dejarlos en <i>Facturas pendientes\\_Entrada\\&lt;área&gt;</i>."},
    {"tab": "tab-cargar", "sel": "que-se-lee", "titulo": "Qué se va a leer",
     "texto": "Cuenta lo que hay en cada <i>_Entrada</i> y lo subido acá, con una X para quitar un archivo subido."},
    {"tab": "tab-cargar", "sel": "leer", "titulo": "Leer facturas",
     "texto": "Lee todo y arma el lote para revisar. <b>Leer no mueve ni escribe nada.</b>"},
    {"tab": "tab-revision", "sel": "avisos", "titulo": "2. Revisión: los avisos",
     "texto": "Lo que conviene mirar antes de cargar: <b>frenan</b> (no se cargan), <b>para revisar</b> y "
              "<b>para tener en cuenta</b>. <i>Lo que ves ahora es un lote de ejemplo: no se carga nada.</i>",
     "falta": "Aparecen después de leer un lote."},
    {"tab": "tab-revision", "sel": "grilla", "titulo": "Las filas que se cargan",
     "texto": "Doble clic para corregir proveedor, fechas, N° de tarea, servicio o monto. Destildá <i>Cargar</i> "
              "para dejar una fila afuera. <b>Amarillo</b>: posible duplicado; <b>rojo</b>: frenada; "
              "<b>violeta</b>: la leyó Claude.",
     "falta": "La tabla aparece después de leer un lote."},
    {"tab": "tab-revision", "sel": "confirmar", "titulo": "Confirmar y cargar",
     "texto": "Mueve cada PDF a la carpeta de su proveedor, da de alta los nuevos y escribe las filas en la planilla. "
              "Todo se puede deshacer después.",
     "falta": "El botón aparece con un lote leído."},
    {"tab": "tab-resultado", "sel": "resultado", "titulo": "3. Resultado",
     "texto": "Qué se hizo con el último lote: archivos, filas, errores. Desde acá se reintenta lo que quedó "
              "pendiente o se deshace el lote."},
    {"tab": "tab-historial", "sel": "tanda", "titulo": "Historial: la tanda de pagos",
     "texto": "<b>Cerrar la tanda</b> guarda una copia de la planilla como <i>tanda N</i>, anota sus lotes y, si "
              "querés, la deja vacía para la próxima."},
    {"tab": "tab-historial", "sel": "lotes", "titulo": "Lotes anteriores",
     "texto": "Todos los lotes, con fecha, tanda y dónde se cargaron (PRUEBA o REAL). El último se puede deshacer, "
              "o terminar si quedó a medias."},
    {"tab": "tab-revisar", "sel": "revisar", "titulo": "Revisar la planilla",
     "texto": "Compara la planilla con las carpetas: links rotos, filas repetidas o sin link, PDF que no están en "
              "ninguna fila. Solo lee."},
    {"tab": "tab-proveedores", "sel": "proveedores", "titulo": "Proveedores",
     "texto": "Pegá los datos de un proveedor como vengan (mail, WhatsApp, .txt): la app arma la ficha con CUIT, "
              "CBU y alias, los chequea y los guarda en la memoria de proveedores."},
    {"tab": "tab-config", "sel": "modo-trabajo", "titulo": "Configuración: el modo",
     "texto": "Acá elegís <b>Solo simular</b>, <b>Practicar en la carpeta de prueba</b> o <b>Cargar en la planilla "
              "real</b>. Para empezar, practicá en la carpeta de prueba."},
    {"tab": "tab-config", "sel": "setup", "titulo": "Setup",
     "texto": "Crea las carpetas que falten en tu OneDrive de RUA y revisa que esté todo bien. Nunca pisa nada."},
    {"tab": "tab-config", "sel": "reset", "titulo": "Empezar la planilla de cero",
     "texto": "Vacía las filas de la planilla (guarda respaldos antes, y se puede volver atrás)."},
    {"tab": "tab-config", "sel": "claude", "titulo": "Claude para lo que no se puede leer",
     "texto": "Con una clave de la API de Claude, las facturas que la app no pudo leer (escaneadas, por ejemplo) se "
              "leen con un clic desde Revisión."},
    {"titulo": "¡Listo!",
     "texto": "Eso es todo. Si querés ver esto de nuevo, tocá <b>?</b> arriba. Cualquier duda o error, avisale a "
              "Matheo."},
]


def _archivo(datos: Path) -> Path:
    return datos / "recorrido.json"


def visto(datos: Path) -> bool:
    try:
        return bool(json.loads(_archivo(datos).read_text(encoding="utf-8")).get("visto"))
    except (OSError, ValueError):
        return False


def marcar_visto(datos: Path, terminado: bool) -> None:
    datos.mkdir(parents=True, exist_ok=True)
    from .util import guardar_json
    guardar_json(_archivo(datos), {"visto": True, "terminado": terminado,
                                   "fecha": datetime.now().isoformat(timespec="seconds")})


def js() -> str:
    """The call that starts the tour in the page."""
    return f"window.ezRecorrido && window.ezRecorrido({json.dumps(PASOS, ensure_ascii=False)})"


# ------------------------------------------------------------------ example batch shown during the tour

def lote_demo(cfg):
    """An invented batch, only in memory, so the Revisión steps show a real-looking grid, warnings and
    the confirm button. It points at no file and cannot be confirmed (ui.py refuses while S.demo is on)."""
    from datetime import date, timedelta
    from decimal import Decimal as D

    from .model import Archivo, Aviso, Estado, Extraccion, Factura, Fila, Lote, Nivel
    hoy = date.today()
    base = Path(cfg.base) if cfg and cfg.base else Path("Facturas pendientes")
    lote = Lote(id="EJEMPLO", simular=True)
    ejemplos = [
        # id, provider, file, letter, amount, IVA (H), task, service, source, colour
        ("f01", "Grúas del Ejemplo SRL", "factura_gruas.pdf", "A", D("100000.00"), None, "412345",
         "Auxilio mecánico", "arca", ""),
        ("f02", "Remolques Modelo", "remolque_0007.pdf", "C", D("58000.00"), D("0"), "412346",
         "Remolque 40 km", "arca", "amarilla"),
        ("f03", "Cerrajería Inventada", "cerrajeria.pdf", "B", D("36300.00"), D("0"), "412347",
         "Apertura de puerta", "generico", "roja"),
        ("f04", "Plomería de Muestra", "escaneada.pdf", "C", D("42000.00"), D("0"), "412348",
         "Destapación", "claude", "violeta"),
    ]
    for i, (fid, prov, archivo, letra, monto, iva, tarea, servicio, fuente, color) in enumerate(ejemplos):
        area = "Hogar" if fid in ("f03", "f04") else "Vehicular"
        nro = 101 + i
        ext = Extraccion(fuente=fuente, letra=letra, pv=3, nro=nro, razon_social=prov, emision=hoy - timedelta(days=5),
                         vto=hoy + timedelta(days=25), total=monto + (iva or 0))
        f = Factura(id=fid, archivo=Archivo(base / "_Entrada" / area / archivo, archivo, area, "carpeta"), ext=ext,
                    estado=Estado.LISTA, proveedor=prov, carpeta=prov, texto_link=f"Factura 0003-{nro:08d}",
                    proveedor_nuevo=(fid == "f01"))
        lote.facturas.append(f)
        lote.filas.append(Fila(factura_id=fid, hoja=area, proveedor=prov, tarea=tarea, servicio=servicio,
                               fecha=ext.emision, vto=ext.vto, monto=monto, iva=iva, texto_link=f.texto_link,
                               resaltar=(color == "amarilla"), excluida=(color == "roja"),
                               bloqueada=(color == "roja"),
                               comentario="Posible duplicado:\n- la tarea 412346 ya figura" if color == "amarilla" else ""))
        if color != "roja":
            lote.plan.append({"op": "mover", "de": str(f.archivo.ruta), "a": str(base / area / prov / f"{f.texto_link}.pdf"),
                              "factura": fid})
    lote.avisos += [
        Aviso(Nivel.BLOQUEA, "mismo_numero", "cerrajeria.pdf: tiene el mismo número que otra factura pero otro importe "
              "(ejemplo): no se carga hasta revisarla", "f03"),
        Aviso(Nivel.REVISAR, "proveedor_nuevo", "factura_gruas.pdf: Grúas del Ejemplo SRL no está en proveedores.csv; "
              "se agregaría como proveedor nuevo (ejemplo)", "f01"),
        Aviso(Nivel.REVISAR, "posible_duplicado", "remolque_0007.pdf: la tarea 412346 ya figura en la planilla (ejemplo)",
              "f02"),
        Aviso(Nivel.REVISAR, "nota", "escaneada.pdf: la leyó Claude: revisá CUIT, número y montos (ejemplo)", "f04"),
        Aviso(Nivel.INFO, "factura_c", "remolque_0007.pdf: es Factura C, H va en 0 (ejemplo)", "f02"),
    ]
    lote.contexto.update(proveedores_csv=12, guardadas=148, planilla="40 filas cargadas (ejemplo)")
    return lote


def resultado_demo() -> dict:
    """What 'Resultado' shows after confirming, for the tour."""
    return {"hechos": [r"factura_gruas.pdf -> Vehicular\Grúas del Ejemplo SRL\Factura 0003-00000101.pdf (ejemplo)",
                       r"remolque_0007.pdf -> Vehicular\Remolques Modelo\Factura 0003-00000102.pdf (ejemplo)",
                       r"escaneada.pdf -> Hogar\Plomería de Muestra\Factura 0003-00000104.pdf (ejemplo)"],
            "errores": [], "excel": "escrito", "excel_mensaje": "", "excel_pendiente": 0,
            "excel_escrito": {"filas": 3, "proveedores": 1}, "memoria": 0, "sin_archivo": 0}
