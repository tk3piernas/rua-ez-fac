r"""Parity test against the old PS1 (retired): compare both simulations invoice by invoice.
Kept as a tool, outside the app package.

  .venv\Scripts\python.exe app\herramientas\paridad.py preparar    -> rebuilds the test folder
  .venv\Scripts\python.exe app\herramientas\paridad.py comparar    -> datos\paridad_resultado.txt
"""
from __future__ import annotations

import json
import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from appfacturas.carpeta_prueba import carpeta_prueba, preparar  # noqa: E402,F401
from appfacturas.config import carpeta_app  # noqa: E402


# ------------------------------------------------------------------ PS1 transcript

def leer_ps1(ruta: Path) -> dict:
    t = ruta.read_text(encoding="utf-8-sig", errors="replace").replace("\r", "")
    r = {"pdf": {}, "dup": {}, "pend": {}, "filas": [], "no_carga": [], "txt": {}, "prov": []}
    seccion = ""
    ultima = None
    for linea in t.split("\n"):
        l = linea.rstrip()
        if m := re.match(r"^PDF  (\w+)\\(.+?)  ->  (.+)$", l):
            r["pdf"][m.group(2)] = m.group(3).replace("\\", "/")
        elif m := re.match(r"^DUP  (\w+)\\(.+?) (es el mismo archivo que|es la misma factura que|ya estaba guardada como) (.+?);", l):
            r["dup"][m.group(2)] = (m.group(3), m.group(4).replace("\\", "/"))
        elif m := re.match(r"^TXT  (\w+)\\(.+?)  ->  (.+?)  \(CUIT (\d*), (.*)\)$", l):
            r["txt"][m.group(2)] = {"proveedor": m.group(3), "cuit": m.group(4), "banco": m.group(5)}
        elif l.startswith("Quedaron en _Entrada"):
            seccion = "pend"
        elif l.startswith("FILAS QUE SE AGREGAR"):
            seccion = "filas"
        elif l.startswith("FILAS PARA LA HOJA PROVEEDORES"):
            seccion = "prov"
        elif seccion == "pend" and (m := re.match(r"^  - (\w+)\\(.+?)  -> (.+)$", l)):
            r["pend"][m.group(2)] = m.group(3)
        elif seccion == "filas" and (m := re.match(r"^  NO SE CARGA: (.+?) \| (.+?) (ya está|su link)(.*)$", l)):
            r["no_carga"].append({"proveedor": m.group(1), "texto": m.group(2), "motivo": m.group(3) + m.group(4)})
        elif seccion == "filas" and (m := re.match(r"^  (\w+): (.+?) \| (.+?) \| (.+?) \| desde (.*?) \| vence (.*?) \| (.+?) \| (.+)$", l)):
            ultima = {"hoja": m.group(1), "proveedor": m.group(2), "texto": m.group(3),
                      "tareas": "" if m.group(4) == "sin tarea" else m.group(4).replace("tarea ", ""),
                      "desde": m.group(5), "vto": m.group(6), "monto": m.group(7),
                      "iva": "fórmula" if "fórmula" in m.group(8) else m.group(8).replace("IVA ", ""), "posible": False}
            r["filas"].append(ultima)
        elif seccion == "filas" and "posible duplicado" in l and ultima:
            ultima["posible"] = True
        elif seccion == "prov" and l.startswith("  ") and " | " in l:
            r["prov"].append([c.strip() for c in l.strip().split(" | ")])
        elif l and not l.startswith(" "):
            if seccion in ("filas", "prov") and not l.startswith(("  ", "     ")):
                seccion = "" if not l.startswith("FILAS") else seccion
    return r


# ------------------------------------------------------------------ comparison

def _monto(v) -> str:
    from decimal import Decimal
    from appfacturas.util import fmt_monto
    return fmt_monto(Decimal(v)) if v not in (None, "") else "sin monto"


def comparar(ps1: Path, lote_dir: Path, salida: Path) -> int:
    p = leer_ps1(ps1)
    L = json.loads((lote_dir / "lote.json").read_text(encoding="utf-8"))
    facturas = {f["archivo"]["nombre"]: f for f in L["facturas"]}
    filas_app: dict[str, list] = {}
    for r in L["filas"]:
        filas_app.setdefault(r["factura_id"], []).append(r)
    out, sin_explicar, explicadas, iguales = [], 0, 0, 0

    def linea(archivo, campo, a, b, nota=""):
        nonlocal sin_explicar, explicadas, iguales
        if a == b:
            iguales += 1
            return
        if nota:
            explicadas += 1
            out.append(f"  [explicada] {archivo} | {campo}: PS1 = {a!r} / app = {b!r} -> {nota}")
        else:
            sin_explicar += 1
            out.append(f"  [SIN EXPLICAR] {archivo} | {campo}: PS1 = {a!r} / app = {b!r}")

    nombres = sorted(set(facturas) | set(p["pdf"]) | set(p["dup"]) | set(p["pend"]))
    for n in nombres:
        f = facturas.get(n)
        e = (f or {}).get("ext") or {}
        # where it ends up
        if n in p["pdf"]:
            st = "guarda en " + p["pdf"][n]
        elif n in p["dup"]:
            st = "duplicado"
        elif n in p["pend"]:
            st = "pendiente"
        else:
            st = "no aparece"
        if not f:
            app = "no aparece"
        elif f["estado"] == "no leída":
            app = "pendiente"
        elif f["a_duplicados"] or (f["dup"] and f["estado"] == "duplicada"):
            app = "duplicado"
        elif f["estado"] == "revisar" and f["dup"]:
            app = "bloqueada (mismo número, otro importe)"
        else:
            app = "guarda en " + _rel(f["destino"])
        nota = ""
        if st == "pendiente" and app.startswith("guarda") and "no está en proveedores.csv" in p["pend"].get(n, ""):
            nota = "proveedor nuevo: la app lo da de alta (el PS1 pedía agregarlo a mano al csv)"
        elif st == "pendiente" and app != "pendiente" and "no pude leer" in p["pend"].get(n, ""):
            nota = "la app lee este formato (lector genérico / no fiscal)"
        elif st.startswith("guarda") and app.startswith("guarda") and st.split("/")[-1] == app.split("/")[-1]:
            nota = "misma factura y mismo archivo; cambia el nombre de la carpeta del proveedor nuevo"
        linea(n, "destino", st, app, nota)
        if not f or not e:
            continue
        # rows: compare against the PS1 row with the same text and sheet
        texto = f.get("texto_link")
        rp = next((r for r in p["filas"] if r["texto"] == texto and r["hoja"] == f["archivo"]["area"]), None)
        nc = next((r for r in p["no_carga"] if r["texto"] == texto), None)
        ra = filas_app.get(f["id"], [])
        dup_tanda = n in p["dup"] and p["dup"][n][0] == "es la misma factura que"
        if nc or dup_tanda or (not ra and f["dup"]):
            ps1_carga = "no se carga" if (nc or dup_tanda or not rp) else "se carga"
            linea(n, "carga", ps1_carga, "no se carga" if not ra else "se carga")
            continue
        if not rp and not ra:
            continue
        if not rp:
            linea(n, "fila", "-", "se carga", "el PS1 no cargaba esta factura (ver destino)" if st != app else "")
            continue
        if not ra:
            linea(n, "fila", "se carga", "-")
            continue
        r0 = ra[0]
        tareas_app = " / ".join(dict.fromkeys(t for r in ra for t in (r["tarea"].split(" / ") if r["tarea"] else [])))
        linea(n, "tarea", rp["tareas"], tareas_app,
              "la app lee tareas separadas por coma o barra" if rp["tareas"] and set(rp["tareas"].split(" / ")) < set(tareas_app.split(" / ")) else "")
        desde_app = _fecha(r0["fecha"])
        linea(n, "fecha (E)", rp["desde"], desde_app,
              "el PS1 deja E vacía cuando la factura no dice 'Desde'; la app usa la fecha de emisión" if not rp["desde"] else "")
        linea(n, "vencimiento (F)", rp["vto"], _fecha(r0["vto"]))
        g_ps1 = _monto_txt(rp["monto"])
        g_app = sum(_d(r["monto"]) for r in ra)
        neto, total = _d(e.get("neto")), _d(e.get("total"))
        linea(n, "monto (G)", _monto(g_ps1), _monto(g_app),
              "Factura A sin 'Neto Gravado:' reconocible: el PS1 cargaba el total con IVA en G y dejaba la fórmula en H "
              "(IVA contado dos veces); la app carga el neto" if g_ps1 is not None and abs(g_ps1) == total and abs(g_app) == neto else "")
        # what the PS1 really writes in H: the printed IVA only if it differs from the formula by more than $1
        h_ps1 = rp["iva"]
        if h_ps1 != "fórmula":
            iva = _monto_txt(h_ps1)
            if iva is not None and iva != 0 and g_ps1 is not None and abs(iva - (g_ps1 * _d("0.21")).quantize(_d("0.01"))) <= 1:
                h_ps1 = "fórmula"
            else:
                h_ps1 = _monto(iva)
        iva_app = "fórmula" if r0["iva"] is None else _monto(r0["iva"])
        linea(n, "IVA (H)", h_ps1, iva_app,
              "consecuencia de la diferencia en G (Factura A)" if h_ps1 == "fórmula" and g_ps1 == total and total != neto else "")
        linea(n, "posible duplicado", rp["posible"], any(r["resaltar"] for r in ra))
        if rp["proveedor"] != r0["proveedor"]:
            linea(n, "proveedor (A)", rp["proveedor"], r0["proveedor"],
                  "proveedor nuevo: el PS1 usa el nombre del archivo .txt o de la línea del csv; la app el de la factura o del .txt"
                  if f.get("proveedor_nuevo") else "")
    # .txt
    for t in L.get("txts", []):
        tp = p["txt"].get(t["nombre"])
        if not tp:
            linea(t["nombre"], "txt", "no aparece", t["proveedor"])
            continue
        linea(t["nombre"], "txt proveedor", tp["proveedor"], t["proveedor"],
              "el .txt no trae 'Proveedor:': el PS1 usa el nombre del archivo, la app el de la factura asociada"
              if tp["proveedor"].lower().startswith("datos para") else "")
        linea(t["nombre"], "txt CUIT", tp["cuit"], t["cuit"],
              "el .txt no trae CUIT: la app lo toma de la factura del mismo titular" if not tp["cuit"] else "")
        linea(t["nombre"], "txt banco", tp["banco"], t["banco"] or "",
              "tabla de bancos actualizada" if "384" in (t.get("cbu") or "")[:3] else "")

    res = [f"Paridad PS1 vs App Facturas  ({ps1.name} / {lote_dir.name})",
           f"Iguales: {iguales}   Diferencias explicadas: {explicadas}   SIN EXPLICAR: {sin_explicar}", ""] + out
    salida.parent.mkdir(parents=True, exist_ok=True)
    salida.write_text("\n".join(res), encoding="utf-8")
    print("\n".join(res))
    return sin_explicar


def _monto_txt(s):
    """'$ 260,000.00' or '$ -260.000,00' (the PS1 prints with the PC's culture) -> Decimal."""
    from appfacturas.util import parse_monto
    if not s or "sin monto" in s:
        return None
    return parse_monto(s.replace("$", "").replace(" ", ""))


def _rel(p) -> str:
    if not p:
        return "?"
    parts = re.split(r"[\\/]", str(p))
    return "/".join(parts[-3:])


def _fecha(iso) -> str:
    if not iso:
        return ""
    y, m, d = iso.split("-")
    return f"{d}/{m}/{y}"


def _d(v):
    from decimal import Decimal
    return Decimal(v) if v not in (None, "") else Decimal(0)


def ultimo_lote() -> Path:
    d = carpeta_app() / "datos" / "lotes"
    return max((x for x in d.iterdir() if x.is_dir()), key=lambda x: x.name)


if __name__ == "__main__":
    import argparse
    ap = argparse.ArgumentParser(description="paridad contra el script viejo (OrdenarFacturas.ps1)")
    ap.add_argument("accion", choices=["preparar", "comparar"])
    ap.add_argument("--ps1", type=Path, help="transcript del PS1")
    ap.add_argument("--lote", type=Path, help="carpeta del lote de la app (por defecto, el último)")
    a = ap.parse_args()
    if a.accion == "preparar":
        preparar()
        sys.exit(0)
    datos = carpeta_app() / "datos"
    sys.exit(1 if comparar(a.ps1 or datos / "ps1_paridad.txt", a.lote or ultimo_lote(), datos / "paridad_resultado.txt") else 0)
