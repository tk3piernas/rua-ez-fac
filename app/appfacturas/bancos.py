"""Bank from the first 3 digits of a CBU (table from the PS1, plus updates)."""
from __future__ import annotations

from .util import norm, solo_digitos

BANCOS = {
    "007": "Banco Galicia", "011": "Banco Nación", "014": "Banco Provincia", "015": "ICBC",
    "016": "Citibank", "017": "BBVA", "020": "Bancor (Banco de Córdoba)", "027": "Banco Supervielle",
    "029": "Banco Ciudad", "034": "Banco Patagonia", "044": "Banco Hipotecario", "045": "Banco San Juan",
    "060": "Banco de Tucumán", "065": "Banco Municipal de Rosario", "072": "Banco Santander",
    "083": "Banco del Chubut", "086": "Banco de Santa Cruz", "093": "Banco de La Pampa",
    "094": "Banco de Corrientes", "097": "Banco Provincia del Neuquén", "143": "Brubank",
    "147": "Banco Interfinanzas", "150": "HSBC", "158": "Openbank", "165": "JP Morgan",
    "191": "Banco Credicoop", "198": "Banco de Valores", "247": "Banco Roela", "254": "Banco Mariva",
    "259": "Banco Itaú", "266": "BNP Paribas", "268": "Banco Tierra del Fuego",
    "269": "Banco República Oriental del Uruguay", "277": "Banco Sáenz", "281": "Banco Meridian",
    "285": "Banco Macro", "299": "Banco Comafi", "300": "BICE", "301": "Banco Piano", "305": "Banco Julio",
    "309": "Banco Rioja", "310": "Banco del Sol", "311": "Nuevo Banco del Chaco", "312": "Banco Voii",
    "315": "Banco de Formosa", "319": "Banco CMF", "321": "Banco de Santiago del Estero",
    "322": "Banco Industrial (BIND)", "330": "Nuevo Banco de Santa Fe", "331": "Banco Cetelem",
    "332": "Banco de Servicios Financieros", "338": "Banco de Servicios y Transacciones",
    "339": "RCI Banque", "340": "BACS", "341": "Banco Masventas",
    "384": "Ualá Bank (ex Wilobank)",
    "386": "Nuevo Banco de Entre Ríos", "389": "Banco Columbia", "426": "Banco Bica",
    "431": "Banco Coinag", "432": "Banco de Comercio", "435": "Banco Sucrédito", "448": "Banco Dino",
    "515": "Bank of China",
}

# words that identify each bank when the provider writes it by hand
_ALIAS = {"014": ["provincia", "bapro"], "011": ["nacion"], "007": ["galicia"], "072": ["santander"],
          "017": ["bbva", "frances"], "285": ["macro"], "029": ["ciudad"], "191": ["credicoop"],
          "384": ["uala", "wilobank"], "143": ["brubank"], "150": ["hsbc"], "259": ["itau"],
          "027": ["supervielle"], "015": ["icbc"], "034": ["patagonia"], "044": ["hipotecario"]}


def banco_de(cbu: str) -> str:
    c = solo_digitos(cbu)
    if len(c) != 22:
        return ""
    if c.startswith("000"):
        if c.startswith("0000003"):
            return "Mercado Pago"
        if c.startswith("0000007"):
            return "Ualá"
        return "Billetera virtual (revisar)"
    return BANCOS.get(c[:3], f"Banco cód. {c[:3]} (revisar)")


def coincide(cbu: str, banco_escrito: str) -> bool:
    """Does the bank the provider wrote match the CBU code? (True when it can't tell)."""
    c = solo_digitos(cbu)
    if len(c) != 22 or not banco_escrito:
        return True
    escrito = norm(banco_escrito)
    if c.startswith("000"):
        return any(w in escrito for w in ("mercado", "uala", "billetera", "cvu", "naranja", "personal", "brubank"))
    claves = _ALIAS.get(c[:3]) or [w for w in norm(BANCOS.get(c[:3], "")).split() if w not in ("banco", "de", "del", "la", "nuevo")]
    return any(k in escrito for k in claves) if claves else True
