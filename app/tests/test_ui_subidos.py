"""'Quitar lo subido' deletes only what was uploaded through the app, and says what it did."""
from appfacturas import ui


def test_quitar_subidos(tmp_path, monkeypatch):
    monkeypatch.setenv("APPFACTURAS_HOME", str(tmp_path))
    sub = ui.carpeta_subidos()
    (sub / "_Detectar").mkdir(parents=True)
    (sub / "Hogar").mkdir()
    (sub / "_Detectar" / "a.pdf").write_bytes(b"x")
    (sub / "Hogar" / "b.txt").write_text("CBU", encoding="utf-8")
    assert ui.contar_entrada(ui.config.Config())["subidos"] == 2
    assert ui.quitar_subidos() == (2, 0)
    assert not [f for f in sub.rglob("*") if f.is_file()]
    assert ui.quitar_subidos() == (0, 0)


def test_columnas_de_la_grilla_no_se_achican():
    """The review grid keeps every fixed width as its minimum (squeezed headers were unreadable)."""
    for c in ui.COLUMNAS:
        assert c["headerTooltip"] == c["headerName"]
        if "width" in c:
            assert c["minWidth"] == c["width"]
        else:
            assert c["minWidth"] >= 190 and c.get("flex")


def test_rutas_no_se_cortan_a_mitad_de_palabra():
    """Paths wrapped with word-break: break-all showed 'DEU / DA A PAGAR ARG.xlsx'."""
    from pathlib import Path
    fuente = Path(ui.__file__).read_text(encoding="utf-8")
    assert "break-all" not in fuente and ".ruta { overflow-wrap: anywhere; word-break: normal; }" in ui.CSS


def test_en_solo_simular_no_se_toca_lo_real(tmp_path, monkeypatch):
    """A real batch can't be undone or continued while only simulating; a test batch always can."""
    import json
    from appfacturas import config
    real = tmp_path / "real"
    prueba = tmp_path / "prueba"
    for d, modo in ((real, "real"), (prueba, "prueba")):
        d.mkdir()
        (d / "estado.json").write_text(json.dumps({"estado": "confirmado", "modo": modo}), encoding="utf-8")
    monkeypatch.setattr(config, "cargar", lambda *a, **k: config.Config(modo_real=False))
    assert not ui.puede_tocar(real) and ui.puede_tocar(prueba)
    monkeypatch.setattr(config, "cargar", lambda *a, **k: config.Config(modo_real=True))
    assert ui.puede_tocar(real)


def test_boton_no_permitido_queda_apagado():
    """mientras_libre(..., permitido=False) never binds (the binding used to turn a disabled button back on)."""
    class Boton:
        enabled, binds, tip = True, 0, ""

        def disable(self):
            self.enabled = False

        def tooltip(self, t):
            self.tip = t

        def bind_enabled_from(self, *a, **k):
            self.binds += 1
            return self
    b = ui.mientras_libre(Boton(), False, "no")
    assert (b.enabled, b.binds, b.tip) == (False, 0, "no")
