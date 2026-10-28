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
