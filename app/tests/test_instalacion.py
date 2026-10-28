import zipfile

from appfacturas import instalacion


def test_empaquetar_no_lleva_pruebas_ni_datos(tmp_path, monkeypatch):
    raiz = tmp_path / "RUA EZ-Fac"
    (raiz / "app/appfacturas/__pycache__").mkdir(parents=True)
    (raiz / "app/appfacturas/ui.py").write_text("x")
    (raiz / "app/appfacturas/__pycache__/ui.pyc").write_text("x")
    (raiz / "app/tests").mkdir()
    (raiz / "app/tests/factura_real.pdf").write_text("cbu")
    (raiz / "datos").mkdir()
    (raiz / "datos/app.log").write_text("x")
    (raiz / "config.toml").write_text("x")
    (raiz / "requirements.txt").write_text("x")
    (raiz / "Instalar o actualizar RUA EZ-Fac.bat").write_text("x")
    monkeypatch.setenv("APPFACTURAS_HOME", str(raiz))
    z = instalacion.empaquetar(tmp_path / "p.zip")
    nombres = zipfile.ZipFile(z).namelist()
    assert "RUA EZ-Fac/app/appfacturas/ui.py" in nombres
    assert "RUA EZ-Fac/Instalar o actualizar RUA EZ-Fac.bat" in nombres
    assert not any("tests" in n or "datos" in n or "config.toml" in n or ".pyc" in n for n in nombres)
