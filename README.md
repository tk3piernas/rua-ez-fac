<p align="center"><img src="app/appfacturas/static/logo.png" width="96" alt=""></p>

<h1 align="center">RUA EZ-Fac</h1>

<p align="center">Carga de facturas de proveedores de RUA Asistencia (Vehicular y Hogar) en <code>DEUDA A PAGAR ARG.xlsx</code>.</p>

RUA EZ-Fac lee los PDF, zip y .txt que mandan los proveedores. Detecta duplicados y te muestra todo para revisar antes de cargar. Cuando confirmás:

- renombra y guarda cada factura en la carpeta de su proveedor, dentro del OneDrive de RUA;
- da de alta los proveedores nuevos;
- escribe las filas en la planilla, esté abierta o cerrada.

Cada lote se puede deshacer. Reemplaza al script `OrdenarFacturas.ps1`.

- **Guía de uso (1 página):** [docs/Guia de uso RUA EZ-Fac.pdf](docs/Guia%20de%20uso%20RUA%20EZ-Fac.pdf)
- **Documentación técnica:** [docs/DOCUMENTACION.md](docs/DOCUMENTACION.md)

## Qué necesitás

- Windows con Excel de escritorio.
- El OneDrive de RUA sincronizado en la PC, con:
  - `Facturas pendientes`, que adentro tiene `_Entrada\Vehicular`, `_Entrada\Hogar`, `Vehicular`, `Hogar` y `_Script\proveedores.csv`;
  - `DEUDA A PAGAR ARG.xlsx` en la raíz, con las hojas Vehicular, Hogar y Proveedores.
- Python 3.15. Si no lo tenés, el instalador te lo consigue.

## Instalar

1. En esta página: **Code → Download ZIP**.
2. Descomprimilo en **Documentos**, por ejemplo `C:\Users\<usuario>\Documents\RUA EZ-Fac`. **Nunca adentro de OneDrive**: el OneDrive de RUA no deja guardar programas.
3. Si Windows marca los archivos como bajados de internet, hacé clic derecho en el zip **antes de descomprimir** → Propiedades → tildá **Desbloquear** → Aceptar.
4. Doble clic en **`Instalar o actualizar RUA EZ-Fac.bat`**.
   - Si aparece "Windows protegió su PC", tocá *Más información* → *Ejecutar de todas formas*.
   - La primera vez tarda unos minutos.
   - Al final te muestra una lista de OK / FALTA y queda el ícono **RUA EZ-Fac** en el Escritorio.

La primera vez la app arranca con la **escritura real apagada**: sobre tus carpetas y tu planilla solo lee y simula. Para practicar, usá Configuración → *Trabajar sobre la carpeta de prueba*. Cuando estés seguro, prendé *Habilitar escritura real*.

## Actualizar

- Bajá el ZIP nuevo.
- Copiá su contenido encima de tu carpeta. Se conservan `config.toml`, `datos\` y `.venv\`.
- Abrí el ícono. Si había una versión abierta, se cierra sola.
- Solo si cambió algún `requirements*.txt`, corré de nuevo el instalador.

## Qué hay en el repo

```
Instalar o actualizar RUA EZ-Fac.bat   instala / actualiza / crea los íconos / revisa la PC
Abrir RUA EZ-Fac (con consola).bat     abre la app mostrando lo que pasa (para diagnosticar)
Armar instalador para otra PC.bat      arma un zip de la instalación actual para pasarlo
Simular (CLI).bat                      simula _Entrada sin interfaz y deja un informe en datos\
requirements*.txt                      librerías con versión fija (ver docs para el --no-deps)
app/appfacturas/                       código (Python + NiceGUI)
app/tests/                             pruebas que corren sin facturas reales
docs/                                  guía de uso y documentación técnica
```

Lo que crea cada instalación queda fuera del repo: `.venv\`, `datos\` (lotes, informes, respaldos), `prueba\` y `config.toml`.

## Privacidad

- El repo **no** incluye facturas, CUIT ni CBU de proveedores.
- La batería completa de pruebas, que usa facturas reales, y la prueba de paridad contra el script viejo quedan solo en la PC de desarrollo (`app/tests/fixtures/` y `paridad/` están en `.gitignore`).
- Nada sale de la PC: la app corre en `127.0.0.1` y escribe en el OneDrive sincronizado.

## Desarrollo

```bat
.venv\Scripts\python.exe -m pytest -q app\tests
.venv\Scripts\python.exe -m appfacturas --help
```

Uso interno de RUA Asistencia. Contacto: Matheo Kouroussis (Alianzas).
