<p align="center"><img src="app/appfacturas/static/logo.png" width="96" alt=""></p>

<h1 align="center">RUA EZ-Fac</h1>

<p align="center">Carga de facturas de proveedores de RUA Asistencia (Vehicular, Hogar, Persona y ServiciosYa) en <code>DEUDA A PAGAR ARG.xlsx</code>.</p>

RUA EZ-Fac lee los PDF, zip y .txt que mandan los proveedores. Detecta duplicados y te muestra todo para revisar antes de cargar. Cuando confirmás:

- renombra y guarda cada factura en la carpeta de su proveedor, dentro del OneDrive de RUA;
- da de alta los proveedores nuevos;
- escribe las filas en la planilla, esté abierta o cerrada.

Las facturas que se frenan (en rojo en la grilla, por ejemplo una Factura A sin neto) se ven pero no se cargan, y el resto del lote sigue. Cada lote se puede deshacer. Reemplaza al script `OrdenarFacturas.ps1`.

- **Guía de uso (2 páginas):** [docs/Guia de uso RUA EZ-Fac.pdf](docs/Guia%20de%20uso%20RUA%20EZ-Fac.pdf)
- **Documentación técnica:** [docs/DOCUMENTACION.md](docs/DOCUMENTACION.md)

## Qué necesitás

- Windows con Excel de escritorio.
- El OneDrive de RUA sincronizado en la PC, con:
  - `Facturas pendientes`, que adentro tiene `_Entrada\<área>` y `<área>` para cada área (Vehicular, Hogar, Persona y ServiciosYa) y `_Script\proveedores.csv`;
  - `DEUDA A PAGAR ARG.xlsx` en la raíz, con una hoja por área y la hoja Proveedores (si faltan Persona o ServiciosYa, Setup las agrega como copia vacía de las otras);
  - `PROVEEDORES RUA.xlsx` al lado de la planilla: la memoria de datos de proveedores (la crea la app con el primer guardado).
- Python 3.15. Si no lo tenés, el instalador te lo consigue.

¿Te falta alguna carpeta o la planilla? Abrí la app y tocá **Configuración → Setup → Revisar y preparar todo**: crea lo que falte (también una planilla vacía con la misma estructura) y te muestra un chequeo con OK / FALTA.

## Instalar

1. En esta página: **Code → Download ZIP**.
2. Descomprimilo en **Documentos**, por ejemplo `C:\Users\<usuario>\Documents\RUA EZ-Fac`. **Nunca adentro de OneDrive**: el OneDrive de RUA no deja guardar programas.
3. Si Windows marca los archivos como bajados de internet, hacé clic derecho en el zip **antes de descomprimir** → Propiedades → tildá **Desbloquear** → Aceptar.
4. Doble clic en **`Instalar o actualizar RUA EZ-Fac.bat`**.
   - Si aparece "Windows protegió su PC", tocá *Más información* → *Ejecutar de todas formas*.
   - La primera vez tarda unos minutos.
   - Al final te muestra una lista de OK / FALTA y queda el ícono **RUA EZ-Fac** en el Escritorio.

> **¡Esto es una alpha!** Mandá bugs y sugerencias a Matheo.

## Datos de proveedores

Pestaña **Proveedores**: pegá los datos como vengan (mail, WhatsApp, .txt) y tocá *Leer los datos*. La app saca CUIT, CBU/CVU, alias, banco y contacto, chequea el CUIT y el CBU y te muestra la ficha para corregir antes de guardar. Se guardan en `PROVEEDORES RUA.xlsx` (OneDrive, al lado de la planilla), que no se toca con las actualizaciones. Cuando un lote trae un proveedor que no está en la hoja Proveedores de la planilla, la app copia sus datos desde ahí; y los datos que llegan en .txt con un lote se guardan solos. Cada cambio se puede deshacer.

La primera vez la app arranca con la **escritura real apagada**: sobre tus carpetas y tu planilla solo lee y simula. Para practicar, usá Configuración → *Trabajar sobre la carpeta de prueba*. En la carpeta de prueba la app nunca escribe en la planilla real, aunque la tengas abierta en Excel. Cuando estés seguro, prendé *Habilitar escritura real*.

## Abrir y exportar la planilla

Arriba a la derecha: **Abrir planilla** abre en Excel la planilla en la que está escribiendo la app (la real o la de prueba). **Exportar planilla** guarda una copia con el nombre que elijas en `Planillas exportadas`, al lado de la planilla; la planilla sigue igual. Si después querés empezarla de cero, usá el Reset.

## Empezar la planilla de cero

**Configuración → Reset de la planilla…** (hay que escribir BORRAR). Vacía las filas 12–210 de las hojas de cada área en el mismo archivo, así los links compartidos siguen andando: vuelve a poner la fórmula del IVA, no toca la columna I ni el formato, y si querés también vacía Proveedores. Antes guarda dos respaldos.

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
CLAUDE.md                              reglas para Claude Code (qué nunca se sube, cómo probar y publicar)
docs/                                  guía de uso (PDF y su HTML en docs/guia), documentación técnica y prompt de revisión
```

Lo que crea cada instalación queda fuera del repo: `.venv\`, `datos\` (lotes, informes, respaldos), `prueba\` y `config.toml`.

## Privacidad

- El repo **no** incluye facturas, CUIT ni CBU de proveedores.
- La batería completa de pruebas, que usa facturas reales, y la prueba de paridad contra el script viejo quedan solo en la PC de desarrollo (`app/tests_privados/` y `paridad/` están en `.gitignore`).
- Nada sale de la PC: la app corre en `127.0.0.1` y escribe en el OneDrive sincronizado.

## Desarrollo

```bat
.venv\Scripts\python.exe -m pytest -q app\tests
.venv\Scripts\python.exe -m appfacturas --help
```

Uso interno de RUA Asistencia. Contacto: Matheo Kouroussis (Alianzas).
