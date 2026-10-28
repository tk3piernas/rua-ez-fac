"""The app's own identity in the Windows taskbar.

The window is Edge in app mode, so Windows groups it with Edge and pinning it pins Edge. The fix is the
AppUserModelID: the shortcuts (Desktop and Start menu) carry "RUA.EZFac", and as soon as the window opens it
gets the same ID plus a relaunch command and icon. Then the running window groups with the pinned icon, and
"Pin to taskbar" pins RUA EZ-Fac (its icon, and opening it starts the app, not a bare Edge)."""
from __future__ import annotations

import sys
import threading
import time
from pathlib import Path

AUMID = "RUA.EZFac"
NOMBRE = "RUA EZ-Fac"
ICONO = Path(__file__).parent / "static" / "ezfac.ico"


def pythonw() -> Path:
    """The windowless Python of the app's .venv (what the shortcut runs)."""
    exe = Path(sys.executable)
    w = exe.with_name("pythonw.exe")
    return w if w.exists() else exe


def comando() -> str:
    return f'"{pythonw()}" -m appfacturas ui'


def _variant(texto: str):
    import pythoncom
    from win32comext.propsys import propsys
    return propsys.PROPVARIANTType(texto, pythoncom.VT_LPWSTR)


def marcar(hwnd: int) -> None:
    """Gives one window the app's identity, relaunch command, name and icon."""
    from win32comext.propsys import propsys, pscon
    store = propsys.SHGetPropertyStoreForWindow(hwnd, propsys.IID_IPropertyStore)
    # relaunch properties first: Windows reads them when the ID is set
    store.SetValue(pscon.PKEY_AppUserModel_RelaunchCommand, _variant(comando()))
    store.SetValue(pscon.PKEY_AppUserModel_RelaunchDisplayNameResource, _variant(NOMBRE))
    store.SetValue(pscon.PKEY_AppUserModel_RelaunchIconResource, _variant(f"{ICONO},0"))
    store.SetValue(pscon.PKEY_AppUserModel_ID, _variant(AUMID))
    store.Commit()


def id_de(hwnd: int) -> str:
    from win32comext.propsys import propsys, pscon
    store = propsys.SHGetPropertyStoreForWindow(hwnd, propsys.IID_IPropertyStore)
    return store.GetValue(pscon.PKEY_AppUserModel_ID).GetValue() or ""


def ventanas() -> list[int]:
    """Visible top-level windows titled like the app (the Edge app window is titled with the page title)."""
    import win32gui
    out = []

    def ver(h, _):
        if win32gui.IsWindowVisible(h) and win32gui.GetWindowText(h).strip() == NOMBRE:
            out.append(h)
        return True
    win32gui.EnumWindows(ver, None)
    return out


def marcar_cuando_aparezca(espera: float = 25.0) -> None:
    """Waits (in the background) for the app window and marks it. Does nothing outside Windows."""
    if sys.platform != "win32":
        return

    def correr():
        import pythoncom
        pythoncom.CoInitialize()
        hecho, fin = set(), time.time() + espera
        while time.time() < fin:
            for h in ventanas():
                if h not in hecho:
                    try:
                        marcar(h)
                        hecho.add(h)
                    except Exception:
                        pass
            if hecho:
                return
            time.sleep(0.5)
    threading.Thread(target=correr, daemon=False).start()   # not daemon: also when the launcher ends right away


_candado = None                     # kept for the whole life of the process that starts the server


def tomar_candado(puerto: int = 8765) -> bool:
    """True if this is the only launcher starting the app right now (a named mutex). Several quick clicks on
    the icon start several launchers: only the first starts the server, the rest wait and show its window."""
    global _candado
    if sys.platform != "win32":
        return True
    try:
        import win32api
        import win32event
        import winerror
        # one lock per port: another installation (or the test one) on another port is another app
        _candado = win32event.CreateMutex(None, False, f"Local\\RUA.EZFac.inicio.{puerto}")
        return win32api.GetLastError() != winerror.ERROR_ALREADY_EXISTS
    except Exception:
        return True


def traer_al_frente(hwnd: int) -> bool:
    """Restores the window if minimized and brings it to the front."""
    try:
        import win32con
        import win32gui
        if win32gui.IsIconic(hwnd):
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
        else:
            win32gui.ShowWindow(hwnd, win32con.SW_SHOW)
        win32gui.SetForegroundWindow(hwnd)
        return True
    except Exception:
        try:
            import win32gui
            win32gui.BringWindowToTop(hwnd)
            return True
        except Exception:
            return False


def mostrar_la_abierta() -> bool:
    """If the app already has a window, brings it to the front instead of opening another one."""
    if sys.platform != "win32":
        return False
    try:
        hs = ventanas()
    except Exception:
        return False
    return bool(hs) and traer_al_frente(hs[0])


def acceso_directo(ruta: Path, carpeta_app: Path) -> None:
    """A .lnk that starts the app, with its icon and the same AppUserModelID as the window."""
    import pythoncom
    from win32comext.propsys import propsys, pscon
    from win32comext.shell import shell
    link = pythoncom.CoCreateInstance(shell.CLSID_ShellLink, None, pythoncom.CLSCTX_INPROC_SERVER,
                                      shell.IID_IShellLink)
    link.SetPath(str(pythonw()))
    link.SetArguments("-m appfacturas ui")
    link.SetWorkingDirectory(str(carpeta_app))
    link.SetDescription("RUA EZ-Fac: carga de facturas de proveedores")
    link.SetIconLocation(str(ICONO), 0)
    store = link.QueryInterface(propsys.IID_IPropertyStore)
    store.SetValue(pscon.PKEY_AppUserModel_ID, _variant(AUMID))
    store.Commit()
    ruta.parent.mkdir(parents=True, exist_ok=True)
    link.QueryInterface(pythoncom.IID_IPersistFile).Save(str(ruta), 0)


def crear_accesos(carpeta_app: Path, log=print) -> list[Path]:
    """Desktop and Start menu shortcuts (replaces the old 'App Facturas' ones)."""
    from win32comext.shell import shell, shellcon
    hechos = []
    for csidl in (shellcon.CSIDL_PROGRAMS, shellcon.CSIDL_DESKTOPDIRECTORY):
        d = Path(shell.SHGetFolderPath(0, csidl, None, 0))
        viejo = d / "App Facturas.lnk"
        if viejo.exists():
            viejo.unlink()
        ruta = d / f"{NOMBRE}.lnk"
        try:
            acceso_directo(ruta, carpeta_app)
            hechos.append(ruta)
            log(f"      {ruta}")
        except Exception as e:
            log(f"      No pude crear el icono en {d}: {e}")
    return hechos


def id_de_acceso(ruta: Path) -> str:
    import pythoncom
    from win32comext.propsys import propsys, pscon
    from win32comext.shell import shell
    link = pythoncom.CoCreateInstance(shell.CLSID_ShellLink, None, pythoncom.CLSCTX_INPROC_SERVER,
                                      shell.IID_IShellLink)
    link.QueryInterface(pythoncom.IID_IPersistFile).Load(str(ruta), 0)
    return link.QueryInterface(propsys.IID_IPropertyStore).GetValue(pscon.PKEY_AppUserModel_ID).GetValue() or ""
