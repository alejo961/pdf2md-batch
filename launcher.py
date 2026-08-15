"""Windowless Windows launcher for the locally hosted PDF2MD application."""

from __future__ import annotations

import ctypes
import faulthandler
import json
import logging
import os
import socket
import sys
import threading
import time
import urllib.request
import webbrowser
from logging.handlers import RotatingFileHandler
from pathlib import Path

from PIL import Image
import pystray
from waitress.server import create_server

from app import app, initialize_application
from runtime_config import (
    APP_NAME,
    APP_VERSION,
    DEFAULT_PORT,
    LOG_DIR,
    RUNTIME_FILE,
    configure_tesseract,
    ensure_runtime_directories,
    resource_root,
)


MUTEX_NAME = r"Local\PDF2MD_Desktop_Application"
ERROR_ALREADY_EXISTS = 183
server = None
server_thread = None
tray_icon = None
mutex_handle = None
runtime_url = None
crash_log_stream = None
shutdown_requested = threading.Event()
server_failed = threading.Event()
launcher_logger = logging.getLogger("pdf2md.launcher")


def configure_launcher_logging() -> None:
    global crash_log_stream
    ensure_runtime_directories()
    launcher_logger.setLevel(logging.INFO)
    launcher_logger.propagate = False
    if not launcher_logger.handlers:
        handler = RotatingFileHandler(
            LOG_DIR / "launcher.log",
            maxBytes=2 * 1024 * 1024,
            backupCount=3,
            encoding="utf-8",
        )
        handler.setFormatter(logging.Formatter(
            "%(asctime)s | %(levelname)s | %(threadName)s | %(message)s"
        ))
        launcher_logger.addHandler(handler)

    crash_log_stream = (LOG_DIR / "crash.log").open("a", encoding="utf-8")
    faulthandler.enable(file=crash_log_stream, all_threads=True)
    launcher_logger.info(
        "Inicio PDF2MD %s | frozen=%s | executable=%s | cwd=%s",
        APP_VERSION,
        bool(getattr(sys, "frozen", False)),
        sys.executable,
        os.getcwd(),
    )


def acquire_single_instance() -> bool:
    global mutex_handle
    kernel32 = ctypes.windll.kernel32
    mutex_handle = kernel32.CreateMutexW(None, False, MUTEX_NAME)
    return bool(mutex_handle) and kernel32.GetLastError() != ERROR_ALREADY_EXISTS


def read_existing_url() -> str | None:
    try:
        data = json.loads(RUNTIME_FILE.read_text(encoding="utf-8"))
        url = str(data.get("url") or "")
        if url.startswith("http://127.0.0.1:") and is_healthy(url):
            return url
    except (OSError, ValueError, TypeError):
        return None
    return None


def find_available_port(preferred: int = DEFAULT_PORT) -> int:
    for port in (preferred, 0):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as probe:
            try:
                probe.bind(("127.0.0.1", port))
                return int(probe.getsockname()[1])
            except OSError:
                continue
    raise RuntimeError("Windows no pudo asignar un puerto local para PDF2MD.")


def is_healthy(base_url: str) -> bool:
    try:
        with urllib.request.urlopen(base_url + "/health", timeout=1.5) as response:
            payload = json.loads(response.read().decode("utf-8"))
            return response.status == 200 and payload.get("status") == "ok"
    except (OSError, ValueError):
        return False


def wait_until_ready(base_url: str, timeout: float = 25.0) -> None:
    deadline = time.time() + timeout
    while time.time() < deadline:
        if is_healthy(base_url):
            return
        time.sleep(0.25)
    raise RuntimeError("PDF2MD tardó demasiado en iniciar. Consulta el diagnóstico para obtener ayuda.")


def write_runtime_file(base_url: str) -> None:
    temporary = RUNTIME_FILE.with_suffix(".json.tmp")
    temporary.write_text(json.dumps({
        "pid": os.getpid(),
        "url": base_url,
        "version": APP_VERSION,
        "started_at": time.time(),
    }, ensure_ascii=False, indent=2), encoding="utf-8")
    os.replace(temporary, RUNTIME_FILE)


def open_application(_icon=None, _item=None) -> None:
    if runtime_url:
        webbrowser.open(runtime_url, new=2)


def stop_application(_icon=None, _item=None) -> None:
    launcher_logger.info("Cierre solicitado por el usuario")
    shutdown_requested.set()
    try:
        if server is not None:
            server.close()
    finally:
        try:
            if RUNTIME_FILE.exists():
                RUNTIME_FILE.unlink()
        except OSError:
            pass
        if tray_icon is not None:
            tray_icon.stop()


def tray_image() -> Image.Image:
    icon_path = resource_root() / "assets" / "pdf2md.ico"
    if icon_path.is_file():
        with Image.open(icon_path) as source:
            return source.convert("RGBA").copy()
    return Image.new("RGBA", (64, 64), (255, 107, 53, 255))


def show_startup_error(message: str) -> None:
    launcher_logger.error("Error de inicio: %s", message)
    detail = (
        f"{message}\n\n"
        f"Diagnóstico: {LOG_DIR}\n\n"
        "Selecciona Sí para copiar esta información al portapapeles."
    )
    result = ctypes.windll.user32.MessageBoxW(
        None,
        detail,
        f"{APP_NAME} - No fue posible iniciar",
        0x00000004 | 0x00000010,
    )
    if result == 6:
        copy_to_clipboard(detail)


def copy_to_clipboard(text: str) -> None:
    """Copy Unicode text using the native Windows clipboard API."""
    CF_UNICODETEXT = 13
    GMEM_MOVEABLE = 0x0002
    user32 = ctypes.windll.user32
    kernel32 = ctypes.windll.kernel32
    kernel32.GlobalAlloc.restype = ctypes.c_void_p
    kernel32.GlobalLock.restype = ctypes.c_void_p
    kernel32.GlobalLock.argtypes = [ctypes.c_void_p]
    kernel32.GlobalUnlock.argtypes = [ctypes.c_void_p]
    user32.SetClipboardData.argtypes = [ctypes.c_uint, ctypes.c_void_p]
    if not user32.OpenClipboard(None):
        return
    try:
        user32.EmptyClipboard()
        data = (text + "\0").encode("utf-16-le")
        handle = kernel32.GlobalAlloc(GMEM_MOVEABLE, len(data))
        if not handle:
            return
        pointer = kernel32.GlobalLock(handle)
        if not pointer:
            return
        ctypes.memmove(pointer, data, len(data))
        kernel32.GlobalUnlock(handle)
        user32.SetClipboardData(CF_UNICODETEXT, handle)
    finally:
        user32.CloseClipboard()


def run_server() -> None:
    """Run Waitress and record unexpected native or Python-level termination."""
    try:
        server.run()
        if not shutdown_requested.is_set():
            launcher_logger.error("Waitress terminó sin que el usuario solicitara el cierre")
            server_failed.set()
    except BaseException:
        launcher_logger.exception("El servidor local terminó inesperadamente")
        server_failed.set()


def monitor_server() -> None:
    """Release the tray loop if the HTTP server terminates unexpectedly."""
    if server_thread is None:
        return
    server_thread.join()
    if server_failed.is_set() and tray_icon is not None:
        try:
            tray_icon.stop()
        except Exception:
            launcher_logger.exception("No se pudo detener el icono de bandeja")


def run() -> int:
    global server, server_thread, tray_icon, runtime_url
    configure_launcher_logging()
    if not acquire_single_instance():
        existing_url = read_existing_url()
        if existing_url:
            webbrowser.open(existing_url, new=2)
            return 0
        show_startup_error("PDF2MD ya se está iniciando. Espera unos segundos y vuelve a intentarlo.")
        return 1

    try:
        tessdata = configure_tesseract()
        if tessdata is None:
            raise RuntimeError("No se encontraron los datos OCR en español e inglés incluidos con la aplicación.")

        initialize_application()
        port = find_available_port()
        runtime_url = f"http://127.0.0.1:{port}"
        server = create_server(app, host="127.0.0.1", port=port, threads=6)
        server_thread = threading.Thread(target=run_server, daemon=False, name="pdf2md-server")
        server_thread.start()
        wait_until_ready(runtime_url)
        write_runtime_file(runtime_url)
        launcher_logger.info("Servidor disponible en %s", runtime_url)
        open_application()

        tray_icon = pystray.Icon(
            "PDF2MD",
            tray_image(),
            f"PDF2MD {APP_VERSION}",
            menu=pystray.Menu(
                pystray.MenuItem("Abrir PDF2MD", open_application, default=True),
                pystray.MenuItem("Cerrar PDF2MD", stop_application),
            ),
        )
        threading.Thread(
            target=monitor_server,
            daemon=True,
            name="pdf2md-server-monitor",
        ).start()
        try:
            tray_icon.run()
        except Exception:
            # The tray is a convenience. Some locked-down Windows profiles do
            # not allow notification icons; the local server must stay alive.
            launcher_logger.exception(
                "El icono de bandeja falló; PDF2MD continuará disponible en el navegador"
            )

        if server_thread.is_alive() and not shutdown_requested.is_set():
            launcher_logger.warning(
                "El icono de bandeja terminó, pero el servidor sigue activo"
            )
            server_thread.join()
        if server_failed.is_set():
            raise RuntimeError(
                "El servidor local se detuvo inesperadamente. Abre nuevamente PDF2MD y comparte el diagnóstico con soporte."
            )
        return 0
    except Exception as exc:
        launcher_logger.exception("Fallo fatal de inicio")
        show_startup_error(str(exc))
        return 1
    finally:
        if server is not None and server_thread is not None and server_thread.is_alive():
            if server_failed.is_set() or shutdown_requested.is_set():
                try:
                    server.close()
                    server_thread.join(timeout=10)
                except Exception:
                    launcher_logger.exception("No se pudo cerrar Waitress de forma limpia")
        try:
            if RUNTIME_FILE.exists():
                current = json.loads(RUNTIME_FILE.read_text(encoding="utf-8"))
                if int(current.get("pid") or 0) == os.getpid():
                    RUNTIME_FILE.unlink()
        except (OSError, ValueError, TypeError):
            pass
        launcher_logger.info("Fin del proceso PDF2MD")


if __name__ == "__main__":
    sys.exit(run())
