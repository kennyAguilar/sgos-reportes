import os
import sys
import threading
import webview

# Fijar directorio de trabajo al lado del .exe (o del script en dev)
if getattr(sys, 'frozen', False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
os.chdir(BASE_DIR)

# Indicar a Flask que estamos en modo desktop (HTTP local, sin HTTPS)
os.environ["SGOS_DESKTOP"] = "1"

from sgos_web.app import app

PORT = 5000


def start_flask():
    app.run(port=PORT, use_reloader=False)


if __name__ == "__main__":
    t = threading.Thread(target=start_flask, daemon=True)
    t.start()

    window = webview.create_window(
        "SGOS Reportes",
        f"http://127.0.0.1:{PORT}",
        width=1280,
        height=800,
    )
    # private_mode=False permite persistir cookies/sesión y habilita descargas
    webview.start(private_mode=False)
