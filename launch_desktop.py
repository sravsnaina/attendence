
"""Desktop launcher for Attendance Manager.
Keeps the existing app.py and HTML files unchanged.
"""
from __future__ import annotations

import os
import sys
import threading
from pathlib import Path

from werkzeug.serving import make_server
import webview


def user_data_dir() -> Path:
    if os.name == "nt":
        base = Path(
            os.environ.get("LOCALAPPDATA", Path.home() / "AppData" / "Local")
        )
    else:
        base = Path(
            os.environ.get("XDG_DATA_HOME", Path.home() / ".local" / "share")
        )

    path = base / "AttendanceManager" / "data"
    path.mkdir(parents=True, exist_ok=True)
    return path


def prepare_bundled_tesseract() -> None:
    if getattr(sys, "frozen", False):
        app_dir = Path(sys.executable).resolve().parent
    else:
        app_dir = Path(__file__).resolve().parent

    tess_dir = app_dir / "tesseract"

    if tess_dir.is_dir():
        os.environ["PATH"] = (
            str(tess_dir) + os.pathsep + os.environ.get("PATH", "")
        )

        tessdata = tess_dir / "tessdata"
        if tessdata.is_dir():
            os.environ["TESSDATA_PREFIX"] = str(tessdata)


def main() -> None:
    prepare_bundled_tesseract()

    # Import the existing application without modifying app.py.
    import app as attendance_app

    data_dir = user_data_dir()
    flask_app = attendance_app.create_app(data_dir=str(data_dir))

    server = make_server("127.0.0.1", 0, flask_app, threaded=True)
    port = server.server_port

    thread = threading.Thread(
        target=server.serve_forever,
        daemon=True,
    )
    thread.start()

    try:
        webview.create_window(
            "Attendance Manager",
            f"http://127.0.0.1:{port}",
            width=1280,
            height=820,
            min_size=(960, 640),
        )
        webview.start()
    finally:
        server.shutdown()


if __name__ == "__main__":
    main()
