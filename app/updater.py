"""Small external updater for source-based Urdu Unicoder installations."""

from __future__ import annotations

import os
import shutil
import subprocess
import sys
import tempfile
import time
import traceback
import urllib.request
import zipfile
from pathlib import Path


SOURCE_ARCHIVE = "https://github.com/MianAshfaq/Urdu-Unicoder/archive/refs/heads/main.zip"
COPY_DIRECTORIES = ("app", "assets")
COPY_FILES = (
    ".gitignore", "README.md", "LICENSE.txt", "requirements.txt", "version.json",
    "setup_windows.bat", "run_windows.bat", "build_exe_windows.bat",
    "test_unicode_recovery.py",
)


def wait_for_process(process_id: int):
    if sys.platform == "win32":
        try:
            import ctypes
            synchronize = 0x00100000
            handle = ctypes.windll.kernel32.OpenProcess(synchronize, False, process_id)
            if handle:
                ctypes.windll.kernel32.WaitForSingleObject(handle, 20000)
                ctypes.windll.kernel32.CloseHandle(handle)
                return
        except Exception:
            pass
    time.sleep(3)


def safe_extract(archive: Path, destination: Path):
    destination = destination.resolve()
    with zipfile.ZipFile(archive) as package:
        for member in package.infolist():
            target = (destination / member.filename).resolve()
            if destination not in target.parents and target != destination:
                raise RuntimeError("The update package contains an unsafe path.")
        package.extractall(destination)


def copy_update(source_root: Path, destination_root: Path):
    for directory in COPY_DIRECTORIES:
        source = source_root / directory
        if source.exists():
            shutil.copytree(
                source,
                destination_root / directory,
                dirs_exist_ok=True,
                ignore=shutil.ignore_patterns("__pycache__", "*.pyc"),
            )
    for filename in COPY_FILES:
        source = source_root / filename
        if source.exists():
            shutil.copy2(source, destination_root / filename)


def show_error(message: str):
    if sys.platform == "win32":
        try:
            import ctypes
            ctypes.windll.user32.MessageBoxW(
                0, message, "Urdu Unicoder update failed", 0x10
            )
        except Exception:
            pass


def main() -> int:
    if len(sys.argv) != 3:
        return 2
    destination_root = Path(sys.argv[1]).resolve()
    process_id = int(sys.argv[2])
    log_path = destination_root / "update.log"
    if not (destination_root / "app" / "main.py").exists():
        raise RuntimeError("The Urdu Unicoder installation folder is not valid.")

    wait_for_process(process_id)
    try:
        with log_path.open("w", encoding="utf-8") as log:
            log.write("Downloading Urdu Unicoder update...\n")
            with tempfile.TemporaryDirectory(prefix="urdu-unicoder-update-") as temp_name:
                temp_dir = Path(temp_name)
                archive = temp_dir / "update.zip"
                request = urllib.request.Request(
                    SOURCE_ARCHIVE, headers={"User-Agent": "Urdu-Unicoder-Updater"}
                )
                with urllib.request.urlopen(request, timeout=60) as response:
                    archive.write_bytes(response.read())
                extract_dir = temp_dir / "extracted"
                extract_dir.mkdir()
                safe_extract(archive, extract_dir)
                roots = [item for item in extract_dir.iterdir() if item.is_dir()]
                if len(roots) != 1 or not (roots[0] / "app" / "main.py").exists():
                    raise RuntimeError("The downloaded update package is not valid.")
                copy_update(roots[0], destination_root)
                log.write("Application files updated.\n")

            requirements = destination_root / "requirements.txt"
            if requirements.exists():
                result = subprocess.run(
                    [sys.executable, "-m", "pip", "install", "-r", str(requirements)],
                    cwd=str(destination_root),
                    stdout=log,
                    stderr=subprocess.STDOUT,
                    timeout=300,
                    check=False,
                )
                if result.returncode:
                    raise RuntimeError("A required package could not be updated. See update.log.")
            log.write("Update completed. Restarting Urdu Unicoder.\n")

        restart = [sys.executable, str(destination_root / "app" / "main.py")]
        kwargs = {"cwd": str(destination_root)}
        if sys.platform == "win32":
            kwargs["creationflags"] = subprocess.CREATE_NO_WINDOW
        subprocess.Popen(restart, **kwargs)
        return 0
    except Exception:
        details = traceback.format_exc()
        try:
            log_path.write_text(details, encoding="utf-8")
        except Exception:
            pass
        show_error("The update could not be installed. Your project files were not changed.\n\nSee update.log for details.")
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
