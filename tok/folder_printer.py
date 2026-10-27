"""Send every top-level file in a selected folder to the default printer."""

from __future__ import annotations

import os
import shutil
import subprocess
import time
from pathlib import Path
from typing import Callable


class PdfPrintUnavailable(RuntimeError):
    """Raised when Windows has no command capable of printing PDF files."""


class PrintCancelled(RuntimeError):
    """Raised internally when a running backend is cancelled."""


def _sumatra_executable() -> str | None:
    candidates = [
        shutil.which("SumatraPDF"),
        os.path.join(os.environ.get("LOCALAPPDATA", ""), "SumatraPDF", "SumatraPDF.exe"),
        r"C:\Program Files\SumatraPDF\SumatraPDF.exe",
        r"C:\Program Files (x86)\SumatraPDF\SumatraPDF.exe",
    ]
    return next((path for path in candidates if path and os.path.isfile(path)), None)


def _print_with_windows(path: str, is_cancelled: Callable[[], bool]) -> None:
    """Use SumatraPDF for PDFs, falling back to the registered shell verb."""
    if Path(path).suffix.lower() == ".pdf":
        sumatra = _sumatra_executable()
        if sumatra:
            process = subprocess.Popen(
                [sumatra, "-print-to-default", "-silent", path],
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
            )
            deadline = time.monotonic() + 120
            while process.poll() is None:
                if is_cancelled():
                    process.terminate()
                    try:
                        process.wait(timeout=3)
                    except subprocess.TimeoutExpired:
                        process.kill()
                    raise PrintCancelled()
                if time.monotonic() >= deadline:
                    process.kill()
                    raise TimeoutError("SumatraPDF svaraði ekki innan 120 sekúndna.")
                time.sleep(0.1)
            stdout, stderr = process.communicate()
            completed_returncode = process.returncode
            if completed_returncode != 0:
                detail = (stderr or stdout or "óþekkt villa").strip()
                raise OSError(f"SumatraPDF gat ekki prentað skrána: {detail}")
            return
    try:
        os.startfile(path, "print")
    except OSError as exc:
        if Path(path).suffix.lower() == ".pdf" and getattr(exc, "winerror", None) == 1155:
            raise PdfPrintUnavailable(
                "Windows er ekki með PDF-prentforrit tengt við Print-aðgerðina. "
                "Settu upp SumatraPDF eða skráðu PDF-lesara sem styður Print og reyndu aftur."
            ) from exc
        raise


def printable_files(folder: str | os.PathLike) -> list[Path]:
    """Return top-level files ordered from oldest-created to newest-created."""
    directory = Path(folder)
    if not directory.is_dir():
        raise ValueError("Valin mappa fannst ekki.")

    def creation_key(path: Path):
        try:
            return (path.stat().st_ctime_ns, path.name.casefold())
        except OSError:
            return (0, path.name.casefold())

    return sorted((path for path in directory.iterdir() if path.is_file()), key=creation_key)


def print_folder_files(
    folder: str | os.PathLike,
    *,
    emit: Callable[[str], None] | None = None,
    print_file: Callable[[str], object] | None = None,
    is_cancelled: Callable[[], bool] | None = None,
    delay_seconds: float = 0.5,
) -> dict:
    """Send every file to the Windows default printer and report each result."""
    emit = emit or (lambda _message: None)
    if os.name != "nt" and print_file is None:
        raise RuntimeError("Möppuprentun er aðeins studd á Windows.")
    is_cancelled = is_cancelled or (lambda: False)
    print_file = print_file or (lambda path: _print_with_windows(path, is_cancelled))
    files = printable_files(folder)
    if not files:
        raise ValueError("Engar skrár fundust í valinni möppu.")

    printed = 0
    failed: list[dict[str, str]] = []
    cancelled = False
    emit(f"Folder: {Path(folder)}")
    emit(f"Files found: {len(files)}")
    emit("Printer: Windows default printer")
    emit("")

    for number, path in enumerate(files, start=1):
        if is_cancelled():
            cancelled = True
            emit(f"Printing cancelled before file {number} of {len(files)}.")
            break
        try:
            print_file(str(path))
            printed += 1
            emit(f"[{number}/{len(files)}] Sent to printer: {path.name}")
        except PrintCancelled:
            cancelled = True
            emit(f"Printing cancelled while processing file {number} of {len(files)}.")
            break
        except PdfPrintUnavailable:
            # This is a machine-wide setup problem; repeating it for every PDF
            # only floods the log and cannot succeed during the same run.
            raise
        except Exception as exc:
            failed.append({"file": str(path), "error": str(exc)})
            emit(f"[{number}/{len(files)}] Could not print: {path.name} — {exc}")
        if delay_seconds > 0 and number < len(files):
            time.sleep(delay_seconds)

    emit("")
    emit("Summary")
    emit(f"Files found: {len(files)}")
    emit(f"Sent to printer: {printed}")
    emit(f"Failed: {len(failed)}")
    emit(f"Cancelled: {'yes' if cancelled else 'no'}")
    return {
        "folder": str(Path(folder)),
        "total": len(files),
        "printed": printed,
        "failed": failed,
        "cancelled": cancelled,
    }
