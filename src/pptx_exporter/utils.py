"""Shared helpers: file validation, logging setup, OS detection."""

import logging
import os
import platform
from pathlib import Path


def configure_logging() -> None:
    """Configure root logger from LOG_LEVEL environment variable."""
    level_name = os.environ.get("LOG_LEVEL", "INFO").upper()
    level = getattr(logging, level_name, logging.INFO)
    logging.basicConfig(
        level=level,
        format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
        datefmt="%Y-%m-%dT%H:%M:%S",
    )


def validate_pptx(path: str) -> Path:
    """Validate that *path* points to an existing .pptx file.

    Returns a resolved :class:`pathlib.Path`.
    Raises :class:`ValueError` on invalid input.
    """
    p = Path(path)
    if not p.exists():
        raise ValueError(f"File not found: {path}")
    if not p.is_file():
        raise ValueError(f"Not a file: {path}")
    if p.suffix.lower() != ".pptx":
        raise ValueError(f"Expected a .pptx file, got: {p.suffix}")
    return p.resolve()


def validate_output_dir(path: str) -> Path:
    """Validate (and create if needed) an output directory.

    Returns a resolved :class:`pathlib.Path`.
    Raises :class:`ValueError` on invalid input.
    """
    p = Path(path)
    try:
        p.mkdir(parents=True, exist_ok=True)
    except OSError as exc:
        raise ValueError(f"Cannot create output directory '{path}': {exc}") from exc
    if not os.access(p, os.W_OK):
        raise ValueError(f"Output directory is not writable: {path}")
    return p.resolve()


def slide_output_name(slide_index: int, total: int) -> str:
    """Return a zero-padded filename for *slide_index* (0-based).

    Example: slide_index=0, total=12  →  "slide_01.png"
    """
    width = len(str(total))
    return f"slide_{slide_index + 1:0{width}d}.png"


def detect_os() -> str:
    """Return 'macos', 'windows', or 'other'."""
    system = platform.system()
    if system == "Darwin":
        return "macos"
    if system == "Windows":
        return "windows"
    return "other"


def is_powerpoint_installed_macos() -> bool:
    """Return True if Microsoft PowerPoint.app is present on macOS."""
    app_path = Path("/Applications/Microsoft PowerPoint.app")
    return app_path.exists()


def is_powerpoint_installed_windows() -> bool:
    """Return True if PowerPoint COM server is registered on Windows."""
    try:
        import win32com.client  # noqa: F401
        import win32api  # noqa: F401

        win32com.client.Dispatch("PowerPoint.Application")
        return True
    except Exception:
        return False


def detect_backend() -> str:
    """Return the backend that will be used: 'macos', 'windows', or 'not_found'.

    Microsoft PowerPoint is required. If it is not found, returns 'not_found'.
    """
    os_name = detect_os()
    if os_name == "macos" and is_powerpoint_installed_macos():
        return "macos"
    if os_name == "windows" and is_powerpoint_installed_windows():
        return "windows"
    return "not_found"


def backend_description(backend: str) -> str:
    """Return a human-readable description of the active backend."""
    descriptions = {
        "macos": "macOS — Microsoft PowerPoint (AppleScript automation)",
        "windows": "Windows — Microsoft PowerPoint (COM automation)",
        "not_found": "PowerPoint not found — please install Microsoft PowerPoint",
    }
    return descriptions.get(backend, backend)
