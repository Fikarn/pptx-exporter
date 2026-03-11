"""Vendored tkdnd 2.9.5 loader and minimal Python DnD wrapper.

Provides file drag-and-drop for macOS (Tcl 9) and Windows (Tcl 8.6)
using vendored tkdnd binaries — no external Python package required.
"""

import logging
import os
import platform
import tkinter

logger = logging.getLogger(__name__)

DND_FILES = "DND_Files"

_loaded = False


def _vendored_dir() -> str | None:
    """Return path to the vendored tkdnd directory for this platform."""
    system = platform.system()
    machine = platform.machine()
    base = os.path.dirname(__file__)

    if system == "Darwin":
        sub = "macos-arm64" if machine == "arm64" else "macos-x86_64"
    elif system == "Windows":
        sub = "windows-x64"
    else:
        return None

    path = os.path.join(base, sub)
    return path if os.path.isdir(path) else None


def init_dnd(tkroot) -> bool:
    """Load tkdnd into the Tk interpreter using vendored binaries.

    Call once after the Tk root window has been created.
    Returns True on success, False if tkdnd is unavailable.
    """
    global _loaded
    tkdnd_dir = _vendored_dir()
    if not tkdnd_dir:
        logger.debug("No vendored tkdnd binary for this platform")
        return False
    try:
        tkroot.tk.call("lappend", "auto_path", tkdnd_dir)
        version = tkroot.tk.call("package", "require", "tkdnd")
        _loaded = True
        logger.debug("tkdnd %s loaded from %s", version, tkdnd_dir)
        return True
    except tkinter.TclError as exc:
        logger.debug("Failed to load tkdnd: %s", exc)
        return False


def dnd_available() -> bool:
    """Return True if tkdnd has been successfully loaded."""
    return _loaded


def register_drop_target(widget) -> None:
    """Register *widget* as a file drop target."""
    widget.tk.call("tkdnd::drop_target", "register", widget._w, DND_FILES)


def bind_drop(widget, callback) -> None:
    """Bind <<Drop>> on *widget*.

    *callback* receives a tuple of file-path strings.
    """
    def _on_drop(data):
        try:
            paths = widget.tk.splitlist(data)
        except tkinter.TclError:
            paths = (data,)
        callback(paths)
        return "copy"

    cmd = widget._register(_on_drop)
    widget.tk.call("bind", widget._w, "<<Drop>>", f"{cmd} %D")


def bind_drop_enter(widget, callback) -> None:
    """Bind <<DropEnter>> on *widget*. *callback* takes no arguments."""
    def _on_enter(*_args):
        callback()
        return "copy"

    cmd = widget._register(_on_enter)
    widget.tk.call("bind", widget._w, "<<DropEnter>>", cmd)


def bind_drop_leave(widget, callback) -> None:
    """Bind <<DropLeave>> on *widget*. *callback* takes no arguments."""
    cmd = widget._register(lambda *_args: callback())
    widget.tk.call("bind", widget._w, "<<DropLeave>>", cmd)
