"""All UI code for pptx-exporter — built with CustomTkinter."""

import json
import logging
import os
import subprocess
import sys
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog
from typing import Optional

import customtkinter as ctk

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    _DND_AVAILABLE = True
except Exception:
    _DND_AVAILABLE = False

from .exporter import Exporter
from .utils import configure_logging

logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Appearance
# ---------------------------------------------------------------------------

ctk.set_appearance_mode("system")       # follows OS light / dark setting
ctk.set_default_color_theme("blue")

# Design tokens — (light, dark)
_COLOR_BG_SURFACE = ("#FFFFFF", "#2C2C2E")
_COLOR_BG_BASE = ("#F5F5F7", "#1C1C1E")
_COLOR_BORDER = ("#D1D1D6", "#3A3A3C")
_COLOR_TEXT_PRIMARY = ("#1D1D1F", "#F5F5F7")
_COLOR_TEXT_SECONDARY = ("#86868B", "#8E8E93")
_COLOR_ACCENT = ("#0071E3", "#0A84FF")
_COLOR_ACCENT_HOVER = ("#0077ED", "#409CFF")
_COLOR_CANCEL = ("#636366", "#636366")
_COLOR_CANCEL_HOVER = ("#48484A", "#48484A")
_COLOR_ERROR_BG = ("#FEE2E2", "#3B0A0A")
_COLOR_ERROR_TEXT = ("#991B1B", "#FCA5A5")
_COLOR_PILL_READY_BG = ("#D1FAE5", "#052E16")
_COLOR_PILL_READY_TEXT = ("#065F46", "#6EE7B7")
_COLOR_PILL_ERROR_BG = ("#FEE2E2", "#3B0A0A")
_COLOR_PILL_ERROR_TEXT = ("#991B1B", "#FCA5A5")

# System font — renders correctly on both macOS (SF Pro) and Windows (Segoe UI)
_FONT_BODY = ("system-ui", 13)
_FONT_BODY_BOLD = ("system-ui", 13, "bold")
_FONT_SMALL = ("system-ui", 11)
_FONT_TITLE = ("system-ui", 15, "bold")

_PPI_OPTIONS = [72, 150, 300]
_PPI_LABELS = ["72 dpi", "150 dpi", "300 dpi"]

# ---------------------------------------------------------------------------
# Settings persistence (~/.pptx-exporter-settings.json)
# ---------------------------------------------------------------------------

_SETTINGS_PATH = Path.home() / ".pptx-exporter-settings.json"


def _load_settings() -> dict:
    try:
        with open(_SETTINGS_PATH) as fh:
            return json.load(fh)
    except Exception:
        return {}


def _save_settings(data: dict) -> None:
    try:
        with open(_SETTINGS_PATH, "w") as fh:
            json.dump(data, fh, indent=2)
    except Exception as exc:
        logger.debug("Could not save settings: %s", exc)


def _load_tkdnd(tkroot) -> bool:
    """Load the tkdnd extension, with a Tcl 9 compatibility fallback.

    tkinterdnd2 bundles a tkdnd binary compiled for Tcl 8.  On Tcl 9 the
    standard ``_require`` call fails.  This function pre-installs a
    Tcl-9-compatible binary from ``vendor/tkdnd/<platform>/`` (shipped with
    this repo) and patches the tkinterdnd2 package directory in-place *before*
    any load attempt, so only the correct binary is ever loaded.
    """
    if not _DND_AVAILABLE:
        return False

    import os
    import platform as _platform

    system = _platform.system()
    machine = _platform.machine()
    platform_map = {
        ("Darwin", "arm64"): "osx-arm64",
        ("Darwin", "x86_64"): "osx-x64",
        ("Linux", "aarch64"): "linux-arm64",
        ("Linux", "x86_64"): "linux-x64",
    }
    platform_dir = platform_map.get((system, machine))

    if platform_dir:
        # Check whether a vendored Tcl-9-compatible binary is available
        here = os.path.dirname(os.path.abspath(__file__))
        vendor_dir = os.path.normpath(
            os.path.join(here, "..", "..", "vendor", "tkdnd", platform_dir)
        )
        exts = (".dylib", ".so", ".dll")
        vendor_lib = next(
            (f for f in os.listdir(vendor_dir)
             if f.startswith("libtkdnd") and f.endswith(exts)),
            None,
        ) if os.path.isdir(vendor_dir) else None

        if vendor_lib:
            try:
                import shutil
                import tkinterdnd2

                mod_dir = os.path.join(
                    os.path.dirname(tkinterdnd2.__file__), "tkdnd", platform_dir
                )
                dest_lib = os.path.join(mod_dir, vendor_lib)

                # Install binary if not already present
                if not os.path.exists(dest_lib):
                    shutil.copy2(os.path.join(vendor_dir, vendor_lib), dest_lib)
                    logger.debug("Installed vendored %s", vendor_lib)

                # Patch pkgIndex.tcl to point at the new binary (once only)
                pkg_index = os.path.join(mod_dir, "pkgIndex.tcl")
                with open(pkg_index) as fh:
                    content = fh.read()
                version = (vendor_lib
                           .replace("libtkdnd", "")
                           .replace(".dylib", "")
                           .replace(".so", ""))
                old_lib = next(
                    (f for f in os.listdir(mod_dir)
                     if f.startswith("libtkdnd") and f.endswith(exts)
                     and f != vendor_lib),
                    None,
                )
                if old_lib and old_lib in content:
                    old_ver = (old_lib
                               .replace("libtkdnd", "")
                               .replace(".dylib", "")
                               .replace(".so", ""))
                    content = content.replace(f"{old_lib} tkdnd", f"{vendor_lib} Tkdnd")
                    content = content.replace(f"tkdnd {old_ver}", f"tkdnd {version}")
                    with open(pkg_index, "w") as fh:
                        fh.write(content)
                    logger.debug("Patched pkgIndex.tcl → tkdnd %s", version)

            except Exception as exc:
                logger.debug("tkdnd vendor setup failed: %s", exc)

    # Standard load — now works whether the binary was already compatible
    # or was just installed/patched above
    try:
        TkinterDnD._require(tkroot)
        return True
    except Exception as exc:
        logger.debug("tkdnd load failed: %s", exc)
        return False


class App(ctk.CTk):
    """Main application window."""

    def __init__(self) -> None:
        super().__init__()
        self._dnd_loaded = _load_tkdnd(self)

        configure_logging()

        self.title("pptx-exporter")
        self.resizable(True, False)
        self.minsize(520, 0)

        self._exporter = Exporter()
        self._pptx_path: Optional[str] = None
        self._powerpoint_available = self._exporter.backend != "not_found"
        self._cancel_event: Optional[threading.Event] = None

        settings = _load_settings()
        self._ppi: int = settings.get("ppi", 300)
        saved_out = settings.get("output_dir")
        self._output_dir: Optional[str] = saved_out if saved_out and os.path.isdir(saved_out) else None

        self._build_ui()
        self._update_run_button_state()
        # Defer DnD registration until the first event-loop tick so that the
        # native macOS NSWindow is fully realized before macdnd::registerdragwidget
        # is called.  Calling it during __init__ (before mainloop) means
        # _window_exists is still False and the C-level registration silently fails.
        if self._dnd_loaded:
            self.after(0, self._register_drop_target)

    # ------------------------------------------------------------------
    # UI construction
    # ------------------------------------------------------------------

    def _build_ui(self) -> None:
        PAD = 16

        self.grid_columnconfigure(0, weight=1)

        # ── Header ────────────────────────────────────────────────────
        header = ctk.CTkFrame(self, fg_color=_COLOR_BG_SURFACE, corner_radius=0)
        header.grid(row=0, column=0, sticky="ew", padx=0, pady=0)
        header.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            header,
            text="pptx-exporter",
            font=_FONT_TITLE,
            text_color=_COLOR_TEXT_PRIMARY,
            anchor="w",
        ).grid(row=0, column=0, sticky="w", padx=PAD, pady=(PAD, PAD))

        self._backend_pill = ctk.CTkLabel(
            header,
            text="",
            font=_FONT_SMALL,
            corner_radius=10,
            padx=10,
            pady=3,
        )
        self._backend_pill.grid(row=0, column=1, sticky="e", padx=PAD, pady=PAD)
        self._refresh_backend_pill()

        # Separator
        ctk.CTkFrame(self, height=1, fg_color=_COLOR_BORDER, corner_radius=0).grid(
            row=1, column=0, sticky="ew"
        )

        # ── Main content area ─────────────────────────────────────────
        content = ctk.CTkFrame(self, fg_color=_COLOR_BG_BASE, corner_radius=0)
        content.grid(row=2, column=0, sticky="ew")
        content.grid_columnconfigure(0, weight=1)

        # Input file section
        ctk.CTkLabel(
            content,
            text="INPUT FILE",
            font=_FONT_SMALL,
            text_color=_COLOR_TEXT_SECONDARY,
            anchor="w",
        ).grid(row=0, column=0, sticky="w", padx=PAD, pady=(PAD, 4))

        self._file_card = _FileCard(
            content, on_browse=self._browse_pptx, on_clear=self._clear_pptx,
            dnd_available=self._dnd_loaded,
        )
        self._file_card.grid(row=1, column=0, sticky="ew", padx=PAD, pady=(0, PAD))

        # Separator
        ctk.CTkFrame(content, height=1, fg_color=_COLOR_BORDER, corner_radius=0).grid(
            row=2, column=0, sticky="ew", padx=PAD
        )

        # Output folder section
        ctk.CTkLabel(
            content,
            text="OUTPUT FOLDER",
            font=_FONT_SMALL,
            text_color=_COLOR_TEXT_SECONDARY,
            anchor="w",
        ).grid(row=3, column=0, sticky="w", padx=PAD, pady=(PAD, 4))

        out_row = ctk.CTkFrame(content, fg_color="transparent")
        out_row.grid(row=4, column=0, sticky="ew", padx=PAD, pady=(0, PAD))
        out_row.grid_columnconfigure(0, weight=1)

        self._out_var = tk.StringVar(value=self._output_dir or "(none selected)")
        ctk.CTkLabel(
            out_row,
            textvariable=self._out_var,
            font=_FONT_BODY,
            text_color=_COLOR_TEXT_SECONDARY,
            anchor="w",
        ).grid(row=0, column=0, sticky="ew", padx=(0, 8))

        ctk.CTkButton(
            out_row,
            text="Browse…",
            command=self._browse_output,
            width=90,
            font=_FONT_BODY,
            fg_color=_COLOR_BG_SURFACE,
            text_color=_COLOR_TEXT_PRIMARY,
            border_width=1,
            border_color=_COLOR_BORDER,
            hover_color=_COLOR_BG_BASE,
        ).grid(row=0, column=1)

        # Separator
        ctk.CTkFrame(content, height=1, fg_color=_COLOR_BORDER, corner_radius=0).grid(
            row=5, column=0, sticky="ew", padx=PAD
        )

        # Resolution section
        ctk.CTkLabel(
            content,
            text="RESOLUTION",
            font=_FONT_SMALL,
            text_color=_COLOR_TEXT_SECONDARY,
            anchor="w",
        ).grid(row=6, column=0, sticky="w", padx=PAD, pady=(PAD, 4))

        initial_label = _PPI_LABELS[_PPI_OPTIONS.index(self._ppi)]
        self._ppi_seg = ctk.CTkSegmentedButton(
            content,
            values=_PPI_LABELS,
            command=self._on_ppi_change,
            font=_FONT_BODY,
        )
        self._ppi_seg.set(initial_label)
        self._ppi_seg.grid(row=7, column=0, sticky="w", padx=PAD, pady=(0, PAD))

        # Separator
        ctk.CTkFrame(self, height=1, fg_color=_COLOR_BORDER, corner_radius=0).grid(
            row=3, column=0, sticky="ew"
        )

        # ── Footer (progress + action) ────────────────────────────────
        footer = ctk.CTkFrame(self, fg_color=_COLOR_BG_SURFACE, corner_radius=0)
        footer.grid(row=4, column=0, sticky="ew")
        footer.grid_columnconfigure(0, weight=1)

        # Progress area (hidden until export starts)
        self._progress_frame = ctk.CTkFrame(footer, fg_color="transparent")
        self._progress_frame.grid(
            row=0, column=0, columnspan=2, sticky="ew", padx=PAD, pady=(PAD, 0)
        )
        self._progress_frame.grid_columnconfigure(0, weight=1)
        self._progress_frame.grid_remove()

        self._progress_bar = ctk.CTkProgressBar(
            self._progress_frame,
            mode="determinate",
            height=6,
        )
        self._progress_bar.set(0)
        self._progress_bar.grid(row=0, column=0, sticky="ew", pady=(0, 6))

        self._status_var = tk.StringVar(value="")
        self._status_label = ctk.CTkLabel(
            self._progress_frame,
            textvariable=self._status_var,
            font=_FONT_SMALL,
            text_color=_COLOR_TEXT_SECONDARY,
            anchor="w",
        )
        self._status_label.grid(row=1, column=0, sticky="w")

        # Open folder button (right of status, hidden until export done)
        self._open_folder_btn = ctk.CTkButton(
            self._progress_frame,
            text="Open Folder ↗",
            command=self._open_output_folder,
            width=110,
            height=26,
            font=_FONT_SMALL,
            fg_color="transparent",
            text_color=_COLOR_ACCENT,
            hover_color=_COLOR_BG_BASE,
            border_width=1,
            border_color=_COLOR_ACCENT,
            corner_radius=4,
        )
        self._open_folder_btn.grid(row=1, column=1, padx=(8, 0))
        self._open_folder_btn.grid_remove()

        # Action buttons row
        btn_row = ctk.CTkFrame(footer, fg_color="transparent")
        btn_row.grid(row=1, column=0, sticky="ew", padx=PAD, pady=PAD)
        btn_row.grid_columnconfigure(0, weight=1)

        self._run_btn = ctk.CTkButton(
            btn_row,
            text="Export PNGs",
            command=self._on_run,
            font=_FONT_BODY_BOLD,
            height=40,
            corner_radius=8,
            fg_color=_COLOR_ACCENT,
            hover_color=_COLOR_ACCENT_HOVER,
        )
        self._run_btn.grid(row=0, column=0, sticky="ew")

        self._cancel_btn = ctk.CTkButton(
            btn_row,
            text="Cancel",
            command=self._on_cancel,
            font=_FONT_BODY_BOLD,
            height=40,
            corner_radius=8,
            fg_color=_COLOR_CANCEL,
            hover_color=_COLOR_CANCEL_HOVER,
        )
        self._cancel_btn.grid(row=0, column=0, sticky="ew")
        self._cancel_btn.grid_remove()

        # Error banner (hidden until an error occurs)
        self._error_banner = _ErrorBanner(footer, on_dismiss=self._dismiss_error)
        self._error_banner.grid(row=2, column=0, sticky="ew", padx=PAD, pady=(0, PAD))
        self._error_banner.grid_remove()

        # Final geometry update
        self.update_idletasks()
        self.minsize(520, self.winfo_reqheight())

    # ------------------------------------------------------------------
    # Drag-and-drop
    # ------------------------------------------------------------------

    def _register_drop_target(self) -> None:
        """Register the whole window as a DnD drop target for files.

        ctk.CTk does not inherit from tkinter.BaseWidget, so we cannot use
        the DnDWrapper helper methods — call the Tcl tkdnd commands directly.

        tkdnd requires <<DropEnter>> and <<DropPosition>> handlers to return
        an action string ("copy") before it will ever fire <<Drop>>.
        """
        self.tk.call("tkdnd::drop_target", "register", self._w, (DND_FILES,))

        # tkdnd evaluates the binding script and substitutes ALL percent fields
        # (%A %a %b … %Y) before calling the registered command.  If the command
        # is registered with no arguments Tcl raises an error and the drop is
        # silently aborted.  Pass the full substitution string so every field is
        # forwarded as a positional argument; the lambdas accept *_ to discard them.
        _subst = "%A %a %b %C %c {%CST} {%CTT} %D %e {%L} {%m} {%ST} %T {%t} {%TT} %W %X %Y"

        accept_cb = self.register(lambda *_: "copy")
        drop_cb = self.register(self._on_drop_data)

        self.tk.call("bind", self._w, "<<DropEnter>>", f"{accept_cb} {_subst}")
        self.tk.call("bind", self._w, "<<DropPosition>>", f"{accept_cb} {_subst}")
        self.tk.call("bind", self._w, "<<Drop>>", f"{drop_cb} {_subst}")

    def _on_drop_data(self, _A, _a, _b, _C, _c, _CST, _CTT,
                      data: str, *_rest) -> None:
        """Handle a <<Drop>> event.  ``data`` is the %D substitution (file path)."""
        raw = data.strip()
        # tkinterdnd2 wraps paths with spaces in braces: {/path/to file.pptx}
        if raw.startswith("{") and raw.endswith("}"):
            raw = raw[1:-1]
        # If multiple files were dropped, take only the first
        path = raw.split("} {")[0].lstrip("{")
        if path.lower().endswith(".pptx") and os.path.isfile(path):
            self._set_pptx(path)
        else:
            logger.debug("Dropped file ignored (not a .pptx): %s", path)

    # ------------------------------------------------------------------
    # Backend pill
    # ------------------------------------------------------------------

    def _refresh_backend_pill(self) -> None:
        if self._powerpoint_available:
            self._backend_pill.configure(
                text="● PowerPoint ready",
                fg_color=_COLOR_PILL_READY_BG,
                text_color=_COLOR_PILL_READY_TEXT,
            )
        else:
            self._backend_pill.configure(
                text="● PowerPoint not found",
                fg_color=_COLOR_PILL_ERROR_BG,
                text_color=_COLOR_PILL_ERROR_TEXT,
            )

    # ------------------------------------------------------------------
    # Event handlers
    # ------------------------------------------------------------------

    def _browse_pptx(self) -> None:
        path = filedialog.askopenfilename(
            title="Select PowerPoint file",
            filetypes=[("PowerPoint files", "*.pptx"), ("All files", "*.*")],
        )
        if path:
            self._set_pptx(path)

    def _clear_pptx(self) -> None:
        self._pptx_path = None
        self._file_card.show_empty()
        self._update_run_button_state()

    def _set_pptx(self, path: str) -> None:
        self._pptx_path = path
        p = Path(path)
        # Set a smart output default: sibling folder named {stem}_pngs
        if self._output_dir is None:
            default_out = str(p.parent / f"{p.stem}_pngs")
            self._output_dir = default_out
            self._out_var.set(default_out)
        self._file_card.show_file(p, ppi=self._ppi)
        logger.debug("Selected pptx: %s", path)
        self._update_run_button_state()

    def _browse_output(self) -> None:
        path = filedialog.askdirectory(title="Select output folder")
        if path:
            self._output_dir = path
            self._out_var.set(path)
            logger.debug("Selected output dir: %s", path)
            _save_settings({"ppi": self._ppi, "output_dir": path})
            self._update_run_button_state()

    def _on_ppi_change(self, label: str) -> None:
        idx = _PPI_LABELS.index(label)
        self._ppi = _PPI_OPTIONS[idx]
        logger.debug("PPI set to %d", self._ppi)
        _save_settings({"ppi": self._ppi, "output_dir": self._output_dir})
        if self._pptx_path:
            self._file_card.show_file(Path(self._pptx_path), ppi=self._ppi)

    def _on_run(self) -> None:
        if not self._pptx_path or not self._output_dir:
            return
        self._error_banner.grid_remove()
        self._cancel_event = threading.Event()
        self._set_ui_busy(True)
        thread = threading.Thread(
            target=self._run_export, daemon=True
        )
        thread.start()

    def _on_cancel(self) -> None:
        if self._cancel_event:
            self._cancel_event.set()
        self._cancel_btn.configure(state="disabled", text="Cancelling…")
        self._status_var.set("Cancelling — finishing current slide…")

    def _open_output_folder(self) -> None:
        if not self._output_dir:
            return
        if sys.platform == "darwin":
            subprocess.run(["open", self._output_dir], check=False)
        elif sys.platform == "win32":
            os.startfile(self._output_dir)  # type: ignore[attr-defined]

    # ------------------------------------------------------------------
    # Export thread
    # ------------------------------------------------------------------

    def _run_export(self) -> None:
        try:
            self._exporter.export(
                self._pptx_path,
                self._output_dir,
                progress_callback=self._on_progress,
                cancel_event=self._cancel_event,
                ppi=self._ppi,
            )
            self.after(0, self._on_export_done)
        except InterruptedError:
            self.after(0, self._on_export_cancelled)
        except Exception as exc:
            logger.exception("Export failed")
            self.after(0, self._on_export_error, str(exc))

    def _on_progress(self, current: int, total: int) -> None:
        if total == 0:
            return
        fraction = current / total
        msg = (
            f"Processing slide {current + 1} of {total}…"
            if current < total
            else "Finalising…"
        )
        self.after(0, self._update_progress, fraction, msg)

    def _update_progress(self, fraction: float, msg: str) -> None:
        self._progress_bar.set(fraction)
        self._status_var.set(msg)

    def _on_export_done(self) -> None:
        self._progress_bar.set(1.0)
        self._status_var.set("Done — all slides exported.")
        self._set_ui_busy(False)
        self._open_folder_btn.grid()

    def _on_export_cancelled(self) -> None:
        self._progress_bar.set(0)
        self._status_var.set("Export cancelled.")
        self._set_ui_busy(False)

    def _on_export_error(self, message: str) -> None:
        self._set_ui_busy(False)
        self._error_banner.show(message)
        self._error_banner.grid()

    def _dismiss_error(self) -> None:
        self._error_banner.grid_remove()
        self._status_var.set("")

    # ------------------------------------------------------------------
    # Helpers
    # ------------------------------------------------------------------

    def _update_run_button_state(self) -> None:
        ready = bool(self._powerpoint_available and self._pptx_path and self._output_dir)
        self._run_btn.configure(state="normal" if ready else "disabled")

    def _set_ui_busy(self, busy: bool) -> None:
        if busy:
            self._progress_bar.set(0)
            self._status_var.set("Starting export…")
            self._open_folder_btn.grid_remove()
            self._progress_frame.grid()
            self._run_btn.grid_remove()
            self._cancel_btn.configure(state="normal", text="Cancel")
            self._cancel_btn.grid()
        else:
            self._cancel_btn.grid_remove()
            self._run_btn.grid()
            self._update_run_button_state()
        self.update_idletasks()


# ---------------------------------------------------------------------------
# Reusable sub-widgets
# ---------------------------------------------------------------------------

class _FileCard(ctk.CTkFrame):
    """Shows either a drop-prompt or a file summary card."""

    def __init__(self, parent, on_browse, on_clear, dnd_available: bool = False):
        super().__init__(
            parent,
            fg_color=_COLOR_BG_SURFACE,
            border_color=_COLOR_BORDER,
            border_width=1,
            corner_radius=8,
        )
        self.grid_columnconfigure(0, weight=1)
        self._on_browse = on_browse
        self._on_clear = on_clear
        self._dnd_available = dnd_available
        self._build_empty()
        self._empty_visible = True

    # ── Empty / drop-prompt state ──────────────────────────────────────

    def _build_empty(self) -> None:
        self._empty_frame = ctk.CTkFrame(self, fg_color="transparent")
        self._empty_frame.grid(row=0, column=0, sticky="ew")
        self._empty_frame.grid_columnconfigure(0, weight=1)

        headline = "Drop a .pptx file here" if self._dnd_available else "No file selected"
        subline = (
            "or click Browse to open a file" if self._dnd_available
            else "Click Browse to open a .pptx file"
        )

        ctk.CTkLabel(
            self._empty_frame,
            text=headline,
            font=_FONT_BODY_BOLD,
            text_color=_COLOR_TEXT_PRIMARY,
        ).grid(row=0, column=0, pady=(16, 2))

        ctk.CTkLabel(
            self._empty_frame,
            text=subline,
            font=_FONT_SMALL,
            text_color=_COLOR_TEXT_SECONDARY,
        ).grid(row=1, column=0, pady=(0, 12))

        ctk.CTkButton(
            self._empty_frame,
            text="Browse…",
            command=self._on_browse,
            width=110,
            font=_FONT_BODY,
            fg_color=_COLOR_ACCENT,
            hover_color=_COLOR_ACCENT_HOVER,
        ).grid(row=2, column=0, pady=(0, 16))

    def _build_file(self) -> None:
        self._file_frame = ctk.CTkFrame(self, fg_color="transparent")
        self._file_frame.grid(row=0, column=0, sticky="ew", padx=12, pady=10)
        self._file_frame.grid_columnconfigure(0, weight=1)

        name_row = ctk.CTkFrame(self._file_frame, fg_color="transparent")
        name_row.grid(row=0, column=0, sticky="ew")
        name_row.grid_columnconfigure(0, weight=1)

        self._filename_label = ctk.CTkLabel(
            name_row,
            text="",
            font=_FONT_BODY_BOLD,
            text_color=_COLOR_TEXT_PRIMARY,
            anchor="w",
        )
        self._filename_label.grid(row=0, column=0, sticky="w")

        ctk.CTkButton(
            name_row,
            text="✕",
            command=self._on_clear,
            width=24,
            height=24,
            font=_FONT_SMALL,
            fg_color="transparent",
            text_color=_COLOR_TEXT_SECONDARY,
            hover_color=_COLOR_BORDER,
            corner_radius=4,
        ).grid(row=0, column=1, padx=(8, 0))

        self._meta_label = ctk.CTkLabel(
            self._file_frame,
            text="",
            font=_FONT_SMALL,
            text_color=_COLOR_TEXT_SECONDARY,
            anchor="w",
        )
        self._meta_label.grid(row=1, column=0, sticky="w", pady=(2, 0))

    # ── Public API ─────────────────────────────────────────────────────

    def show_empty(self) -> None:
        if hasattr(self, "_file_frame"):
            self._file_frame.grid_remove()
        self._empty_frame.grid()
        self._empty_visible = True

    def show_file(self, path: Path, ppi: int = 300) -> None:
        self._empty_frame.grid_remove()
        if not hasattr(self, "_file_frame"):
            self._build_file()
        self._filename_label.configure(text=path.name)
        self._meta_label.configure(text=self._read_meta(path, ppi))
        self._file_frame.grid()
        self._empty_visible = False

    @staticmethod
    def _read_meta(path: Path, ppi: int = 300) -> str:
        try:
            from pptx import Presentation
            prs = Presentation(str(path))
            n = len(prs.slides)
            w_pt = prs.slide_width / 12700
            h_pt = prs.slide_height / 12700
            w_px = int(round(w_pt / 72 * ppi))
            h_px = int(round(h_pt / 72 * ppi))
            size_mb = path.stat().st_size / 1_048_576
            slides = f"{n} slide{'s' if n != 1 else ''}"
            return f"{slides}  ·  {size_mb:.1f} MB  ·  {w_px} × {h_px} px @ {ppi} dpi"
        except Exception:
            size_mb = path.stat().st_size / 1_048_576
            return f"{size_mb:.1f} MB"


class _ErrorBanner(ctk.CTkFrame):
    """Inline error banner — shown below the run button on failure."""

    def __init__(self, parent, on_dismiss):
        super().__init__(
            parent,
            fg_color=_COLOR_ERROR_BG,
            corner_radius=8,
        )
        self.grid_columnconfigure(0, weight=1)
        self._on_dismiss = on_dismiss

        self._msg_label = ctk.CTkLabel(
            self,
            text="",
            font=_FONT_SMALL,
            text_color=_COLOR_ERROR_TEXT,
            anchor="w",
            wraplength=400,
            justify="left",
        )
        self._msg_label.grid(row=0, column=0, sticky="ew", padx=12, pady=10)

        ctk.CTkButton(
            self,
            text="Dismiss",
            command=self._on_dismiss,
            width=70,
            height=26,
            font=_FONT_SMALL,
            fg_color="transparent",
            text_color=_COLOR_ERROR_TEXT,
            hover_color=_COLOR_ERROR_BG,
            border_width=1,
            border_color=_COLOR_ERROR_TEXT,
            corner_radius=4,
        ).grid(row=0, column=1, padx=(0, 12), pady=10)

    def show(self, message: str) -> None:
        self._msg_label.configure(text=message)
