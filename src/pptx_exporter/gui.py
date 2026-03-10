"""All UI code for pptx-exporter — built with CustomTkinter."""

import logging
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog
from typing import Optional

import customtkinter as ctk

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
_COLOR_SUCCESS_BG = ("#D1FAE5", "#052E16")
_COLOR_SUCCESS_TEXT = ("#065F46", "#6EE7B7")
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


class App(ctk.CTk):
    """Main application window."""

    def __init__(self) -> None:
        super().__init__()
        configure_logging()

        self.title("pptx-exporter")
        self.resizable(True, False)
        self.minsize(520, 0)

        self._exporter = Exporter()
        self._pptx_path: Optional[str] = None
        self._output_dir: Optional[str] = None
        self._powerpoint_available = self._exporter.backend != "not_found"

        self._build_ui()
        self._update_run_button_state()

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
            content, on_browse=self._browse_pptx, on_clear=self._clear_pptx
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

        self._out_var = tk.StringVar(value="(none selected)")
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
        ctk.CTkFrame(self, height=1, fg_color=_COLOR_BORDER, corner_radius=0).grid(
            row=3, column=0, sticky="ew"
        )

        # ── Footer (progress + action) ────────────────────────────────
        footer = ctk.CTkFrame(self, fg_color=_COLOR_BG_SURFACE, corner_radius=0)
        footer.grid(row=4, column=0, sticky="ew")
        footer.grid_columnconfigure(0, weight=1)

        # Progress bar (hidden until export starts)
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
        ctk.CTkLabel(
            self._progress_frame,
            textvariable=self._status_var,
            font=_FONT_SMALL,
            text_color=_COLOR_TEXT_SECONDARY,
            anchor="w",
        ).grid(row=1, column=0, sticky="w")

        # Run button
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

        # Error banner (hidden until an error occurs)
        self._error_banner = _ErrorBanner(footer, on_dismiss=self._dismiss_error)
        self._error_banner.grid(row=2, column=0, sticky="ew", padx=PAD, pady=(0, PAD))
        self._error_banner.grid_remove()

        # Final geometry update
        self.update_idletasks()
        self.minsize(520, self.winfo_reqheight())

    # ------------------------------------------------------------------
    # Backend pill
    # ------------------------------------------------------------------

    def _refresh_backend_pill(self) -> None:
        if self._powerpoint_available:
            label = "● PowerPoint ready"
            self._backend_pill.configure(
                text=label,
                fg_color=_COLOR_PILL_READY_BG,
                text_color=_COLOR_PILL_READY_TEXT,
            )
        else:
            label = "● PowerPoint not found"
            self._backend_pill.configure(
                text=label,
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
        self._file_card.show_file(p)
        logger.debug("Selected pptx: %s", path)
        self._update_run_button_state()

    def _browse_output(self) -> None:
        path = filedialog.askdirectory(title="Select output folder")
        if path:
            self._output_dir = path
            self._out_var.set(path)
            logger.debug("Selected output dir: %s", path)
            self._update_run_button_state()

    def _on_run(self) -> None:
        if not self._pptx_path or not self._output_dir:
            return
        self._error_banner.grid_remove()
        self._set_ui_busy(True)
        thread = threading.Thread(target=self._run_export, daemon=True)
        thread.start()

    # ------------------------------------------------------------------
    # Export thread
    # ------------------------------------------------------------------

    def _run_export(self) -> None:
        try:
            self._exporter.export(
                self._pptx_path,
                self._output_dir,
                progress_callback=self._on_progress,
            )
            self.after(0, self._on_export_done)
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
        self._status_var.set(f"Done — PNGs saved to {self._output_dir}")
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
            self._progress_frame.grid()
            self._run_btn.configure(state="disabled", text="Exporting…")
        else:
            self._run_btn.configure(state="normal", text="Export PNGs")
        self.update_idletasks()


# ---------------------------------------------------------------------------
# Reusable sub-widgets
# ---------------------------------------------------------------------------

class _FileCard(ctk.CTkFrame):
    """Shows either a drop-prompt or a file summary card."""

    def __init__(self, parent, on_browse, on_clear):
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
        self._build_empty()
        self._empty_visible = True

    # ── Empty / drop-prompt state ──────────────────────────────────────

    def _build_empty(self) -> None:
        self._empty_frame = ctk.CTkFrame(self, fg_color="transparent")
        self._empty_frame.grid(row=0, column=0, sticky="ew")
        self._empty_frame.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            self._empty_frame,
            text="No file selected",
            font=_FONT_BODY_BOLD,
            text_color=_COLOR_TEXT_PRIMARY,
        ).grid(row=0, column=0, pady=(16, 2))

        ctk.CTkLabel(
            self._empty_frame,
            text="Click Browse to open a .pptx file",
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

    def show_file(self, path: Path) -> None:
        self._empty_frame.grid_remove()
        if not hasattr(self, "_file_frame"):
            self._build_file()
        self._filename_label.configure(text=path.name)
        self._meta_label.configure(text=self._read_meta(path))
        self._file_frame.grid()
        self._empty_visible = False

    @staticmethod
    def _read_meta(path: Path) -> str:
        try:
            from pptx import Presentation
            prs = Presentation(str(path))
            n = len(prs.slides)
            w_pt = prs.slide_width / 12700
            h_pt = prs.slide_height / 12700
            w_px = int(round(w_pt / 72 * 300))
            h_px = int(round(h_pt / 72 * 300))
            size_mb = path.stat().st_size / 1_048_576
            slides = f"{n} slide{'s' if n != 1 else ''}"
            return f"{slides}  ·  {size_mb:.1f} MB  ·  {w_px} × {h_px} px @ 300 dpi"
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
