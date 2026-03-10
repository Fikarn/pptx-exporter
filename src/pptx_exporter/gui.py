"""All Tkinter UI code for pptx-exporter."""

import logging
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from typing import Optional

from .exporter import Exporter
from .utils import configure_logging

logger = logging.getLogger(__name__)


class App(tk.Tk):
    """Main application window."""

    def __init__(self) -> None:
        super().__init__()
        configure_logging()

        self.title("pptx-exporter")
        self.resizable(False, False)
        self._exporter = Exporter()

        self._pptx_path: Optional[str] = None
        self._output_dir: Optional[str] = None

        self._build_ui()
        self._update_run_button_state()

    # ------------------------------------------------------------------
    # UI construction
    # ------------------------------------------------------------------

    def _build_ui(self) -> None:
        PAD = 12

        # ── Backend status banner ──────────────────────────────────────
        banner_frame = tk.Frame(self, bg="#f0f4ff", pady=6)
        banner_frame.pack(fill=tk.X, padx=PAD, pady=(PAD, 0))

        tk.Label(
            banner_frame,
            text="Mode:",
            font=("Helvetica", 10, "bold"),
            bg="#f0f4ff",
        ).pack(side=tk.LEFT, padx=(6, 4))

        tk.Label(
            banner_frame,
            text=self._exporter.backend_label,
            fg="#2255aa",
            bg="#f0f4ff",
            font=("Helvetica", 10),
            wraplength=380,
            justify=tk.LEFT,
        ).pack(side=tk.LEFT, padx=(0, 6))

        # ── Input file row ─────────────────────────────────────────────
        file_frame = tk.Frame(self)
        file_frame.pack(fill=tk.X, padx=PAD, pady=(PAD, 0))

        tk.Label(file_frame, text="Input .pptx:", width=12, anchor="w").pack(
            side=tk.LEFT
        )

        self._pptx_var = tk.StringVar(value="(none selected)")
        tk.Label(
            file_frame,
            textvariable=self._pptx_var,
            fg="#555",
            anchor="w",
            wraplength=260,
            justify=tk.LEFT,
        ).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(4, 8))

        tk.Button(
            file_frame,
            text="Browse…",
            command=self._browse_pptx,
            width=9,
        ).pack(side=tk.RIGHT)

        # ── Output folder row ──────────────────────────────────────────
        out_frame = tk.Frame(self)
        out_frame.pack(fill=tk.X, padx=PAD, pady=(6, 0))

        tk.Label(out_frame, text="Output folder:", width=12, anchor="w").pack(
            side=tk.LEFT
        )

        self._out_var = tk.StringVar(value="(none selected)")
        tk.Label(
            out_frame,
            textvariable=self._out_var,
            fg="#555",
            anchor="w",
            wraplength=260,
            justify=tk.LEFT,
        ).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(4, 8))

        tk.Button(
            out_frame,
            text="Browse…",
            command=self._browse_output,
            width=9,
        ).pack(side=tk.RIGHT)

        # ── Progress bar ───────────────────────────────────────────────
        progress_frame = tk.Frame(self)
        progress_frame.pack(fill=tk.X, padx=PAD, pady=(PAD, 0))

        self._progress = ttk.Progressbar(
            progress_frame, orient=tk.HORIZONTAL, mode="determinate", length=420
        )
        self._progress.pack(fill=tk.X)

        # ── Status label ───────────────────────────────────────────────
        self._status_var = tk.StringVar(value="Ready.")
        tk.Label(
            self,
            textvariable=self._status_var,
            fg="#333",
            font=("Helvetica", 10),
            wraplength=420,
            justify=tk.LEFT,
            anchor="w",
        ).pack(fill=tk.X, padx=PAD, pady=(4, 0))

        # ── Run button ─────────────────────────────────────────────────
        btn_frame = tk.Frame(self)
        btn_frame.pack(fill=tk.X, padx=PAD, pady=(10, PAD))

        self._run_btn = tk.Button(
            btn_frame,
            text="Run Export",
            command=self._on_run,
            width=14,
            font=("Helvetica", 11, "bold"),
            bg="#2255aa",
            fg="white",
            activebackground="#1a4090",
            activeforeground="white",
            relief=tk.FLAT,
            padx=8,
            pady=6,
        )
        self._run_btn.pack(side=tk.RIGHT)

        # Minimum window width
        self.update_idletasks()
        self.minsize(460, self.winfo_reqheight())

    # ------------------------------------------------------------------
    # Event handlers
    # ------------------------------------------------------------------

    def _browse_pptx(self) -> None:
        path = filedialog.askopenfilename(
            title="Select PowerPoint file",
            filetypes=[("PowerPoint files", "*.pptx"), ("All files", "*.*")],
        )
        if path:
            self._pptx_path = path
            self._pptx_var.set(Path(path).name)
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
            messagebox.showerror("Missing input", "Please select both a .pptx file and an output folder.")
            return

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
        pct = int(current / total * 100)
        msg = (
            f"Processing slide {current + 1} of {total}…"
            if current < total
            else "Finalising…"
        )
        self.after(0, self._update_progress, pct, msg)

    def _update_progress(self, pct: int, msg: str) -> None:
        self._progress["value"] = pct
        self._status_var.set(msg)

    def _on_export_done(self) -> None:
        self._progress["value"] = 100
        self._status_var.set(f"Done! PNGs saved to: {self._output_dir}")
        self._set_ui_busy(False)
        messagebox.showinfo("Export complete", f"All slides exported to:\n{self._output_dir}")

    def _on_export_error(self, message: str) -> None:
        self._status_var.set(f"Error: {message}")
        self._set_ui_busy(False)
        messagebox.showerror("Export failed", message)

    # ------------------------------------------------------------------
    # Helpers
    # ------------------------------------------------------------------

    def _update_run_button_state(self) -> None:
        ready = bool(self._pptx_path and self._output_dir)
        self._run_btn.config(state=tk.NORMAL if ready else tk.DISABLED)

    def _set_ui_busy(self, busy: bool) -> None:
        state = tk.DISABLED if busy else tk.NORMAL
        self._run_btn.config(state=state)
        if busy:
            self._progress["value"] = 0
            self._status_var.set("Starting export…")
        self.update_idletasks()
