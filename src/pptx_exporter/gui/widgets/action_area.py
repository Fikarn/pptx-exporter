"""ActionArea — export button, cancel, progress bar, status text, open folder."""

import tkinter as tk

import customtkinter as ctk

from ..tokens import COLORS, FONTS, RADIUS, SP


class ActionArea(ctk.CTkFrame):
    """Bottom action area: export/cancel button, progress, status, open folder."""

    def __init__(self, parent, on_run, on_cancel, on_open_folder):
        super().__init__(parent, fg_color="transparent")
        self.grid_columnconfigure(0, weight=1)
        self._on_run = on_run
        self._on_cancel = on_cancel
        self._on_open_folder = on_open_folder

        # Top divider
        ctk.CTkFrame(
            self, height=1, fg_color=COLORS["border"], corner_radius=0,
        ).grid(row=0, column=0, sticky="ew", pady=(0, SP["sm"]))

        # -- Hint text (shown when export is disabled) -------------------------
        self._hint_var = tk.StringVar(value="")
        self._hint_label = ctk.CTkLabel(
            self,
            textvariable=self._hint_var,
            font=FONTS["caption"],
            text_color=COLORS["text_tertiary"],
            anchor="center",
        )
        self._hint_label.grid(row=1, column=0, sticky="ew", pady=(0, SP["xs"]))
        self._hint_label.grid_remove()

        # -- Export button — tall, accent green --------------------------------
        self._run_btn = ctk.CTkButton(
            self,
            text="EXPORT PNGS",
            command=self._on_run,
            font=FONTS["body_bold"],
            height=48,
            corner_radius=RADIUS["sm"],
            fg_color=COLORS["accent"],
            hover_color=COLORS["accent_hover"],
            text_color=COLORS["on_accent"],
        )
        self._run_btn.grid(row=2, column=0, sticky="ew")

        # -- Cancel button (same grid slot, hidden) ----------------------------
        self._cancel_btn = ctk.CTkButton(
            self,
            text="Cancel",
            command=self._on_cancel,
            font=FONTS["body_bold"],
            height=48,
            corner_radius=RADIUS["sm"],
            fg_color=COLORS["cancel"],
            hover_color=COLORS["cancel_hover"],
            text_color=COLORS["text_primary"],
        )
        self._cancel_btn.grid(row=2, column=0, sticky="ew")
        self._cancel_btn.grid_remove()

        # -- Progress area (hidden until export starts) ------------------------
        self._progress_frame = ctk.CTkFrame(self, fg_color="transparent")
        self._progress_frame.grid(row=3, column=0, sticky="ew", pady=(SP["sm"], 0))
        self._progress_frame.grid_columnconfigure(0, weight=1)
        self._progress_frame.grid_remove()

        self._progress_bar = ctk.CTkProgressBar(
            self._progress_frame,
            mode="determinate",
            height=6,
            corner_radius=3,
            progress_color=COLORS["accent"],
            fg_color=COLORS["border"],
        )
        self._progress_bar.set(0)
        self._progress_bar.grid(row=0, column=0, sticky="ew", pady=(0, SP["xs"]))

        self._status_var = tk.StringVar(value="")
        self._status_label = ctk.CTkLabel(
            self._progress_frame,
            textvariable=self._status_var,
            font=FONTS["caption"],
            text_color=COLORS["text_secondary"],
            anchor="w",
        )
        self._status_label.grid(row=1, column=0, sticky="w")

        # Open folder button (right of status)
        self._open_btn = ctk.CTkButton(
            self._progress_frame,
            text="Open Folder \u2197",
            command=self._on_open_folder,
            width=110,
            height=26,
            font=FONTS["caption"],
            fg_color="transparent",
            text_color=COLORS["accent"],
            hover_color=COLORS["accent_muted"],
            border_width=1,
            border_color=COLORS["accent"],
            corner_radius=RADIUS["sm"],
        )
        self._open_btn.grid(row=1, column=1, padx=(SP["sm"], 0))
        self._open_btn.grid_remove()

    # -- Public API ------------------------------------------------------------

    def set_ready(self, ready: bool, hint: str = "") -> None:
        """Enable/disable the export button with optional hint text."""
        self._run_btn.configure(state="normal" if ready else "disabled")
        if hint and not ready:
            self._hint_var.set(hint)
            self._hint_label.grid()
        else:
            self._hint_label.grid_remove()

    def set_busy(self, busy: bool) -> None:
        """Switch between export/cancel modes."""
        if busy:
            self._progress_bar.set(0)
            self._status_var.set("Starting export\u2026")
            self._open_btn.grid_remove()
            self._progress_frame.grid()
            self._run_btn.grid_remove()
            self._hint_label.grid_remove()
            self._cancel_btn.configure(state="normal", text="Cancel")
            self._cancel_btn.grid()
        else:
            self._cancel_btn.grid_remove()
            self._run_btn.grid()

    def set_cancelling(self) -> None:
        self._cancel_btn.configure(state="disabled", text="Cancelling\u2026")
        self._status_var.set("Cancelling \u2014 finishing current slide\u2026")

    def update_progress(self, fraction: float, msg: str) -> None:
        self._progress_bar.set(fraction)
        self._status_var.set(msg)

    def show_done(self, msg: str) -> None:
        self._progress_bar.set(1.0)
        self._status_var.set(f"\u2713 {msg}")
        self._status_label.configure(text_color=COLORS["accent"])
        self._open_btn.grid()

    def show_cancelled(self) -> None:
        self._progress_bar.set(0)
        self._status_var.set("Export cancelled.")
        self._status_label.configure(text_color=COLORS["text_secondary"])

    def reset_progress(self) -> None:
        self._progress_frame.grid_remove()
        self._open_btn.grid_remove()
        self._status_var.set("")
        self._status_label.configure(text_color=COLORS["text_secondary"])
