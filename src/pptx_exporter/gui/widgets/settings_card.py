"""SettingsCard — resolution, slide selection, and output folder settings."""

import tkinter as tk

import customtkinter as ctk

from ..tokens import (
    COLORS, FONTS, RADIUS, SP,
    PPI_PRESETS, PPI_SEGMENT_VALUES, PPI_MIN, PPI_MAX,
)


class SettingsCard(ctk.CTkFrame):
    """Combined settings panel: resolution, slides, and output folder."""

    def __init__(
        self,
        parent,
        initial_ppi: int,
        initial_output: str | None,
        on_ppi_change,
        on_browse_output,
        on_slide_toggle,
    ):
        super().__init__(
            parent,
            fg_color=COLORS["surface"],
            border_color=COLORS["border"],
            border_width=1,
            corner_radius=RADIUS["md"],
        )
        self.grid_columnconfigure(0, weight=1)

        self._on_ppi_change_cb = on_ppi_change
        self._on_browse_output = on_browse_output
        self._on_slide_toggle_cb = on_slide_toggle

        row_idx = 0

        # -- RESOLUTION -------------------------------------------------------
        res_frame = ctk.CTkFrame(self, fg_color="transparent")
        res_frame.grid(
            row=row_idx, column=0, sticky="ew",
            padx=SP["md"], pady=(SP["md"], SP["sm"]),
        )
        res_frame.grid_columnconfigure(1, weight=1)
        row_idx += 1

        ctk.CTkLabel(
            res_frame,
            text="RESOLUTION",
            font=FONTS["label"],
            text_color=COLORS["text_secondary"],
            anchor="w",
        ).grid(row=0, column=0, sticky="w", padx=(0, SP["md"]))

        initial_label = next(
            (lbl for lbl, val in PPI_PRESETS.items() if val == initial_ppi),
            "Custom",
        )
        self._ppi_seg = ctk.CTkSegmentedButton(
            res_frame,
            values=PPI_SEGMENT_VALUES,
            command=self._on_ppi_seg_change,
            font=FONTS["caption"],
            selected_color=COLORS["accent"],
            selected_hover_color=COLORS["accent_hover"],
            unselected_color=COLORS["surface"],
            unselected_hover_color=COLORS["surface_hover"],
            text_color=COLORS["on_accent"],
        )
        self._ppi_seg.set(initial_label)
        self._ppi_seg.grid(row=0, column=1, sticky="e")

        # Custom PPI entry (inline, hidden by default)
        self._custom_frame = ctk.CTkFrame(self, fg_color="transparent")
        self._custom_frame.grid(
            row=row_idx, column=0, sticky="ew",
            padx=SP["md"], pady=(0, SP["sm"]),
        )
        row_idx += 1

        self._custom_entry = ctk.CTkEntry(
            self._custom_frame, width=70, font=FONTS["mono"],
            placeholder_text="dpi",
            fg_color=COLORS["bg"],
            border_color=COLORS["border"],
            text_color=COLORS["text_primary"],
        )
        self._custom_entry.grid(row=0, column=0)
        self._custom_entry.bind("<Return>", lambda _: self._apply_custom_ppi())
        self._custom_entry.bind("<FocusOut>", lambda _: self._apply_custom_ppi())

        self._custom_hint = ctk.CTkLabel(
            self._custom_frame,
            text=f"{PPI_MIN}\u2013{PPI_MAX}",
            font=FONTS["mono_sm"],
            text_color=COLORS["text_tertiary"],
        )
        self._custom_hint.grid(row=0, column=1, padx=(SP["sm"], 0))

        if initial_label == "Custom":
            self._custom_entry.insert(0, str(initial_ppi))
        else:
            self._custom_frame.grid_remove()

        # -- Divider -----------------------------------------------------------
        ctk.CTkFrame(
            self, height=1, fg_color=COLORS["border"], corner_radius=0,
        ).grid(row=row_idx, column=0, sticky="ew", padx=SP["md"], pady=SP["xs"])
        row_idx += 1

        # -- SLIDES (hidden until single file loaded) --------------------------
        self._slides_frame = ctk.CTkFrame(self, fg_color="transparent")
        self._slides_frame.grid(
            row=row_idx, column=0, sticky="ew",
            padx=SP["md"], pady=SP["md"],
        )
        self._slides_frame.grid_columnconfigure(1, weight=1)
        self._slides_row_idx = row_idx
        row_idx += 1

        ctk.CTkLabel(
            self._slides_frame,
            text="SLIDES",
            font=FONTS["label"],
            text_color=COLORS["text_secondary"],
            anchor="w",
        ).grid(row=0, column=0, sticky="w", padx=(0, SP["md"]))

        slide_controls = ctk.CTkFrame(self._slides_frame, fg_color="transparent")
        slide_controls.grid(row=0, column=1, sticky="e")

        self._all_slides_var = tk.BooleanVar(value=True)
        self._all_slides_cb = ctk.CTkCheckBox(
            slide_controls,
            text="All",
            variable=self._all_slides_var,
            command=self._on_slide_toggle,
            font=FONTS["caption"],
            checkbox_width=18,
            checkbox_height=18,
            fg_color=COLORS["accent"],
            hover_color=COLORS["accent_hover"],
            border_color=COLORS["border"],
            text_color=COLORS["text_primary"],
        )
        self._all_slides_cb.grid(row=0, column=0)

        self._slide_range_entry = ctk.CTkEntry(
            slide_controls,
            font=FONTS["mono_sm"],
            placeholder_text="e.g. 1-5, 8",
            width=120,
            height=28,
            fg_color=COLORS["bg"],
            border_color=COLORS["border"],
            text_color=COLORS["text_primary"],
        )
        self._slide_range_entry.grid(row=0, column=1, padx=(SP["sm"], 0))
        self._slide_range_entry.grid_remove()

        # Inline validation label for slide range
        self._slide_error = ctk.CTkLabel(
            self._slides_frame,
            text="",
            font=FONTS["caption"],
            text_color=COLORS["error"],
            anchor="e",
        )
        self._slide_error.grid(row=1, column=0, columnspan=2, sticky="e")
        self._slide_error.grid_remove()

        # Divider after slides
        self._slides_divider = ctk.CTkFrame(
            self, height=1, fg_color=COLORS["border"], corner_radius=0,
        )
        self._slides_divider.grid(row=row_idx, column=0, sticky="ew", padx=SP["md"], pady=SP["xs"])
        row_idx += 1

        # Hide slides section initially
        self._slides_frame.grid_remove()
        self._slides_divider.grid_remove()

        # -- OUTPUT ------------------------------------------------------------
        out_frame = ctk.CTkFrame(self, fg_color="transparent")
        out_frame.grid(
            row=row_idx, column=0, sticky="ew",
            padx=SP["md"], pady=SP["md"],
        )
        out_frame.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(
            out_frame,
            text="OUTPUT",
            font=FONTS["label"],
            text_color=COLORS["text_secondary"],
            anchor="w",
        ).grid(row=0, column=0, sticky="w", padx=(0, SP["md"]))

        self._out_var = tk.StringVar(value=initial_output or "Not set")
        self._out_label = ctk.CTkLabel(
            out_frame,
            textvariable=self._out_var,
            font=FONTS["caption"],
            text_color=COLORS["text_secondary"],
            anchor="e",
        )
        self._out_label.grid(row=0, column=1, sticky="ew", padx=(0, SP["sm"]))

        ctk.CTkButton(
            out_frame,
            text="Change\u2026",
            command=self._on_browse_output,
            width=80,
            height=28,
            font=FONTS["caption"],
            fg_color="transparent",
            text_color=COLORS["accent"],
            hover_color=COLORS["accent_muted"],
            border_width=1,
            border_color=COLORS["accent"],
            corner_radius=RADIUS["sm"],
        ).grid(row=0, column=2)

    # -- Resolution -----------------------------------------------------------

    def _on_ppi_seg_change(self, label: str) -> None:
        if label == "Custom":
            self._custom_frame.grid()
            self._custom_entry.delete(0, tk.END)
            self._custom_entry.focus_set()
            return
        self._custom_frame.grid_remove()
        self._on_ppi_change_cb(PPI_PRESETS[label])

    def _apply_custom_ppi(self) -> None:
        raw = self._custom_entry.get().strip()
        try:
            value = int(raw)
        except ValueError:
            return
        value = max(PPI_MIN, min(PPI_MAX, value))
        self._custom_entry.delete(0, tk.END)
        self._custom_entry.insert(0, str(value))
        self._on_ppi_change_cb(value)

    def set_custom_ppi_text(self, text: str) -> None:
        """Pre-fill the custom PPI entry."""
        self._custom_entry.delete(0, tk.END)
        self._custom_entry.insert(0, text)

    def get_ppi_seg_value(self) -> str:
        return self._ppi_seg.get()

    # -- Slides ---------------------------------------------------------------

    def show_slides(self) -> None:
        """Show the slides row (single file with >1 slide)."""
        self._slides_frame.grid()
        self._slides_divider.grid()
        self._all_slides_var.set(True)
        self._slide_range_entry.grid_remove()
        self._slide_error.grid_remove()

    def hide_slides(self) -> None:
        """Hide the slides row."""
        self._slides_frame.grid_remove()
        self._slides_divider.grid_remove()

    def _on_slide_toggle(self) -> None:
        if self._all_slides_var.get():
            self._slide_range_entry.grid_remove()
            self._slide_error.grid_remove()
        else:
            self._slide_range_entry.grid()
            self._slide_range_entry.focus_set()
        self._on_slide_toggle_cb()

    @property
    def all_slides(self) -> bool:
        return self._all_slides_var.get()

    def get_slide_range_text(self) -> str:
        return self._slide_range_entry.get().strip()

    def show_slide_error(self, msg: str) -> None:
        """Show inline validation error on the slide range."""
        self._slide_error.configure(text=msg)
        self._slide_error.grid()
        self._slide_range_entry.configure(border_color=COLORS["error"])

    def clear_slide_error(self) -> None:
        self._slide_error.grid_remove()
        self._slide_range_entry.configure(border_color=COLORS["border"])

    # -- Output ---------------------------------------------------------------

    def set_output_dir(self, path: str) -> None:
        display = path
        if len(display) > 45:
            display = "\u2026" + display[-42:]
        self._out_var.set(display)
