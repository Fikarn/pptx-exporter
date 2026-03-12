"""Design tokens — colors, fonts, spacing constants for the GUI.

SSE brand visual system: SSE Green #99BA92, Dark Green #004932,
Beige #EDEBD1, Beige Light #F6F5E8. PT Sans (UI) + PT Serif (body).
Dual light/dark theme via (light, dark) tuples for CTk compatibility.
"""

import tkinter.font as tkfont

# ---------------------------------------------------------------------------
# Color palette — dual theme: (light, dark) tuples
# ---------------------------------------------------------------------------

COLORS = {
    # Backgrounds
    "bg":              ("#F6F5E8", "#003D2B"),   # Beige Light / deep dark green
    "surface":         ("#EDEBD1", "#004932"),   # Beige / Dark Green
    "surface_hover":   ("#E4E1C8", "#0A5C3F"),   # darker beige / lighter green
    "border":          ("#C9D4DA", "#1A6B4A"),   # Sky / muted green

    # Text
    "text_primary":    ("#000000", "#F6F5E8"),   # Black / Beige Light
    "text_secondary":  ("#3D3D3D", "#EDEBD1"),   # dark gray / Beige
    "text_tertiary":   ("#6B6B6B", "#99BA92"),   # medium gray / SSE Green muted

    # Accent — SSE Green in both modes
    "accent":          ("#99BA92", "#99BA92"),   # SSE Green (universal)
    "accent_hover":    ("#88AB81", "#AACAA3"),   # darker / lighter SSE Green
    "accent_muted":    ("#D6E5D2", "#1A4A35"),   # light green tint / dark green tint

    # Semantic
    "success":         ("#004932", "#99BA92"),   # Dark Green / SSE Green
    "success_bg":      ("#D6E5D2", "#0A3D2B"),
    "success_text":    ("#004932", "#99BA92"),
    "error":           ("#671919", "#FF7D55"),   # Burgundy / Coral
    "error_bg":        ("#F5E0E0", "#3A1A15"),
    "error_text":      ("#671919", "#FF7D55"),
    "cancel":          ("#9E9E9E", "#6B6B5C"),
    "cancel_hover":    ("#858585", "#525245"),

    # Button text (text on accent-colored buttons)
    "on_accent":       ("#000000", "#004932"),   # Black / Dark Green
}

# ---------------------------------------------------------------------------
# Typography
# ---------------------------------------------------------------------------


def _resolve_fonts():
    """Called once after Tk root exists. Returns resolved font families."""
    try:
        available = set(tkfont.families())
    except Exception:
        available = set()

    def pick(primary, *fallbacks):
        if primary in available:
            return primary
        for f in fallbacks:
            if f in available:
                return f
        return fallbacks[-1]

    sans = pick("PT Sans", "Helvetica Neue", "Segoe UI", "Arial")
    serif = pick("PT Serif", "Georgia", "Times New Roman")
    mono = pick("SF Mono", "Consolas", "Courier New")
    return sans, serif, mono


def _build_fonts(sans, serif, mono):
    """Build the FONTS dict from resolved family names."""
    return {
        "display":      (sans, 20, "bold"),
        "title":        (sans, 15, "bold"),
        "heading":      (sans, 13, "bold"),
        "body":         (serif, 13),
        "body_bold":    (serif, 13, "bold"),
        "body_italic":  (serif, 13, "italic"),
        "caption":      (sans, 11),
        "caption_bold": (sans, 11, "bold"),
        "mono":         (mono, 12),
        "mono_sm":      (mono, 11),
        "label":        (sans, 11, "bold"),
    }


# Placeholder — populated by init_fonts() after Tk root
_SANS = "PT Sans"
_SERIF = "PT Serif"
_MONO = "SF Mono"

FONTS = _build_fonts(_SANS, _SERIF, _MONO)


def init_fonts():
    """Resolve available fonts — call from App.__init__ after CTk root exists."""
    global _SANS, _SERIF, _MONO, FONTS
    _SANS, _SERIF, _MONO = _resolve_fonts()
    FONTS = _build_fonts(_SANS, _SERIF, _MONO)


# ---------------------------------------------------------------------------
# Spacing (4 px grid)
# ---------------------------------------------------------------------------

SP = {
    "xs":  4,
    "sm":  8,
    "md":  16,
    "lg":  24,
    "xl":  32,
    "2xl": 48,
}

# ---------------------------------------------------------------------------
# Radii
# ---------------------------------------------------------------------------

RADIUS = {
    "sm": 4,
    "md": 8,
    "lg": 12,
}

# ---------------------------------------------------------------------------
# PPI presets
# ---------------------------------------------------------------------------

PPI_PRESETS = {
    "72":  72,
    "150": 150,
    "300": 300,
}
PPI_SEGMENT_VALUES = ["72", "150", "300", "Custom"]
PPI_MIN = 36
PPI_MAX = 2400
