"""Design tokens — colors, fonts, spacing constants for the GUI."""

# Color palette — tuples of (light, dark) for CustomTkinter
COLORS = {
    "bg":              ("#F2F2F7", "#1C1C1E"),
    "surface":         ("#FFFFFF", "#2C2C2E"),
    "surface_hover":   ("#F0F0F5", "#3A3A3C"),
    "border":          ("#D1D1D6", "#48484A"),
    "text_primary":    ("#1D1D1F", "#F5F5F7"),
    "text_secondary":  ("#86868B", "#8E8E93"),
    "text_tertiary":   ("#AEAEB2", "#636366"),
    "accent":          ("#0071E3", "#0A84FF"),
    "accent_hover":    ("#0077ED", "#409CFF"),
    "success":         ("#34C759", "#30D158"),
    "success_bg":      ("#D1FAE5", "#052E16"),
    "success_text":    ("#065F46", "#6EE7B7"),
    "error":           ("#FF3B30", "#FF453A"),
    "error_bg":        ("#FFF0F0", "#3B1010"),
    "error_text":      ("#991B1B", "#FCA5A5"),
    "cancel":          ("#636366", "#636366"),
    "cancel_hover":    ("#48484A", "#48484A"),
}

# Typography — system-ui maps to SF Pro on macOS, Segoe UI on Windows
FONTS = {
    "title":     ("system-ui", 18, "bold"),
    "heading":   ("system-ui", 14, "bold"),
    "body":      ("system-ui", 13),
    "body_bold": ("system-ui", 13, "bold"),
    "caption":   ("system-ui", 11),
}

# Spacing (4px grid)
SP = {
    "xs": 4,
    "sm": 8,
    "md": 16,
    "lg": 24,
    "xl": 32,
}

# PPI presets for the resolution picker
PPI_PRESETS = {
    "72": 72,
    "150": 150,
    "300": 300,
}
PPI_SEGMENT_VALUES = ["72", "150", "300", "Custom"]
PPI_MIN = 36
PPI_MAX = 2400
