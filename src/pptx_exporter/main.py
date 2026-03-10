"""Entry point — launches the pptx-exporter GUI."""

import sys


def main() -> None:
    """Start the Tkinter application."""
    from .gui import App

    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()
