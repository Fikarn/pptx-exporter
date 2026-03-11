"""pptx-exporter: Export PowerPoint slide objects as transparent PNG images."""

from importlib.metadata import version, PackageNotFoundError

try:
    __version__ = version("pptx-exporter")
except PackageNotFoundError:
    __version__ = "0.0.0"
