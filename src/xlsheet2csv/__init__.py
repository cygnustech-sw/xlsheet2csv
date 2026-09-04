"""Deterministic XLSX worksheet extraction."""

from importlib.metadata import PackageNotFoundError, version

try:
    __version__ = version("xlsheet2csv")
except PackageNotFoundError:
    __version__ = "1.0.0"

__all__ = ["__version__"]
