"""Backward-compatible re-export of the refactored ExcelQueryEngine.

This module keeps the original import path working while the implementation
lives in the `eqe` package.
"""

from eqe import ExcelQueryEngine

__all__ = ["ExcelQueryEngine"]
