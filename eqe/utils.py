"""Utility helpers for Excel reference parsing.

Importing `openpyxl` can be optional at runtime; the functions that need
`openpyxl` will import it lazily so the package can be imported even when
`openpyxl` is not installed. If you call an Excel-ref parsing function without
`openpyxl` installed, a clear error will be raised.
"""

from typing import Tuple


def _ensure_openpyxl_utils():
    try:
        from openpyxl.utils.cell import coordinate_from_string, column_index_from_string
        return coordinate_from_string, column_index_from_string
    except ModuleNotFoundError as exc:
        raise ModuleNotFoundError(
            "openpyxl is required for Excel-style reference parsing. "
            "Install it with 'pip install openpyxl' or avoid calling the "
            "Excel-ref helper functions." ) from exc


def excel_ref_to_index(cell_ref: str) -> Tuple[int, int]:
    """Convert an Excel cell reference like 'A1' to zero-based (row, col) index.

    Returns (row_index, col_index) both zero-based.
    """
    coordinate_from_string, column_index_from_string = _ensure_openpyxl_utils()
    col_letter, row = coordinate_from_string(cell_ref.upper())
    col = column_index_from_string(col_letter)
    return (row - 1, col - 1)


def parse_excel_range(range_ref: str):
    """Parse an Excel range string and return ((r1,c1), (r2,c2))."""
    if ":" in range_ref:
        start_ref, end_ref = range_ref.split(":")
        return (excel_ref_to_index(start_ref), excel_ref_to_index(end_ref))
    else:
        idx = excel_ref_to_index(range_ref)
        return (idx, idx)


def add_offset(base, offset):
    base_row, base_col = base
    offset_row, offset_col = offset
    return (base_row + offset_row, base_col + offset_col)


def column_letter_to_index(col_letter: str) -> int:
    """Convert a column letter like 'A' to zero-based column index."""
    _, column_index_from_string = _ensure_openpyxl_utils()
    return column_index_from_string(col_letter) - 1
