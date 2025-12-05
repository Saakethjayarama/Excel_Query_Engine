"""Example showing Excel-style range usage from a temporary .xlsx.

Demonstrates `get_range_by_ref`, which requires `openpyxl` for parsing Excel
references.
"""
import os
import tempfile
from ExcelQueryEngine import ExcelQueryEngine


def main():
    try:
        from openpyxl import Workbook
    except ModuleNotFoundError:
        print("openpyxl is required to run this example. Install with: pip install openpyxl")
        return

    wb = Workbook()
    ws = wb.active
    ws.title = 'Sheet1'
    ws.append(['Name', 'Age', 'City'])
    ws.append(['Alice', 30, 'NY'])
    ws.append(['Bob', 25, 'LA'])

    fd, path = tempfile.mkstemp(suffix='.xlsx')
    os.close(fd)
    try:
        wb.save(path)
        engine = ExcelQueryEngine(path)
        rng = engine.get_range_by_ref('Sheet1', 'A2:B3')
        print('Range A2:B3 ->', rng)
    finally:
        try:
            os.remove(path)
        except OSError:
            pass


if __name__ == '__main__':
    main()
