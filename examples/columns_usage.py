"""Example showing how to select specific columns by letter from a file.

Creates a temporary .xlsx and demonstrates `get_columns_from_row` which uses
Excel column letters (requires `openpyxl`).
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
        cols = engine.get_columns_from_row('Sheet1', ['A', 'C'], start_row=2)
        print('Selected columns A and C from row 2:', cols)
    finally:
        try:
            os.remove(path)
        except OSError:
            pass


if __name__ == '__main__':
    main()
