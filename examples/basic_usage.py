"""Basic usage example for `ExcelQueryEngine` using an on-disk .xlsx file.

This example creates a temporary workbook (using `openpyxl`), saves it to a
temporary file and passes the filename to `ExcelQueryEngine`.
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

        print('Cell (1,0):', engine.get_cell('Sheet1', 1, 0))
        print('Find Bob:', engine.find_by_value('Sheet1', 'Bob'))
        print('Adjacent to Alice (right):', engine.get_adjacent_value('Sheet1', 'Alice', offset=(0, 1)))
        print('Extracted table:', engine.extract_table_from_header('Sheet1', 0))
    finally:
        try:
            os.remove(path)
        except OSError:
            pass


if __name__ == '__main__':
    main()
