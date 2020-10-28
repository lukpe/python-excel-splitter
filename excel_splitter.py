import getopt
import sys
from pathlib import Path

import openpyxl as xl
from openpyxl import Workbook


def main(argv):
    file_name = None
    file_ext = None
    col_name = None

    try:
        opts, args = getopt.getopt(argv, "i:c:", ["input_file=", "column_name="])
        for opt, arg in opts:
            if opt == "-i":
                ext_position = arg.find(".")
                file_name = arg[0:ext_position]
                file_ext = arg[ext_position:]
            elif opt == "-c":
                col_name = arg
    except getopt.GetoptError:
        print("excel_splitter.py -i <input_file> -c <column_name>")

    wb_in = xl.load_workbook(filename=file_name + file_ext, data_only=True)
    ws_in = wb_in.worksheets[0]

    for row_in in range(2, ws_in.max_row + 1):
        col_num = get_column_number(ws_in, col_name)
        value = ws_in.cell(column=col_num, row=row_in).value
        if value is None:
            value = "Empty"
        file_out = Path(file_name + "_" + value + file_ext)
        wb_out = xl.load_workbook(create_workbook(ws_in, file_out))
        ws_out = wb_out.worksheets[0]
        row_out = ws_out.max_row + 1
        for col_out in range(1, ws_in.max_column + 1):
            ws_out.cell(column=col_out, row=row_out).value = \
                ws_in.cell(column=col_out, row=row_in).value
        wb_out.save(file_out)
        wb_out.close()
    wb_in.close()


def get_column_number(ws_in, col_name):
    for col_in in range(1, ws_in.max_column + 1):
        if col_name in ws_in.cell(column=col_in, row=1).value:
            return col_in


def create_workbook(ws_in, file):
    if not file.exists():
        wb_out = Workbook()
        ws_out = wb_out.worksheets[0]
        for col_in in range(1, ws_in.max_column + 1):
            ws_out.cell(column=col_in, row=1).value = ws_in.cell(column=col_in, row=1).value
        wb_out.save(file)
    return file


if __name__ == "__main__":
    main(sys.argv[1:])
