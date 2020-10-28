import getopt
import sys
from pathlib import Path

import openpyxl as xl
from openpyxl import Workbook


def main(argv):
    file_name = None
    file_ext = None
    column_name = None

    try:
        opts, args = getopt.getopt(argv, "i:c:", ["input=", "column="])
        for opt, arg in opts:
            if opt == "-i":
                ext_position = arg.find(".")
                file_name = arg[0:ext_position]
                file_ext = arg[ext_position:]
            elif opt == "-c":
                column_name = arg
    except getopt.GetoptError:
        print("excel_splitter.py -i <input_file>")

    wb_in = xl.load_workbook(filename=file_name + file_ext, data_only=True)
    ws_in = wb_in.worksheets[0]

    column_number = get_column_number(ws_in, column_name)

    for r in range(2, ws_in.max_row + 1):
        value = ws_in.cell(column=column_number, row=r).value
        output_file = Path(file_name + "_" + value + file_ext)
        wb_out = xl.load_workbook(create_workbook(ws_in, output_file))
        ws_out = wb_out.worksheets[0]
        out_row = ws_out.max_row + 1
        for cc in range(1, ws_in.max_column + 1):
            ws_out.cell(column=cc, row=out_row).value = ws_in.cell(column=cc, row=r).value
        wb_out.save(output_file)
        wb_out.close()
    wb_in.close()


def get_column_number(ws_in, column_name):
    for c in range(1, ws_in.max_column + 1):
        if column_name in ws_in.cell(column=c, row=1).value:
            return c


def create_workbook(ws_in, file):
    if not file.exists():
        wb_out = Workbook()
        ws_out = wb_out.worksheets[0]
        for j in range(1, ws_in.max_column + 1):
            ws_out.cell(column=j, row=1).value = ws_in.cell(column=j, row=1).value
        wb_out.save(file)
    return file


if __name__ == "__main__":
    main(sys.argv[1:])
