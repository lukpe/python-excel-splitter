import tkinter as tk
from functools import partial
from pathlib import Path
from tkinter import filedialog, messagebox

import openpyxl as xl
from openpyxl import Workbook


class App:
    def __init__(self):
        self.file_path = "None"
        self.wb_in = None
        self.ws_in = None

        self.root = tk.Tk()
        self.root.winfo_toplevel().title("Excel file splitter")
        self.canvas1 = tk.Canvas(self.root, width=400, height=220)
        self.canvas1.pack()

        self.button_file = tk.Button(text="Choose file", command=self.choose_file)
        self.canvas1.create_window(200, 30, window=self.button_file)

        self.label_file = tk.Label(text="File: " + self.file_path)
        self.canvas1.create_window(200, 60, window=self.label_file)

        self.canvas1.create_line(0, 90, 400, 90, fill="#000000")

        self.label_column = tk.Label(text="Column name:")
        self.canvas1.create_window(200, 110, window=self.label_column)

        self.variable = tk.StringVar(self.root)
        self.variable.set("None")
        self.list_column = tk.OptionMenu(self.root, self.variable, "None")
        self.canvas1.create_window(200, 140, window=self.list_column)

        self.canvas1.create_line(0, 170, 400, 170, fill="#000000")

        self.button_split = tk.Button(text="Split", command=self.split_workbook)
        self.canvas1.create_window(200, 200, window=self.button_split)

        self.root.mainloop()

    def choose_file(self):
        self.file_path = filedialog.askopenfilename()
        self.label_file["text"] = "File: " + self.file_path
        self.wb_in = xl.load_workbook(filename=self.file_path, data_only=True)
        self.ws_in = self.wb_in.worksheets[0]
        self.update_list(self.ws_in)

    def split_workbook(self):
        ext_position = self.file_path.find(".")
        file_name = self.file_path[0:ext_position]
        file_ext = self.file_path[ext_position:]
        col_name = self.variable.get()
        sheet_name = self.ws_in.title

        for row_in in range(2, self.ws_in.max_row + 1):
            col_num = self.get_column_number(col_name)
            value = self.ws_in.cell(column=col_num, row=row_in).value
            if value is None:
                value = "EmptyValue"
            file_out = Path(file_name + "_" + value + file_ext)
            wb_out = xl.load_workbook(self.create_workbook(file_out))
            ws_out = wb_out.worksheets[0]
            ws_out.title = sheet_name
            row_out = ws_out.max_row + 1
            for col_out in range(1, self.ws_in.max_column + 1):
                ws_out.cell(column=col_out, row=row_out).value = \
                    self.ws_in.cell(column=col_out, row=row_in).value
            wb_out.save(file_out)
            wb_out.close()
        self.wb_in.close()
        self.message_success()
        self.quit()

    def update_list(self, ws_in):
        options = []
        for col_in in range(1, ws_in.max_column + 1):
            options.append(ws_in.cell(column=col_in, row=1).value)
        list_columns = self.list_column["menu"]
        list_columns.delete(0, "end")
        for option in options:
            list_columns.add_command(label=option, command=partial(self.variable.set, option))
        self.variable.set(options[0])

    def get_column_number(self, col_name):
        for col_in in range(1, self.ws_in.max_column + 1):
            if col_name in self.ws_in.cell(column=col_in, row=1).value:
                return col_in

    def create_workbook(self, file):
        if not file.exists():
            wb_out = Workbook()
            ws_out = wb_out.worksheets[0]
            for col_in in range(1, self.ws_in.max_column + 1):
                ws_out.cell(column=col_in, row=1).value = self.ws_in.cell(column=col_in, row=1).value
            wb_out.save(file)
        return file

    @staticmethod
    def message_success():
        messagebox.showinfo(title="Success", message="File split correctly")

    def quit(self):
        self.root.destroy()


app = App()
