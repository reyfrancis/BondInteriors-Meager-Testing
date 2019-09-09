# CreateSheet.py

from Format import *
from Library import *

class CreateSheet(Format):

    def __init__(self, folder_name, dummy_name, book_name, sheet_name, cols, rows, *data_list):

        dirpath = os.getcwd()
        workbook_path = os.path.join(dirpath, dummy_name)

        self.folder_name = folder_name
        self.book_name = book_name
        self.sheet_name = sheet_name
        self.path = workbook_path

        self.wb = Workbook()
        self.ws = self.wb.active
        self.ws.title = self.sheet_name

        self.cols = cols
        self.rows = rows
        self.data_list = data_list[0]

        self.insert_cells()
        self.apply_format_default()
        self.create_sheet_format(self.data_list)

    def save_excel(self):
        dirpath = os.getcwd()
        self.wb.save(self.path) # Can be another folder?

    def __del__(self):
        self.save_excel()