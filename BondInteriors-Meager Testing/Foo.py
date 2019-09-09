# Foo.py

from CopyFoo import *
from Library import *

class Foo:

    def __init__(self, folder_name):

        self.folder_name = folder_name
        self.smdb_list = []
        self.db_list = []
        self.workbook_path = []
        dirpath = os.getcwd()
        for each_file in os.listdir(dirpath):
            ext = os.path.splitext(each_file)[-1].lower()
            if ext == ".xlsx":
                file_path = each_file
                self.workbook_path.append(os.path.join(dirpath, file_path)) 

        for each_book_path in self.workbook_path:
            self.wb = load_workbook(each_book_path)

            for each_sheet in self.wb.sheetnames:
                if "smdb" in each_sheet.lower():                
                    self.smdb_list.append(each_sheet)
                    self.smdb_list.append(each_book_path) # The signature path

                elif "db" in each_sheet.lower() and "smdb" not in each_sheet.lower():
                    self.db_list.append(each_sheet)
                    self.db_list.append(each_book_path) # The signature path

        i = 0
        while i < len(self.smdb_list):
            book_name = os.path.basename(self.smdb_list[i + 1])
            CopyFoo(self.folder_name, self.smdb_list[i], book_name, self.smdb_list[i + 1], True)
            i += 2

        i = 0
        while i < len(self.db_list):
            book_name = os.path.basename(self.db_list[i + 1])
            CopyFoo(self.folder_name, self.db_list[i], book_name, self.db_list[i + 1], False) # folder_name, sheet_name, book_name, path, bool_sm
            i += 2