# CopyFoo.py

from CreateSheet import *
from Library import *

class CopyFoo(CreateSheet):

    def __init__(self, folder_name, sheet_name, book_name, path, bool_sm): # argument is boolean. True refers to SMDB type. False refers to DB type.

        self.folder_name = folder_name
        self.sheet_name = sheet_name
        self.book_name = book_name
        self.path = path
        self.bool_sm = bool_sm
        self.cols = 7

        if self.bool_sm == True:
            self.rows = 21
            self.circuit_ref_list = []
            self.circuit_designation_list = []
            self.cable_list = []
            self.ecc_list = []
            self.numcore_list = []
            self.mccbamp_list = []
            self.mcbtype_list = []
            self.mcbka_list = []
            self.rccbma_list = []
            self.ry_list = []
            self.yb_list = []
            self.br_list = []
            self.rybn_list = []
            self.rybe_list = []
            self.ne_list = []
            self.continuitytest_list = []
            self.ring_list = []
            self.resistance_list = []
            self.remarks_list = []

            self.paste_list =  [self.circuit_ref_list,
                                self.circuit_designation_list,
                                self.cable_list,
                                self.ecc_list,
                                self.numcore_list,
                                self.mccbamp_list,
                                self.mcbtype_list,
                                self.mcbka_list,
                                self.rccbma_list,
                                self.ry_list,
                                self.yb_list,
                                self.br_list,
                                self.rybn_list,
                                self.rybe_list,
                                self.ne_list,
                                self.continuitytest_list,
                                self.ring_list,
                                self.resistance_list,
                                self.remarks_list]
            
            self.wb = load_workbook(self.path, data_only=True)
            self.ws = self.wb[self.sheet_name]
            self.count = self.get_count(self.bool_sm)
            self.copy_val_smdb()

            smdb_Sheet = CreateSheet(self.folder_name, 'dummy.xlsx', self.book_name, self.sheet_name, self.cols, self.rows,
                        [data_list.SMDB_header_list,
                        data_list.SMDB_merge_x_list,
                        data_list.SMDB_merge_y_list,
                        data_list.SMDB_merge_xy_list,
                        data_list.SMDB_format_list,
                        data_list.SMDB_border_list])
            del smdb_Sheet

        else:
            self.rows = 25
            self.circuit_ref_list = []
            self.circuit_designation_list = []
            self.cable_list = []
            self.cpc_list = []
            self.numcore_list = []
            self.mccbamp_list = []
            self.mcbtype_list = []
            self.mcbka_list = []
            self.elcb_rating_list = []
            self.ry_list = []
            self.yb_list = []
            self.br_list = []
            self.rn_list = []
            self.yn_list = []
            self.bn_list = []
            self.re_list = []
            self.ye_list = []
            self.be_list = []
            self.ne_list = []
            self.continuitytest_list = []
            self.ring_list = []
            self.resistance_list = []
            self.remarks_list = []

            self.paste_list =  [self.circuit_ref_list,
                                self.circuit_designation_list,
                                self.cable_list,
                                self.cpc_list,
                                self.numcore_list,
                                self.mccbamp_list,
                                self.mcbtype_list,
                                self.mcbka_list,
                                self.elcb_rating_list,
                                self.ry_list,
                                self.yb_list,
                                self.br_list,
                                self.rn_list,
                                self.yn_list,
                                self.bn_list,
                                self.re_list,
                                self.ye_list,
                                self.be_list,
                                self.ne_list,
                                self.continuitytest_list,
                                self.ring_list,
                                self.resistance_list,
                                self.remarks_list]

            self.wb = load_workbook(self.path, data_only=True)
            self.ws = self.wb[self.sheet_name]
            self.count = self.get_count(self.bool_sm)
            self.copy_val_db()

            db_Sheet = CreateSheet(self.folder_name, 'dummy.xlsx', self.book_name, self.sheet_name, self.cols, self.rows, 
                        [data_list.DB_header_list,
                        data_list.DB_merge_x_list,
                        data_list.DB_merge_y_list,
                        data_list.DB_merge_xy_list,
                        data_list.DB_format_list,
                        data_list.DB_border_list])
            del db_Sheet

    def __del__(self):

        dirpath = os.getcwd()
        workbook_path = os.path.join(dirpath, 'dummy.xlsx')
        self.wb = load_workbook(workbook_path)
        self.ws = self.wb[self.sheet_name]
        self.paste_val()
        self.wb.save(workbook_path)
        self.copy_sheets_to_file()
        os.remove(self.dummy_path) # Delete the dummy file
        self.final_layout()

    def final_layout(self):
        self.path = self.final_path # Update the path, so we can use the save_excel() method directly

        self.wb = load_workbook(self.final_path)
        self.ws = self.wb[self.sheet_name]


        if self.bool_sm == True:
            # Can't move to data_list.py since not all are raw value
            self.apply_format_specific('general', 'center', 'Arial', 10, 
                            'FF000000', False, False, False, [1, 17], [1, 16 + self.count])
            self.apply_format_specific('left', 'bottom', 'Arial', 10, 
                            'FF000000', True, True, False, [2, 17], [2, 16 + self.count])
            self.apply_format_specific('center', 'center', 'Times New Roman', 10, 
                    'FF000000', False, False, False, [3, 17], [3, 16 + self.count])
            self.apply_format_specific('center', 'bottom', 'Times New Roman', 10, 
            'FF000000', False, False, False, [4, 17], [4, 16 + self.count])
            self.apply_format_specific('center', 'center', 'Calibri', 11, 
            'FF000000', False, False, False, [5, 17], [5, 16 + self.count])
            self.apply_format_specific('center', 'bottom', 'Times New Roman', 10, 
                    'FF000000', True, True, False, [6, 17], [6, 16 + self.count])
            self.apply_format_specific('center', 'bottom', 'Times New Roman', 10, 
            'FF000000', True, True, False, [8, 17], [8, 16 + self.count])
            self.apply_format_specific('center', 'bottom', 'Arial', 10, 
            'FF000000', False, False, False, [10, 17], [10, 16 + self.count])
            self.apply_format_specific('center', 'bottom', 'Arial', 10, 
            'FF000000', False, False, False, [11, 17], [11, 16 + self.count])
            self.apply_format_specific('center', 'bottom', 'Arial', 10, 
            'FF000000', False, False, False, [12, 17], [12, 16 + self.count])
            self.apply_format_specific('center', 'bottom', 'Arial', 10, 
            'FF000000', False, False, False, [13, 17], [13, 16 + self.count])
            self.apply_format_specific('center', 'bottom', 'Arial', 10, 
            'FF000000', False, False, False, [14, 17], [14, 16 + self.count])
            self.apply_format_specific('center', 'bottom', 'Arial', 10, 
            'FF000000', False, False, False, [15, 17], [15, 16 + self.count])
            self.apply_format_specific('center', 'bottom', 'Arial', 10, 
            'FF000000', False, False, False, [16, 17], [16, 16 + self.count])
            self.apply_format_specific('center', 'bottom', 'Arial', 10, 
            'FF000000', False, False, False, [17, 17], [17, 16 + self.count])
            self.apply_format_specific('center', 'bottom', 'Arial', 10, 
            'FF000000', False, False, False, [18, 17], [18, 16 + self.count])

            self.merge_cell_x_iterator(self.count, [self.rows - 2, 17], [self.rows, 17])
            self.merge_cell_xy([1, 17 + self.count], [5, 18 + self.count])
            self.merge_cell_xy([1, 19 + self.count], [5, 19 + self.count])
            self.merge_cell_x_iterator(2, [6, 17 + self.count], [self.rows, 17 + self.count])
            self.merge_cell_x([6, 19 + self.count], [11, 19 + self.count])
            self.merge_cell_x([12, 19 + self.count], [16, 19 + self.count])  
            self.merge_cell_x([17, 19 + self.count], [self.rows, 19 + self.count])
            self.merge_cell_x([1, 20 + self.count], [self.rows, 20 + self.count])

            self.write_cell('NOTES: TEST READING ARE WITHIN THE LIMIT', [1, 17 + self.count])
            self.write_cell('ACCEPTED BY:', [1, 19 + self.count])
            self.write_cell('COLD TESTED BY:', [6, 17 + self.count])
            self.write_cell('WITNESSED BY:', [6, 18 + self.count])
            self.write_cell('TEST INSTRUMENT:', [6, 19 + self.count])
            self.write_cell('SERIAL NUMBER:', [12, 19 + self.count])
            self.write_cell('CALIBRATION DATE: ', [17, 19 + self.count])        

            self.apply_border_default([1, self.cols], [self.rows, 20 + self.count])

            self.apply_format_specific('left', 'center', 'Arial', 11, 
            'FF000000', True, False, True, [1, 17 + self.count], [self.rows, 20 + self.count])

        else: 

            # Can't move to data_list.py since not all are raw value
            self.apply_format_specific('center', 'center', 'Arial', 10, 
                            'FF000000', False, False, False, [1, 17], [1, 16 + self.count])
            self.apply_format_specific('center', 'center', 'Arial', 10, 
                            'FF000000', False, False, False, [2, 17], [2, 16 + self.count])
            self.apply_format_specific('center', 'center', 'Arial', 10, 
                    'FF000000', False, False, False, [3, 17], [3, 16 + self.count])
            self.apply_format_specific('center', 'center', 'Arial', 10, 
            'FF000000', False, False, False, [4, 17], [4, 16 + self.count])
            self.apply_format_specific('center', 'center', 'Calibri', 11, 
            'FF000000', False, False, False, [5, 17], [5, 16 + self.count])
            self.apply_format_specific('center', 'center', 'Arial', 10, 
                    'FF000000', False, False, False, [6, 17], [6, 16 + self.count])
            self.apply_format_specific('center', 'center', 'Calibri', 11, 
            'FF000000', False, False, False, [7, 17], [7, 16 + self.count])
            self.apply_format_specific('center', 'center', 'Calibri', 11, 
            'FF000000', False, False, False, [8, 17], [8, 16 + self.count])
            self.apply_format_specific('center', 'center', 'Arial', 10, 
            'FF000000', False, False, True, [9, 17], [9, 16 + self.count])
            self.apply_format_specific('center', 'center', 'Calibri', 11, 
            'FF000000', False, False, False, [10, 17], [10, 16 + self.count])
            self.apply_format_specific('center', 'center', 'Calibri', 11, 
            'FF000000', False, False, False, [11, 17], [11, 16 + self.count])
            self.apply_format_specific('center', 'center', 'Calibri', 11, 
            'FF000000', False, False, False, [12, 17], [12, 16 + self.count])
            self.apply_format_specific('center', 'center', 'Calibri', 11, 
            'FF000000', False, False, False, [13, 17], [13, 16 + self.count])
            self.apply_format_specific('center', 'center', 'Calibri', 11, 
            'FF000000', False, False, False, [14, 17], [14, 16 + self.count])
            self.apply_format_specific('center', 'center', 'Calibri', 11, 
            'FF000000', False, False, False, [15, 17], [15, 16 + self.count])
            self.apply_format_specific('center', 'center', 'Calibri', 11, 
            'FF000000', False, False, False, [16, 17], [16, 16 + self.count])
            self.apply_format_specific('center', 'center', 'Calibri', 11, 
            'FF000000', False, False, False, [17, 17], [17, 16 + self.count])
            self.apply_format_specific('center', 'center', 'Calibri', 11, 
            'FF000000', False, False, False, [18, 17], [18, 16 + self.count])
            self.apply_format_specific('center', 'center', 'Calibri', 11, 
            'FF000000', False, False, False, [19, 17], [19, 16 + self.count])
            self.apply_format_specific('center', 'center', 'Calibri', 11, 
            'FF000000', False, False, False, [20, 17], [20, 16 + self.count])
            self.apply_format_specific('center', 'center', 'Calibri', 11, 
            'FF000000', False, False, False, [21, 17], [21, 16 + self.count])
            self.apply_format_specific('center', 'center', 'Calibri', 11, 
            'FF000000', False, False, False, [22, 17], [22, 16 + self.count])

            self.merge_cell_x_iterator(self.count, [self.rows - 2, 17], [self.rows, 17])
            self.merge_cell_xy([1, 17 + self.count], [5, 18 + self.count])
            self.merge_cell_x([1, 19 + self.count], [5, 19 + self.count])
            self.merge_cell_x_iterator(2, [6, 17 + self.count], [self.rows - 5, 17 + self.count])
            self.merge_cell_x([6, 19 + self.count], [11, 19 + self.count])
            self.merge_cell_x([12, 19 + self.count], [self.rows, 19 + self.count]) 
            self.merge_cell_x([1, 20 + self.count], [11, 20 + self.count])
            self.merge_cell_x([12, 20 + self.count], [self.rows - 5, 20 + self.count]) 

            self.merge_cell_x([self.rows - 4, 17 + self.count], [self.rows, 17 + self.count])
            self.merge_cell_x([self.rows - 4, 18 + self.count], [self.rows, 18 + self.count])
            self.merge_cell_x([self.rows - 4, 20 + self.count], [self.rows, 20 + self.count])
             
            self.write_cell('NOTES: TEST READING ARE WITHIN THE LIMIT', [1, 17 + self.count])
            self.write_cell('ACCEPTED BY:', [1, 19 + self.count])
            self.write_cell('COLD TESTED BY:', [6, 17 + self.count])
            self.write_cell('WITNESSED BY:', [6, 18 + self.count])
            self.write_cell('TEST INSTRUMENT:', [6, 19 + self.count])
            self.write_cell('SERIAL NUMBER:', [12, 19 + self.count])
            self.write_cell('CALIBRATION DATE: ', [12, 20 + self.count])        
            self.write_cell('DATE:', [self.rows - 4, 17 + self.count])  
            self.write_cell('DATE:', [self.rows - 4, 18 + self.count])
            self.write_cell('DATE: ', [self.rows - 4, 20 + self.count])    

            self.apply_border_default([1, self.cols], [self.rows, 20 + self.count])

            self.apply_format_specific('left', 'center', 'Arial', 11, 
            'FF000000', True, False, True, [1, 17 + self.count], [self.rows, 20 + self.count])

            self.elcb_merge() # Only for DB type

        # Final thick borders
        self.apply_border_rectangular('thick', None, [1, 1], [self.rows, 6])
        self.apply_border_rectangular('thick', 'thin', [1, 7], [self.rows, 13])
        self.apply_border_rectangular('thick', 'thin', [1, 14], [self.rows, 16])
        self.apply_border_rectangular('thick', 'thin', [1, 17], [self.rows, 17+self.count])
        self.apply_border_rectangular('thick', 'thin', [1, 17+self.count], [self.rows, 17+self.count+4])

        self.save_excel() # Always save changes

    def get_count(self, bool_sm):
        if bool_sm == True:
            x = 1
            y = 13
        else:
            x = 3
            y = 9

        if self.ws.cell(y, x).value != 1:
            print("cell ({}, {}) is not equal to 1! ({}, {}) = {}".format(x,y, x, y, self.ws.cell(y, x).value))
        else:
            i = 0
            previous_value = 1
            while self.ws.cell(y + i, x).value == previous_value:
                previous_value += 1
                i += 1

        #------ Error Handling ------#
        try:
            x = self.ws.cell(y, x).value/2 # Dummy code just for the purpose of Error Handling
        except TypeError:
            print("Program terminated!")
            sys.exit(1)
        #------ End of Error Handling ------#
        return i

    def copy_val_smdb(self):
        for i in range(self.count):
            self.circuit_ref_list.append(i + 1)
            self.circuit_designation_list.append(self.ws.cell(13 + i, 2).value)

            if (self.ws.cell(13 + i, 8).value) != None: # Some values could be None!
                if (self.ws.cell(13 + i, 8).value).lower().find('cx') == -1:
                    print("CX is not found!")
                else:
                    m, n = (self.ws.cell(13 + i, 8).value).lower().split('cx')
                    self.numcore_list.append(int(m))
                    self.cable_list.append(int(n))
            else:
                self.cable_list.append(None)
                self.numcore_list.append(None)

            if (self.ws.cell(13 + i, 9).value) != None: # Some values could be None!
                if (self.ws.cell(13 + i, 9).value).lower().find('cx') == -1:
                    print("CX is not found!")
                else:
                    m, n = (self.ws.cell(13 + i, 9).value).lower().split('cx')
                    self.ecc_list.append(int(n))
            else:
                self.ecc_list.append(None)
  
            self.mccbamp_list.append(self.ws.cell(13 + i, 5).value)
            self.mcbtype_list.append(None)
            self.mcbka_list.append(self.ws.cell(13 + i, 6).value)
            self.rccbma_list.append(None)

            if (self.ws.cell(13 + i, 8).value) != None: # I just chose 8, no big deal
                self.ry_list.append('>2000')
                self.yb_list.append('>2000')
                self.br_list.append('>2000')
                self.rybn_list.append('>2000')
                self.rybe_list.append('>2000')
                self.ne_list.append('>2000')
            else:
                self.ry_list.append(None)
                self.yb_list.append(None)
                self.br_list.append(None)
                self.rybn_list.append(None)
                self.rybe_list.append(None)
                self.ne_list.append(None)

            if (self.ws.cell(13 + i, 8).value) != None:
                self.continuitytest_list.append('OK')
                self.ring_list.append('N/A')
                self.resistance_list.append(round(random.uniform(0.1, 1.5), 1))
            else:
                self.continuitytest_list.append(None)
                self.ring_list.append(None)
                self.resistance_list.append(None)

            self.remarks_list.append(None)

    def copy_val_db(self):

        modulo_count = 0

        for i in range(self.count):
            self.circuit_ref_list.append(self.ws.cell(9 + i, 4).value)
            self.circuit_designation_list.append(self.ws.cell(9 + i, 8).value.upper())
            self.cable_list.append(self.ws.cell(9 + i, 6).value)
            self.cpc_list.append(self.ws.cell(9 + i, 6).value)

            #------- @TODO : Need implementation -------#
            self.numcore_list.append(None)
            self.mcbtype_list.append(None)
            self.mcbka_list.append(None)

            self.mccbamp_list.append(self.ws.cell(9 + i, 5).value)
            self.elcb_rating_list.append(self.ws.cell(9 + i, 2).value)

            if (self.ws.cell(9 + i, 8).value.upper()) != 'SPARE':
                self.ry_list.append('N/A')
                self.yb_list.append('N/A')
                self.br_list.append('N/A')
                self.ne_list.append('>2000')
                self.continuitytest_list.append('OK')
                self.ring_list.append('N/A')
                self.resistance_list.append(round(random.uniform(0.1, 1.5), 1))
            else:
                self.ry_list.append(None)
                self.yb_list.append(None)
                self.br_list.append(None)
                self.ne_list.append(None)
                self.continuitytest_list.append(None)
                self.ring_list.append(None)
                self.resistance_list.append(None)

            if (self.ws.cell(9 + i, 8).value.upper()) != 'SPARE':
                if (modulo_count%3) == 0: 
                    self.rn_list.append('>2000')
                    self.re_list.append('>2000')
                else:
                    self.rn_list.append(None)
                    self.re_list.append(None)

                if (modulo_count%3) == 1:
                    self.yn_list.append('>2000')
                    self.ye_list.append('>2000')
                else:
                    self.yn_list.append(None)
                    self.ye_list.append(None)

                if (modulo_count%3) == 2:
                    self.bn_list.append('>2000')
                    self.be_list.append('>2000')
                else:
                    self.bn_list.append(None)
                    self.be_list.append(None)
            else:
                    self.rn_list.append(None)
                    self.re_list.append(None)
                    self.yn_list.append(None)
                    self.ye_list.append(None)
                    self.bn_list.append(None)
                    self.be_list.append(None)

            self.remarks_list.append(None)
            modulo_count += 1

    def elcb_merge(self):
        none_count = 0
        previous_value_cell = 17

        for i in range(self.count):
            if self.ws.cell(18 + i, 9).value == None:
                none_count += 1
            else:
                if none_count > 0:
                    self.merge_cell_y_iterator(1, [9, previous_value_cell], [9, previous_value_cell + none_count])
                previous_value_cell = previous_value_cell + none_count + 1
                none_count = 0
            if i == (self.count - 1):
                self.merge_cell_y_iterator(1, [9, previous_value_cell], [9, previous_value_cell + none_count - 1])

    def paste_val(self):
        for i in range(len(self.paste_list)):
            for j in range(self.count):
                self.ws.cell(17 + j, 1 + i).value = self.paste_list[i][j]

    def copy_sheets_to_file(self):
        dirpath = os.getcwd()
        workbook_path = os.path.join(dirpath, self.folder_name)
        self.dummy_path = os.path.join(dirpath, 'dummy.xlsx')
        self.final_path = os.path.join(workbook_path, 'r_' + self.book_name)

        if os.path.exists(workbook_path) != True:
            os.mkdir(workbook_path)

        self.excel = Dispatch("Excel.Application")
        self.wb1 = self. excel.Workbooks.Open(Filename=self.dummy_path)


        if os.path.isfile(self.final_path) != True:
            self.wb2 = self.excel.Workbooks.Add()
            self.wb2.SaveAs(self.final_path)

        self.wb2 = self.excel.Workbooks.Open(Filename=self.final_path)

        self.ws1 = self.wb1.Worksheets(1)
        self.ws1.Copy(Before=self.wb2.Worksheets(self.wb2.Sheets.count))
        self.wb1.Close(SaveChanges=True)
        self.wb2.Close(SaveChanges=True)
        self.excel.Quit()