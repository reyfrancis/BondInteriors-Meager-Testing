# MergeCell.py

from Library import *

class MergeCell(object): 

    def merge_cell_xy(self, *args):
        __x1, __y1 = args[0]
        __x2, __y2 = args[1]
        self.ws.merge_cells(start_row=__y1, start_column=__x1, end_row=__y2, end_column=__x2)

    def merge_cell_x(self, *args):
        __x1, __y1 = args[0]
        __x2, __y2 = args[1]
        self.ws.merge_cells(start_row=__y1, start_column=__x1, end_row=__y1, end_column=__x2)

    def merge_cell_y(self, *args):
        __x1, __y1 = args[0]
        __x2, __y2 = args[1]
        self.ws.merge_cells(start_row=__y1, start_column=__x1, end_row=__y2, end_column=__x1)

    def merge_cell_x_iterator(self, iterate_count, *args):
        __x1, __y1 = args[0]
        __x2, __y2 = args[1]
        for i in range(iterate_count):   
            self.merge_cell_x([__x1, __y1], [__x2, __y2])
            __y1 += 1

    def merge_cell_y_iterator(self, iterate_count, *args):
        __x1, __y1 = args[0]
        __x2, __y2 = args[1]
        for i in range(iterate_count):   
            self.merge_cell_y([__x1, __y1], [__x2, __y2])
            __x1 = __x1 + 1
