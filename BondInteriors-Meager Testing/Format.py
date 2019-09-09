# Format.py

from MergeCell import *
from Library import *

class Format(MergeCell):

    def write_cell(self, text, *args):
        __x, __y = args[0]
        self.ws.cell(__y, __x).value = text

    def apply_width(self, column, width):
        offset_deficit = 0.72
        self.ws.column_dimensions[column].width = width + offset_deficit

    def apply_height(self, row, height):
        offset_deficit = 0.0
        self.ws.row_dimensions[row].height = height + offset_deficit

    def add_logo(self, image_name):
        dirpath = os.getcwd()
        img_name = image_name
        img_path = os.path.join(dirpath, img_name)
        c2e = cm_to_EMU
        p2e = pixels_to_EMU

        img = Image(img_path)
        position = XDRPoint2D(p2e(121), p2e(7))
        size = XDRPositiveSize2D(p2e(116.78), p2e(120))
        img.anchor = AbsoluteAnchor(pos=position, ext=size)
        self.ws.add_image(img)

    def apply_border_default(self, *args):
        __x1, __y1 = args[0]
        __x2, __y2 = args[1]

        border = Border(left=Side(border_style='thin', color='000000'),
                    right=Side(border_style='thin', color='000000'),
                    top=Side(border_style='thin', color='000000'),
                    bottom=Side(border_style='thin', color='000000'))

        for i in range(__x2 - __x1 + 1):
            for j in range(__y2 - __y1 + 1):
                self.ws.cell(__y1 + j, __x1 + i).border = border

    def apply_border_rectangular(self, border_style, other_style, *args):
        __x1, __y1 = args[0]
        __x2, __y2 = args[1]

        border = Border(left=Side(border_style=other_style, color='000000'),
                    right=Side(border_style=other_style, color='000000'),
                    top=Side(border_style=border_style, color='000000'),
                    bottom=Side(border_style=other_style, color='000000'))

        for i in range(__x2 - __x1 + 1):
            self.ws.cell(__y1, __x1 + i).border = border

        border = Border(left=Side(border_style=other_style, color='000000'),
                    right=Side(border_style=other_style, color='000000'),
                    top=Side(border_style=other_style, color='000000'),
                    bottom=Side(border_style=border_style, color='000000'))

        for i in range(__x2 - __x1 + 1):
            self.ws.cell(__y2, __x1 + i).border = border

        for i in range(__y2 - __y1 + 1):
                if __y1 + i == __y1:
                    border = Border(left=Side(border_style=other_style, color='000000'),
                    right=Side(border_style=border_style, color='000000'),
                    top=Side(border_style=border_style, color='000000'),
                    bottom=Side(border_style=other_style, color='000000'))
                elif __y1 + i == __y2:
                    border = Border(left=Side(border_style=other_style, color='000000'),
                    right=Side(border_style=border_style, color='000000'),
                    top=Side(border_style=other_style, color='000000'),
                    bottom=Side(border_style=border_style, color='000000'))
                else:
                    border = Border(left=Side(border_style=other_style, color='000000'),
                    right=Side(border_style=border_style, color='000000'),
                    top=Side(border_style=other_style, color='000000'),
                    bottom=Side(border_style=other_style, color='000000'))

                self.ws.cell(__y1 + i, __x2).border = border

        for i in range(__y2 - __y1 + 1):
                if __y1 + i == __y1:
                    border = Border(left=Side(border_style=border_style, color='000000'),
                    right=Side(border_style=other_style, color='000000'),
                    top=Side(border_style=border_style, color='000000'),
                    bottom=Side(border_style=other_style, color='000000'))
                elif __y1 + i == __y2:
                    border = Border(left=Side(border_style=border_style, color='000000'),
                    right=Side(border_style=other_style, color='000000'),
                    top=Side(border_style=other_style, color='000000'),
                    bottom=Side(border_style=border_style, color='000000'))
                else:
                    border = Border(left=Side(border_style=border_style, color='000000'),
                    right=Side(border_style=other_style, color='000000'),
                    top=Side(border_style=other_style, color='000000'),
                    bottom=Side(border_style=other_style, color='000000'))

                self.ws.cell(__y1 + i, __x1).border = border

    def apply_format_default(self):
        self.apply_width("B", 48.43)
        self.apply_height(1, 26.25)
        self.apply_height(7, 18.75)

        for i in range(9):
            self.apply_height(8+i, 15.75)
        self.add_logo("BondLogo.png")

    def apply_format_specific(self, horizontal, vertical, font_type, font_size, font_color, bold, italic, wrap_text, *args):

        alignment=Alignment(horizontal=horizontal, vertical=vertical, wrap_text= wrap_text)
        font = Font(name=font_type,
                    size=font_size,
                    bold=bold,
                    italic=italic,
                    color=font_color)

        __x1, __y1 = args[0]
        __x2, __y2 = args[1]
        for i in range(__x2 - __x1 + 1):
            for j in range(__y2 - __y1 + 1):
                self.ws.cell(__y1 + j, __x1 + i).font = font
                self.ws.cell(__y1 + j, __x1 + i).alignment = alignment
    
    def insert_cells(self):
        for x in range(self.cols):
            self.ws.insert_cols(1)
        for x in range(self.rows):
            self.ws.insert_rows(1)

    def create_sheet_format(self, *data):

        header_list, merge_x_list, merge_y_list, merge_xy_list, format_list, border_list = self.data_list

        for i in range(len(header_list)):
            self.write_cell(header_list[i][0], header_list[i][1])

        for i in range(len(merge_x_list)):
            self.merge_cell_x_iterator(merge_x_list[i][0], merge_x_list[i][1], merge_x_list[i][2])

        for i in range(len(merge_y_list)):
            self.merge_cell_y_iterator(merge_y_list[i][0], merge_y_list[i][1], merge_y_list[i][2])

        for i in range(len(merge_xy_list)):
            self.merge_cell_xy(merge_xy_list[i][0], merge_xy_list[i][1])

        for i in range(len(format_list)):
            self.apply_format_specific(format_list[i][0], format_list[i][1], format_list[i][2], format_list[i][3], format_list[i][4], format_list[i][5], format_list[i][6], format_list[i][7], format_list[i][8], format_list[i][9])