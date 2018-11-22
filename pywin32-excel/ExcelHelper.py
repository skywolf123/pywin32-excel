"""
****************************************
* Author: SIRIUS
* Email: xuqingskywolf@outlook.com
* Created Time: 2018/11/18 21:40
****************************************
"""

import win32com.client
import re


class ExcelHelper:
    """

    """

    def __init__(self, file_name=None, is_debug=False):
        self.excel = win32com.client.DispatchEx('Excel.Application')
        if is_debug:
            self.excel.Visible = True
        if file_name:
            self.workbook = self.excel.Workbooks.Open(file_name)
        else:
            self.workbook = self.excel.Workbooks.Add()
        self.worksheet = self.workbook.Worksheets(1)
        self.worksheet_name = ''
        self.row_field_list = []
        self.col_field_list = []
        self.data_field = ''

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.close()

    def close(self):
        self.workbook.Close(SaveChanges=False)
        self.excel.Quit()
        del self.excel

    def save(self, new_file_name=None, is_xls=True):
        if new_file_name:
            if is_xls and int(self.excel.Version) >= 12:
                self.workbook.SaveAs(new_file_name, FileFormat=-4143)
            else:
                self.workbook.SaveAs(new_file_name)
        else:
            self.workbook.Save()

    def save_to_csv(self, new_file_name):
        self.excel.ActiveWorkbook.SaveAs(new_file_name, FileFormat=43)

    def save_to_txt(self, new_file_name):
        self.excel.AciveWorkbook.SaveAs(new_file_name, FileFormat=-4158)

    def show(self):
        self.excel.Visible = True

    def hide(self):
        self.excel.Visible = False

    @staticmethod
    def convert_number_to_alphabet(col_num):
        if col_num > 26:
            return chr(int(col_num / 26 + 64)) + chr(int(col_num % 26 + 64))
        else:
            return chr(col_num + 64)

    @staticmethod
    def convert_address_to_num(address):
        if isinstance(address, str):
            row = re.sub('^\D+', '', address)
            col = re.sub('\d+$', '', address)
            col_index = 0
            for c in col:
                col_index = col_index * 26 + ord(c) - 64
            return row, col_index
        else:
            return address

    def convert_address_to_cell(self, address):
        row, col = self.convert_address_to_num(address)
        return self.worksheet.Cells(row, col)

    def convert_address_to_range(self, address):
        cells = address.split(':')
        cell1 = self.convert_address_to_cell(cells[0])
        cell2 = self.convert_address_to_cell(cells[1])
        return self.worksheet.Range(cell1, cell2)

    def convert_cell_index(self, cell_index):
        if isinstance(cell_index, tuple):
            return self.worksheet.Cells(cell_index)
        else:
            return self.convert_address_to_cell(cell_index)

    def convert_range_index(self, range_index):
        if isinstance(range_index, tuple):
            cell1 = self.worksheet.Cells(range_index[0], range_index[1])
            cell2 = self.worksheet.Cells(range_index[2], range_index[3])
            return self.worksheet.Range(cell1, cell2)
        else:
            return self.convert_address_to_range(range_index)

    def set_worksheet(self, sheet):
        self.worksheet = self.workbook.Worksheets(sheet)

    def get_cell(self, cell_index):
        """
        Get value of one cell
        :param cell_index: tuple (row, column) or string 'ColumnRow'
        e.g. (1, 2) or 'A2'
        :return:
        """
        return self.convert_cell_index(cell_index).Value

    def get_cell_text(self, cell_index):
        return self.worksheet.Cells(cell_index).Text

    def set_cell(self, cell_index, value):
        """
        Set value of one cell
        :param cell_index: tuple (row, column) or string 'ColRow'
        e.g. (1, 2) or 'A2'
        :param value:
        :return:
        """
        self.convert_cell_index(cell_index).Value = value

    def set_cell_font(self, cell_index,
                      style='Regular', name='Arial', size=9, color_index=1):
        cell = self.convert_cell_index(cell_index).Value

        cell.Font.Size = size
        cell.ColorIndex = color_index
        for i, item in enumerate(style):
            if item.lower() == 'bold':
                cell.Font.Bold = True
            elif item.lower() == 'italic':
                cell.Font.Italic = True
            elif item.lower() == 'underline':
                cell.Font.Underline = True
            elif item.lower() == 'regular':
                cell.Font.FontStyle = 'Regular'
        cell.Font.Name = name

    def get_range(self, range_index):
        """
        Get value of a range of cells
        :param range_index: tuple (row1,col1,row2,col2) or string
        'ColRow1:ColRow2'
        e.g. (1,2,10,12) or 'A1:C3'
        :return:
        """
        return self.convert_range_index(range_index).Value

    def set_range(self, top_cell_index, data):
        if isinstance(top_cell_index, tuple):
            row1 = top_cell_index[0]
            col1 = top_cell_index[1]
        else:
            row1, col1 = self.convert_address_to_num(top_cell_index)
        row2 = row1 + len(data) - 1
        col2 = col1 + len(data[0]) - 1
        self.convert_range_index((row1, col1, row2, col2)).Value = data

    def clear_range(self, range_index, clear_contents=True, clear_formats=True):
        if clear_contents:
            self.convert_range_index(range_index).ClearContents()
        if clear_formats:
            self.convert_range_index(range_index).ClearFormats()

    def del_range(self, range_index):
        self.convert_range_index(range_index).Delete()

    def copy_and_paste(self, source, destination):
        source_sheet = self.workbook.Worksheets(source[0])
        destination_sheet = self.workbook.Worksheets(destination[0])
        temp_sheet = self.worksheet.Name
        self.worksheet = destination_sheet
        destination_range = self.convert_range_index(destination[1])
        self.worksheet = source_sheet
        self.convert_range_index(source[1]).Copy(destination_range)
        self.set_worksheet(temp_sheet)

    def add_sheet(self, old_sheet, new_sheet, after=True):
        sheet = self.workbook.Worksheets(old_sheet)
        if after:
            self.workbook.Worksheets.Add(None, sheet).Name = new_sheet
        else:
            self.workbook.Worksheets.Add(sheet).Name = new_sheet

    def del_sheet(self, sheet):
        self.excel.DisplayAlerts = False
        worksheet = self.workbook.Worksheets(sheet)
        worksheet.Activate()
        self.excel.ActivateSheet.Delete()

    def add_picture(self, picture_name, left, top, width, height):
        self.worksheet.Shapes.AddPicture(picture_name, 1, 1, left, top, width,
                                         height)

    def get_num_of_lines(self):
        return self.worksheet.Usedrange.Rows.Count

    def set_cell_font_color(self, row, col, color):
        self.worksheet.Cells(row, col).Font.Color = color

    def set_cell_font_bold(self, row, col):
        self.worksheet.Cells(row, col).Font.Bold = True

    '''
    def copy_sheet(self, before=None):
        shts = self.workbook.Worksheets
        shts(1).Copy(before, shts(1))
    '''


if __name__ == '__main__':
    pass
