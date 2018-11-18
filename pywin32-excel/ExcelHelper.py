"""
****************************************
* Author: SIRIUS
* Email: xuqingskywolf@outlook.com
* Created Time: 2018/11/18 21:40
****************************************
"""

import win32com.client
import logging


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

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.close()

    def close(self):
        num_of_workbooks = self.excel.Workbooks.Count
        if num_of_workbooks > 0:
            logging.debug(
                'there are still %d workbooks opened in excel process, '
                'not quit excel application',
                num_of_workbooks
            )
        else:
            logging.debug(
                'no workbook opened in excel process, quiting excel '
                'application instance ...'
            )
        self.workbook.Close(SaveChanges=0)
        self.excel.Quit()
        del self.excel

    def save(self, new_file_name=None):
        if new_file_name:
            self.workbook.SaveAs(new_file_name)
        else:
            self.workbook.Save()

    def set_worksheet(self, sheet):
        self.worksheet = self.workbook.Worksheets(sheet)

    def get_cell(self, row, col):
        return self.worksheet.Cells(row, col).Value

    def get_cell_text(self, row, col):
        return self.worksheet.Cells(row, col).Text

    def set_cell(self, row, col, value):
        self.worksheet.Cells(row, col).Value = value

    def get_range(self, row1, col1, row2, col2):
        cell1 = self.worksheet.Cells(row1, col1)
        cell2 = self.worksheet.Cells(row2, col2)
        return self.worksheet.Range(cell1, cell2).Value

    def del_range(self, row1, col1, row2, col2):
        cell1 = self.worksheet.Cells(row1, col1)
        cell2 = self.worksheet.Cells(row2, col2)
        return self.worksheet.Range(cell1, cell2).Delete

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
