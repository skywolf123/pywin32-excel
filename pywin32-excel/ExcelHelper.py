"""
****************************************
* Author: SIRIUS
* Email: xuqingskywolf@outlook.com
* Created Time: 2018/11/23
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

    def set_worksheet(self, sheet):
        self.worksheet = self.workbook.Worksheets(sheet)

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

    def convert_name_to_sheet(self, sheet_name=None):
        if sheet_name is None:
            sheet = self.worksheet
        else:
            sheet = self.workbook.Worksheets(sheet_name)
        return sheet

    def convert_address_to_cell(self, address, sheet_name=None):
        sheet = self.convert_name_to_sheet(sheet_name)
        row, col = self.convert_address_to_num(address)
        return sheet.Cells(row, col)

    def convert_address_to_range(self, address, sheet_name=None):
        sheet = self.convert_name_to_sheet(sheet_name)
        cells = address.split(':')
        cell1 = self.convert_address_to_cell(cells[0], sheet_name)
        cell2 = self.convert_address_to_cell(cells[1], sheet_name)
        return sheet.Range(cell1, cell2)

    def convert_cell_index(self, cell_index, sheet_name=None):
        sheet = self.convert_name_to_sheet(sheet_name)
        if isinstance(cell_index, tuple):
            return sheet.Cells(cell_index)
        else:
            return self.convert_address_to_cell(cell_index, sheet_name)

    def convert_range_index(self, range_index, sheet_name=None):
        sheet = self.convert_name_to_sheet(sheet_name)
        if isinstance(range_index, tuple):
            cell1 = sheet.Cells(range_index[0], range_index[1])
            cell2 = sheet.Cells(range_index[2], range_index[3])
            return sheet.Range(cell1, cell2)
        else:
            return self.convert_address_to_range(range_index, sheet_name)

    # cell methods
    def get_cell(self, cell_index, sheet_name=None):
        """
        Get value of one cell
        :param sheet_name:
        :param cell_index: tuple (row, column) or string 'ColumnRow'
        e.g. (1, 2) or 'A2'
        :return:
        """
        return self.convert_cell_index(cell_index, sheet_name).Value

    def get_cell_text(self, cell_index, sheet_name=None):
        return self.worksheet.Cells(cell_index, sheet_name).Text

    def set_cell(self, cell_index, value, sheet_name=None):
        """
        Set value of one cell
        :param sheet_name:
        :param cell_index: tuple (row, column) or string 'ColRow'
        e.g. (1, 2) or 'A2'
        :param value:
        :return:
        """
        self.convert_cell_index(cell_index, sheet_name).Value = value

    def set_cell_font(self, cell_index, style='Regular', name='Arial',
                      size=9, color_index=1, sheet_name=None):
        cell = self.convert_cell_index(cell_index, sheet_name).Value

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

    # range methods
    def get_range(self, range_index, sheet_name=None):
        """
        Get value of a range of cells
        :param sheet_name:
        :param range_index: tuple (row1,col1,row2,col2) or string
        'ColRow1:ColRow2'
        e.g. (1,2,10,12) or 'A1:C3'
        :return:
        """
        return self.convert_range_index(range_index, sheet_name).Value

    def set_range(self, top_cell_index, data, sheet_name=None):
        if isinstance(top_cell_index, tuple):
            row1 = top_cell_index[0]
            col1 = top_cell_index[1]
        else:
            row1, col1 = self.convert_address_to_num(top_cell_index)
        row2 = row1 + len(data) - 1
        col2 = col1 + len(data[0]) - 1
        range = self.convert_range_index((row1, col1, row2, col2), sheet_name)
        range.Value = data

    def clear_range(self, range_index, clear_contents=True,
                    clear_formats=True, sheet_name=None):
        if clear_contents:
            self.convert_range_index(range_index, sheet_name).ClearContents()
        if clear_formats:
            self.convert_range_index(range_index, sheet_name).ClearFormats()

    def del_range(self, range_index, sheet_name=None):
        self.convert_range_index(range_index, sheet_name).Delete()

    def copy_and_paste(self, source, destination):
        source_sheet_name = source[0]
        source_range_index = source[1]
        dest_sheet_name = destination[0]
        dest_range_index = destination[1]
        dest_range = self.convert_range_index(dest_range_index,
                                              dest_sheet_name)
        source_range = self.convert_range_index(source_range_index,
                                                source_sheet_name)
        source_range.Copy(dest_range)

    def set_range_align(self, range_index, alignment='center',
                        sheet_name=None):
        alignment_dict = {'left': 2,
                          'center': 3,
                          'right': 4}
        range = self.convert_range_index(range_index, sheet_name)
        range.HorizontalAlignment = alignment_dict[alignment.lower()]

    def set_range_font(self, range_index, style='Regular', name='Arial',
                       size=9, color_index=1, sheet_name=None):
        range = self.convert_cell_index(range_index, sheet_name).Value

        range.Font.Size = size
        range.ColorIndex = color_index
        for i, item in enumerate(style):
            if item.lower() == 'bold':
                range.Font.Bold = True
            elif item.lower() == 'italic':
                range.Font.Italic = True
            elif item.lower() == 'underline':
                range.Font.Underline = True
            elif item.lower() == 'regular':
                range.Font.FontStyle = 'Regular'
        range.Font.Name = name

    # sheet methods
    def add_sheet(self, new_sheet, old_sheet=None, after=True):
        sheet = self.convert_name_to_sheet(old_sheet)
        if after:
            self.workbook.Worksheets.Add(None, sheet).Name = new_sheet
        else:
            self.workbook.Worksheets.Add(sheet).Name = new_sheet

    def del_sheet(self, sheet_name):
        self.excel.DisplayAlerts = False
        sheet = self.convert_name_to_sheet(sheet_name)
        sheet.Activate()
        self.excel.ActivateSheet.Delete()

    # other items methods
    def add_chart(self, left, top, width=240, height=160, sheet_name=None):
        sheet = self.convert_name_to_sheet(sheet_name)
        return sheet.ChartObjects().Add(left, top, width, height)

    def chart_plot(self, data_ranges, chart_object, chart_type,
                   plot_by=None, category_labels=1, series_labels=0,
                   has_legend=None, title=None, category_title=None,
                   value_title=None, extra_title=None, sheet_name=None):
        sheet = self.convert_name_to_sheet(sheet_name)
        top_row, left_col, bottom_row, right_col = data_ranges[0]
        source = sheet.Range(sheet.Cells(top_row, left_col),
                             sheet.Cells(bottom_row, right_col))
        if len(data_ranges) > 1:
            for count in range(len(data_ranges[1:])):
                top_row, left_col, bottom_row, right_col = data_ranges[
                    count + 1]
                temp_source = sheet.Range(sheet.Cells(top_row, left_col),
                                          sheet.Cells(bottom_row, right_col))
                source = self.excel.Union(source, temp_source)
        plot_type = {'Area': 1,
                     'Bar': 2,
                     'Column': 3,
                     'Line': 4,
                     'Pie': 5,
                     'Radar': -4151,
                     'Scatter': -4169,
                     'XYScatter': 72,  # Smooth
                     'XYScatterLines': 74,
                     'Combination': -4111,
                     '3DArea': -4098,
                     '3DBar': -4099,
                     '3DColumn': -4100,
                     '3DPie': -4101,
                     '3DSurface': -4103,
                     'Doughnut': -4120,
                     'Bubble': 15,
                     'Surface': 83,
                     'Cone': 3,
                     '3DAreaStacked': 78,
                     '3DColumnStacked': 55}
        gallery = plot_type[chart_type]
        chart_object.Chart.ChartWizard(source, gallery, format, plot_by,
                                       category_labels, series_labels,
                                       has_legend, title, category_title,
                                       value_title, extra_title)

    def copy_chart(self, source_chart_object, dest_chart_object):
        if isinstance(dest_chart_object, tuple):
            source_chart_object.Copy()
            sheet = self.workbook.Worksheets(dest_chart_object[0])
            sheet.Paste(sheet.Range(dest_chart_object[1]))
        else:
            source_chart_object.Chart.ChartArea.Copy()
            dest_chart_object.Chart.Paste()

    def cut_chart(self, source_chart_object, dest_chart_object):
        self.copy_chart(source_chart_object, dest_chart_object)
        source_chart_object.Delete()

    def hide_column(self, col, sheet_name=None):
        sheet = self.workbook.Worksheets(sheet_name)
        sheet.Columns(col).Hidden = True

    def hide_row(self, row, sheet_name=None):
        sheet = self.workbook.Worksheets(sheet_name)
        sheet.Rows(row).Hidden = True

    def excel_function(self, range, function, sheet_name=None):
        sheet = self.workbook.Worksheets(sheet_name)
        if isinstance(range, tuple):
            top_row = range[0]
            left_col = range[1]
            bottom_row = range[2]
            right_col = range[3]
            range_str = '(sheet.Range(sheet.Cells(top_row,left_col),' \
                        'sheet.Cells(bottom_row,right_col)))'
        else:
            range_str = '(sheet.Range(' + '"' + range + '"' + '))'
        function_str = 'self.excel.WorksheetFunction.'
        function_str += function + range_str
        return eval(function_str, globals(), locals())

    def add_comment(self, cell_index, comment='', sheet_name=None):
        cell = self.convert_cell_index(cell_index, sheet_name)
        if comment is None:
            cell.ClearComments()
        else:
            cell.AddComment(comment)

    def show_warning(self, enable=True):
        self.excel.DisplayAlerts = enable

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
