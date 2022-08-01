import openpyxl
from openpyxl.utils import get_column_letter #轉換
from tool_style import *

class tool_excel(): #讀取excel 單一零件
    def __init__(self, file, workbook, sh):
        self.file = file
        self.workbook = workbook
        self.sh = sh # excel sheet

    def c_write(self, row, column, value = '', font = font_9, alignment = ah_left, border = no_border, fillcolor = cf_none):
        #寫入儲存格 並設定格式
        cell = self.sh.cell(row, column)
        cell.value = value
        if font:
            cell.font = font
        if alignment:
            cell.alignment = alignment
        if border:
            cell.border = border
        if fillcolor:
            cell.fill = fillcolor

    def c_merge(self, start_row, start_column, end_row, end_column):
        self.sh.merge_cells(start_row=start_row, start_column=start_column, end_row=end_row, end_column=end_column) #合併儲存格

    def set_page_layout(self): # 頁面設定layout
        cm2in = lambda x: x/2.54
        self.sh.page_margins = openpyxl.worksheet.page.PageMargins(
            left=cm2in(1.2),
            right=cm2in(1.2),
            top=cm2in(1.0),
            bottom=cm2in(0.5),
            header=cm2in(1.0),
            footer=cm2in(0.5))

    def c_column_width(self, width_list): # 設定欄寬
        for i in range(len(width_list)):
            self.sh.column_dimensions[get_column_letter(i+1)].width = width_list[i]

    def save_xls(self): # 儲存
        try:
            self.workbook.save(self.file) #save
        except:
            print('儲存時發生錯誤，無法處理該檔案，有可能檔案已被開啟尚未關閉!')

def test1():
    print('test1')

if __name__ == '__main__':
    test1()
    print('ok')