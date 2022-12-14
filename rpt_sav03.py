if True:
    import sys, custom_path
    config_path = custom_path.custom_path['hr_report_202208'] # 取得專案引用路徑
    sys.path.append(config_path) # 載入專案路徑

import os, time

import openpyxl
from tool_excel2 import tool_excel
from tool_style import *
import tool_file
import tool_db_hr
import tool_func
from config import *

class Report_sav03(tool_excel):
    def __init__(self, filename, YM1, YM2):
        # YM1 查詢年月6碼  起
        # YM1 查詢年月6碼  迄
        self.fileName = filename
        self.YM1 = YM1; self.YM2 = YM2
        self.report_name = 'sav03' # 出勤統計表
        self.report_dir = config_report_dir # 資料夾名稱
        self.report_path = os.path.join(os.path.expanduser(r'~\Documents'), self.report_dir) #資料夾路徑

        self.hr = tool_db_hr.db_hr() # 資料庫
        self.file_tool = tool_file.File_tool() # 檔案工具並初始化資料夾
        if self.report_path is None:
            print('找不到路徑')
            raise SystemExit  #結束

        self.file_tool.clear(self.report_name) # 清除舊檔

        self.create_excel()  # 建立
        self.output()
        self.save_xls()
        self.open_xls() # 開啟

    def create_excel(self):
        wb = openpyxl.Workbook()
        sh = wb.active
        # sh.title = self.report_name
        sh.title = self.report_name
        self.xlsfile = os.path.join(self.report_path, self.fileName)
        wb.save(filename = self.xlsfile)
        super().__init__(self.xlsfile, wb, sh) # 傳遞引數給父class
        self.set_page_layout() # 頁面設定layout

    def output(self):
        caption = '出勤統計表' # 標題
        self.c_write(1, 1, caption, font_A_10) #標題
        self.c_write(2, 1, f'出勤年月:{self.YM1}~{self.YM2}', font_A_10) #標題


        df = self.hr.ymGetrd_sum_df(self.YM1, tool_func.getNextymStr(self.YM2)) # 出勤統計
        
        fname = ['人員',      '姓名',    '出勤(天)',   '缺勤(天)','遲到(分鐘)',
                 '加班(小時)','0~2加班', '2以上',      '公休',    '周休',
                 '國定假日',  '無薪公休','不計天數',    '特休',    '公假',
                 '婚假',      '喪假',    '產假',       '病假',    '事假',
                 '陪產假',    '產檢假',  '育嬰假',     '留職停薪', '實際遲到(分鐘)',
                 '補刷卡次數','防疫照顧假','疫苗接種假','返鄉探親假'
                 ]
        fsWidth =[8,10,6,6,6,
                  6,6,6,6,6,
                  6,6,6,6,6,
                  6,6,6,6,6,
                  6,6,6,6,6,
                  6,6,6,6
                  ] #欄寬
        self.c_column_width(fsWidth) # 設定欄寬

        for i, e in enumerate(fname):
            self.c_write(3, i+1, e, font_A_10, alignment=ah_wr, border=bt_border) # 欄位名稱

        font = {'black':font_A_10, 'gray': font_A_10G}
        cr = 4
        for idx, r in df.iterrows():
            v = r['ps02']; self.c_write(cr, 1, v, font_A_10, border=bottom_border)
            v = r['ps03']; self.c_write(cr, 2, v, font_A_10, border=bottom_border)
            v = r['Srd06']; f2 = font['gray' if v == 0 else 'black']; self.c_write(cr, 3, v, f2, alignment = ah_right, border=bottom_border)
            v = r['Srd07']; f2 = font['gray' if v == 0 else 'black']; self.c_write(cr, 4, v, f2, alignment = ah_right, border=bottom_border)
            v = r['Srd08']; f2 = font['gray' if v == 0 else 'black']; self.c_write(cr, 5, v, f2, alignment = ah_right, border=bottom_border)
            v = r['Srd09']; f2 = font['gray' if v == 0 else 'black']; self.c_write(cr, 6, v, f2, alignment = ah_right, border=bottom_border)
            v = r['Srd19']; f2 = font['gray' if v == 0 else 'black']; self.c_write(cr, 7, v, f2, alignment = ah_right, border=bottom_border)
            v = r['Srd20']; f2 = font['gray' if v == 0 else 'black']; self.c_write(cr, 8, v, f2, alignment = ah_right, border=bottom_border)
            v = r['Srd14']; f2 = font['gray' if v == 0 else 'black']; self.c_write(cr, 9, v, f2, alignment = ah_right, border=bottom_border)
            v = r['Srd15']; f2 = font['gray' if v == 0 else 'black']; self.c_write(cr, 10, v, f2, alignment = ah_right, border=bottom_border)
            v = r['Srd16']; f2 = font['gray' if v == 0 else 'black']; self.c_write(cr, 11, v, f2, alignment = ah_right, border=bottom_border)
            v = r['Srd17']; f2 = font['gray' if v == 0 else 'black']; self.c_write(cr, 12, v, f2, alignment = ah_right, border=bottom_border)
            v = r['Srd18']; f2 = font['gray' if v == 0 else 'black']; self.c_write(cr, 13, v, f2, alignment = ah_right, border=bottom_border)
            v = r['Srd21']; f2 = font['gray' if v == 0 else 'black']; self.c_write(cr, 14, v, f2, alignment = ah_right, border=bottom_border)
            v = r['Srd22']; f2 = font['gray' if v == 0 else 'black']; self.c_write(cr, 15, v, f2, alignment = ah_right, border=bottom_border)
            v = r['Srd23']; f2 = font['gray' if v == 0 else 'black']; self.c_write(cr, 16, v, f2, alignment = ah_right, border=bottom_border)
            v = r['Srd24']; f2 = font['gray' if v == 0 else 'black']; self.c_write(cr, 17, v, f2, alignment = ah_right, border=bottom_border)
            v = r['Srd25']; f2 = font['gray' if v == 0 else 'black']; self.c_write(cr, 18, v, f2, alignment = ah_right, border=bottom_border)
            v = r['Srd26']; f2 = font['gray' if v == 0 else 'black']; self.c_write(cr, 19, v, f2, alignment = ah_right, border=bottom_border)
            v = r['Srd27']; f2 = font['gray' if v == 0 else 'black']; self.c_write(cr, 20, v, f2, alignment = ah_right, border=bottom_border)
            v = r['Srd28']; f2 = font['gray' if v == 0 else 'black']; self.c_write(cr, 21, v, f2, alignment = ah_right, border=bottom_border)
            v = r['Srd29']; f2 = font['gray' if v == 0 else 'black']; self.c_write(cr, 22, v, f2, alignment = ah_right, border=bottom_border)
            v = r['Srd30']; f2 = font['gray' if v == 0 else 'black']; self.c_write(cr, 23, v, f2, alignment = ah_right, border=bottom_border)
            v = r['Srd31']; f2 = font['gray' if v == 0 else 'black']; self.c_write(cr, 24, v, f2, alignment = ah_right, border=bottom_border)
            v = r['Srd32']; f2 = font['gray' if v == 0 else 'black']; self.c_write(cr, 25, v, f2, alignment = ah_right, border=bottom_border)
            v = r['Srd33']; f2 = font['gray' if v == 0 else 'black']; self.c_write(cr, 26, v, f2, alignment = ah_right, border=bottom_border)
            v = r['Srd34']; f2 = font['gray' if v == 0 else 'black']; self.c_write(cr, 27, v, f2, alignment = ah_right, border=bottom_border)
            v = r['Srd35']; f2 = font['gray' if v == 0 else 'black']; self.c_write(cr, 28, v, f2, alignment = ah_right, border=bottom_border)
            v = r['Srd36']; f2 = font['gray' if v == 0 else 'black']; self.c_write(cr, 29, v, f2, alignment = ah_right, border=bottom_border)
            cr += 1

        self.c_write(cr, 1, '-結束- 以下空白', alignment=ah_center_top)
        self.c_merge(cr,1,cr, len(fname))

def test1():
    fileName = 'sav03' + '_' + time.strftime("%Y%m%d%H%M%S", time.localtime()) + '.xlsx'
    Report_sav03(fileName, '202211', '202211')
    print('ok')

if __name__ == '__main__':
    test1()
    sys.exit() #正式結束程式