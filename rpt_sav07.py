import sys, os, time

import openpyxl
from tool_excel2 import tool_excel
from tool_style import *
import tool_file
import tool_db_hr
from config import *

class Report_sav07(tool_excel):
    def __init__(self, filename):
        self.fileName = filename
        self.report_path = config_report_path.get(os.getenv('COMPUTERNAME'), None)
        if self.report_path is None:
            print('找不到路徑')
            raise SystemExit  #結束
        self.clear_temp()    # 清除暫存
        self.create_excel()  # 建立

        self.hr = tool_db_hr.db_hr()
        self.output()
        self.save_xls()
        self.open_excel() #開啟

    def clear_temp(self): # 清除暫存
        ftl = tool_file.File_tool()
        ftl.clear()

    def create_excel(self):
        wb = openpyxl.Workbook()
        sh = wb.active
        # sh.title = self.report_name
        sh.title = self.fileName.split('_')[0] # report_name
        self.xlsfile = os.path.join(self.report_path, self.fileName)
        wb.save(filename = self.xlsfile)
        super().__init__(self.xlsfile, wb, sh) # 傳遞引數給父class
        self.set_page_layout() # 頁面設定layout

    def open_excel(self):
        if os.path.exists(self.xlsfile): #檔案存在
            # 使用cmd 使用excel啟動 最大化 該檔案
            cmd = r'start "" /max EXCEL.EXE "' + self.xlsfile + '"'
            # print(cmd)
            os.system(cmd)

    def output(self):
        caption = '職務代理人清冊' # 標題
        self.c_write(1, 1, caption, font_11, ah_center) #標題
        self.c_merge(1,1,1,5)
        self.c_column_width([18, 18, 18, 18, 30]) # 設定欄寬
        lis_title = ['人員','職務代理人','代理誰職務','我請假通知誰','誰請假通知我']
        for i, title in enumerate(lis_title):
            self.c_write(2, i+1, title, font_9, fillcolor=cf_gray) #標題

        df = self.hr.df_ps_atwork() # 在職人員
        cr = 3
        for idx, r in df.iterrows():
            self.c_write(cr, 1, f"{r['ps02']} {r['ps03']}", border=bottom_border) # 人員

            tmp = '\n'.join([f"{no} {self.hr.nogetName(no)}" for no in r['ps12'].split(',')])
            self.c_write(cr, 2, tmp, alignment=ah_wr, border=bottom_border) # 職務代理人

            lis_dku = self.hr.ps02Getdku_lis(r['ps02'])
            tmp = '\n'.join([f"{no} {self.hr.nogetName(no)}" for no in lis_dku])
            self.c_write(cr, 3, tmp, alignment=ah_wr, border=bottom_border) # 代理誰職務

            tmp = '\n'.join([f"{no} {self.hr.nogetName(no)}" for no in r['ps52'].split(',')]) 
            self.c_write(cr, 4, tmp, alignment=ah_wr, border=bottom_border) # 請假通知人

            lis_dbt = self.hr.ps02Getdbt_lis(r['ps02'])
            tmp = '\n'.join([f"{no} {self.hr.nogetName(no)}" for no in lis_dku])
            self.c_write(cr, 5, tmp, alignment=ah_wr, border=bottom_border) # 誰請假通知我

            cr += 1

        self.c_write(cr, 1, '-結束- 以下空白', alignment=ah_center_top)
        self.c_merge(cr,1,cr,5)




def test1():
    fileName = 'sav07' + '_' + time.strftime("%Y%m%d%H%M%S", time.localtime()) + '.xlsx'
    Report_sav07(fileName)
    print('ok')

if __name__ == '__main__':
    test1()
    sys.exit() #正式結束程式