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
from config import *

class Report_sav02(tool_excel):
    def __init__(self, filename):
        self.fileName = filename
        self.report_name = 'sav02' # 薪資項目明細表
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
        self.open_excel() #開啟

    def create_excel(self):
        wb = openpyxl.Workbook()
        sh = wb.active
        # sh.title = self.report_name
        sh.title = self.report_name
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
        caption = '薪資項目明細表' # 標題
        lis_pa = ['0010','0020','0030','0040','0050','0060','0070','0210','0220','A001'] # 薪資項目
        self.c_write(1, 1, caption, font_A_10) #標題
        self.c_write(2, 1, '薪資項目', font_A_10)
        self.c_write(3, 1, '人員', font_A_10, border=bottom_border, fillcolor=cf_gray)
        self.c_write(3, 2, '名稱', font_A_10, border=bottom_border, fillcolor=cf_gray)
        lis_w = [9, 10]; lis_wd = [8]*len(lis_pa); lis_w.extend(lis_wd) # 欄寬
        self.c_column_width(lis_w) # 設定欄寬
        for j, pano in enumerate(lis_pa):
            self.c_write(2, j+3, pano, font_A_10)
            self.c_write(3, j+3, self.hr.pa08Getpa02(pano), font_A_10, alignment = ah_wr, border=bottom_border, fillcolor=cf_gray) # 標題

        whereSTR = "pa08 IN ('" +   "','".join(lis_pa) + "')"
        df_pf = self.hr.wGerpf_df(whereSTR) # 人員薪資項目
        lis_ps = list(set(df_pf['ps02'].tolist()))
        lis_ps.sort()
        cr = 4
        for i, psno in enumerate(lis_ps):
            self.c_write(cr, 1, psno, font_A_10, border=bottom_border) # 工號
            self.c_write(cr, 2, self.hr.nogetName(psno), font_A_10, border=bottom_border) # 姓名
            for j, pano in enumerate(lis_pa):
                df_w = df_pf.loc[(df_pf['ps02'] == psno) & (df_pf['pa08'] == pano)] # 篩選
                data_v = df_w.iloc[0]['pf05'] if len(df_w.index) > 0 else '' # 單價
                self.c_write(cr, j+3, data_v, font_A_10, border=bottom_border)
            cr += 1

        self.c_write(cr, 1, '-結束- 以下空白', alignment=ah_center_top)
        self.c_merge(cr,1,cr, len(lis_pa)+2)

def test1():
    fileName = 'sav02' + '_' + time.strftime("%Y%m%d%H%M%S", time.localtime()) + '.xlsx'
    Report_sav02(fileName)
    print('ok')

if __name__ == '__main__':
    test1()
    sys.exit() #正式結束程式