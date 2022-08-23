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

class Report_sav06(tool_excel):
    def __init__(self, filename, USERNO, YYYY):
        # USERNO 工號
        # YYYY 查詢年4碼
        self.fileName = filename
        self.USERNO = USERNO
        self.YYYY = YYYY
        self.report_name = 'sav06' # 出勤狀況表(個人年度)
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
        # self.set_page_layout() # 頁面設定layout
        self.set_page_layout_horizontal()

    def open_excel(self):
        if os.path.exists(self.xlsfile): #檔案存在
            # 使用cmd 使用excel啟動 最大化 該檔案
            cmd = r'start "" /max EXCEL.EXE "' + self.xlsfile + '"'
            # print(cmd)
            os.system(cmd)

    def output(self):
        caption = '出勤狀況表(個人年度)' # 標題
        self.c_write(1, 1, caption, font_A_10) #標題

        self.c_write(2, 1, f'年度: {self.YYYY}', font_A_10) #標題
        self.c_write(3, 1, f'人員: {self.USERNO} {self.hr.nogetName(self.USERNO)}', font_A_10) #標題
        self.c_column_width([16, 12, 24, 24, 8,8,8,8]) # 設定欄寬
        lis_title = ['日期','班別','刷卡','備註','缺勤(天)','出勤(天)','遲到(分鐘)','加班(小時)']
        for i, title in enumerate(lis_title):
            self.c_write(4, i+1, title, font_A_10, alignment = ah_wr, border=bt_border, fillcolor=cf_gray) #標題

        df_rd = self.hr.userGetrd_df(self.USERNO, self.YYYY) # 出勤資料
        dic_k = {0:'未設定', 1:'出勤日', 2:'公休日', 3:'周休日', 4:'國定假日', 5:'無薪公休日', 6:'不計'}
        font = {'black':font_A_10, 'gray': font_A_10G} # 顏色
        cr = 5
        for i, r in df_rd.iterrows():
            fd = '{0}/{1}({2})'.format(r['rd03'][4:6], r['rd03'][6:8],
                tool_func.getWeekdayStr(self.YYYY, r['rd03'][4:6], r['rd03'][6:8]))
            self.c_write(cr, 1, fd, font_A_10, border=bottom_border) #日期
            self.c_write(cr, 2, dic_k[r['rd04']], font_A_10, border=bottom_border) #班別
            self.c_write(cr, 3, r['rd11'], font_A_10, alignment = ah_wr, border=bottom_border) #刷卡
            self.c_write(cr, 4, r['rd12'], font_A_10, alignment = ah_wr, border=bottom_border) #備註
            color =  cf_yellow if r['rd07'] > 0 else cf_none
            if '未結算' in r['rd12']:
                color = cf_none            
            v = r['rd07']; f2 = font['gray' if v == 0 else 'black']; self.c_write(cr, 5, v, f2, border=bottom_border, fillcolor = color) # 缺勤

            v = r['rd06']; f2 = font['gray' if v == 0 else 'black']; self.c_write(cr, 6, v, f2, border=bottom_border) # 出勤
            v = r['rd08']; f2 = font['gray' if v == 0 else 'black']; self.c_write(cr, 7, v, f2, border=bottom_border) # 遲到
            self.c_write(cr, 8, r['rd10'], font_A_10, alignment = ah_wr, border=bottom_border) #加班
            cr += 1

        self.c_write(cr, 1, '-結束- 以下空白', alignment=ah_center_top)
        self.c_merge(cr,1,cr,8)

def test1():
    fileName = 'sav06' + '_' + time.strftime("%Y%m%d%H%M%S", time.localtime()) + '.xlsx'
    Report_sav06(fileName, 'AA0022', '2022')
    print('ok')

if __name__ == '__main__':
    test1()
    sys.exit() #正式結束程式