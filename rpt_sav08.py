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

class Report_sav08(tool_excel):
    def __init__(self, filename, YMD, H1, H2):
        # YMD 查詢年月日8碼
        # H1, H2 起時2碼, 迄時2碼
        self.fileName = filename
        self.YMD = YMD
        self.H1 = H1
        self.H2 = H2
        self.report_name = 'sav08' # 出勤狀況表(當日)
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
        caption = '出勤狀況表(當日)' # 標題
        self.c_write(1, 1, caption, font_A_10) #標題
        self.c_write(2, 1, f'年月日: {self.YMD}', font_A_10)
        self.c_write(3, 1, f'查詢時間: {self.H1}~{self.H2}', font_A_10)


        lis_col = ['編號','姓名','刷卡時間','訂餐','訂餐備註']
        for i, c in enumerate(lis_col):
            self.c_write(5, i+1, c, font_A_10, border=bt_border)

        lis_w = [16,10,20,6,20]
        self.c_column_width(lis_w) # 設定欄寬

        ymdh1 = f'{self.YMD}{self.H1}'
        ymdh2 = f'{self.YMD}{self.H2}'

        df_sv = self.hr.ymdGersv_df(ymdh1, ymdh2) # 刷卡資料
        lis_ps = list(set(df_sv['ps02'].tolist())) # 人員唯一直
        lis_ps.sort()

        font = {'black':font_A_10, 'gray': font_A_10G} # 顏色
        sum_d = 0 # 訂餐數量
        print(sum_d)
        cr = 6
        for ps02 in lis_ps:
            self.c_write(cr, 1, ps02,   font_A_10, border=bottom_border) # 工號
            psid = self.hr.nogetId(ps02)
            self.c_write(cr, 2, self.hr.idgetName(psid), font_A_10, border=bottom_border) # 姓名
            # self.c_write(cr, 4, self.hr.idgetps32(psid), font_A_10, border=bottom_border) # 訂餐
            v = self.hr.idgetps32(psid); f2 = font['gray' if v == 0 else 'black']; self.c_write(cr, 4, v, f2, border=bottom_border) # 訂餐
            try:
                int(v) # 檢查可否為數字  方加總
                sum_d += v
            except:
                pass
            self.c_write(cr, 5, self.hr.idgetps33(psid), font_A_10, border=bottom_border) # 訂餐備註

            df_w = df_sv.loc[(df_sv['ps02'] == ps02)] # 篩選 (以人為群組，可能有多筆刷卡資料)
            lis_sv = list(map(lambda e: e[8:12], df_w['sv03'].tolist())) # 刷卡資料(取時分4碼)
            if len(df_w.index) > 0:
                self.c_write(cr, 3, ','.join(lis_sv), font_A_10, border=bottom_border) # 刷卡
            else:
                self.c_write(cr, 3, '', font_A_10, border=bottom_border) # 刷卡
            cr += 1

        self.c_write(4, 1, f'訂餐數量: {int(sum_d)}', font_A_10)
        self.c_write(cr, 1, '-結束- 以下空白', alignment=ah_center_top)
        self.c_merge(cr,1,cr,5)

def test1():
    fileName = 'sav08' + '_' + time.strftime("%Y%m%d%H%M%S", time.localtime()) + '.xlsx'
    Report_sav08(fileName, '20220822', '06', '09')
    print('ok')

if __name__ == '__main__':
    test1()
    sys.exit() #正式結束程式