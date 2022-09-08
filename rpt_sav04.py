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

class Report_sav04(tool_excel):
    def __init__(self, filename, YM):
        # YM 查詢年月6碼
        self.fileName = filename
        self.YM = YM
        self.hr = tool_db_hr.db_hr() # 資料庫
        self.df_rs = self.hr.ymGetrs_df(self.YM)
        if self.df_rs is None:
            print('nd data!')
            sys.exit()

        self.report_name = 'sav04' # 薪資轉帳明細表
        self.report_dir = config_report_dir # 資料夾名稱
        self.report_path = os.path.join(os.path.expanduser(r'~\Documents'), self.report_dir) #資料夾路徑
        
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
        self.wb = openpyxl.Workbook()
        wb = self.wb
        self.df_ca = self.hr.Getca_df() # 公司
        for cai, car in self.df_ca.iterrows():
            wb.create_sheet(car['ca02'])
        wb.remove(wb['Sheet'])

        self.xlsfile = os.path.join(self.report_path, self.fileName)
        wb.save(filename = self.xlsfile)
        # super().__init__(self.xlsfile, wb, sh) # 傳遞引數給父class
        # self.set_page_layout() # 頁面設定layout

    def output(self):
        wb = self.wb
        df_ca = self.df_ca
        df_rs = self.df_rs
        for cai, car in self.df_ca.iterrows():
            sh = wb[car['ca02']]
            super().__init__(self.xlsfile, wb, sh) # 傳遞引數給父class
            self.set_page_layout() # 頁面設定layout

            caption = '薪資轉帳明細表' # 標題
            lis_w =[16,16,16,16,16,16]
            self.c_column_width(lis_w) # 設定欄寬

            # lis_ca = list(set(df_rs['ps40'].tolist()))
            # lis_ca.sort()
            font = {'black':font_A_10, 'gray': font_A_10G} # 顏色
            align = {'al': ah_left, 'ar': ah_right}
            cr = 1

            self.c_write(cr, 1, caption, font_A_10) #標題
            self.c_write(cr+1, 1,  f"公司: {car['ca03']}", font_A_10)
            self.c_write(cr+2, 1,  f'薪資年月: {self.YM}', font_A_10)
            lis_col = ['人員','姓名','轉帳(一)','轉帳(二)','轉帳(三)','實領薪資']
            for i, c in enumerate(lis_col):
                a2 = align['ar' if i >=2 else 'al']
                self.c_write(cr+3, i+1, c, font_A_10, alignment=a2, border=bt_border)
            cr += 4

            df_w = df_rs.loc[(df_rs['ps40'] == car['ca01'])] # 篩選 (以公司為群組)
            for i, r in df_w.iterrows():
                self.c_write(cr, 1, r['ps02'], font_A_10, border=bottom_border)
                self.c_write(cr, 2, r['ps03'], font_A_10, border=bottom_border)
                v = r['rs10']; f2 = font['gray' if v == 0 else 'black']; self.c_write(cr, 3, v, f2, alignment=ah_right, border=bottom_border)
                v = r['rs11']; f2 = font['gray' if v == 0 else 'black']; self.c_write(cr, 4, v, f2, alignment=ah_right, border=bottom_border)
                v = r['rs12']; f2 = font['gray' if v == 0 else 'black']; self.c_write(cr, 5, v, f2, alignment=ah_right, border=bottom_border)
                v = r['rs08']; f2 = font['gray' if v == 0 else 'black']; self.c_write(cr, 6, v, f2, alignment=ah_right, border=bottom_border)
                cr += 1

            self.c_write(cr, 1,  '合計', font_A_10, border=bottom_border)
            self.c_write(cr, 2,  f'{len(df_w.index)}人', font_A_10, border=bottom_border)
            self.c_write(cr, 3,  f"{df_w['rs10'].sum():,.0f}", font_A_10, alignment = ah_right, border=bottom_border)
            self.c_write(cr, 4,  f"{df_w['rs11'].sum():,.0f}", font_A_10, alignment = ah_right, border=bottom_border)
            self.c_write(cr, 5,  f"{df_w['rs12'].sum():,.0f}", font_A_10, alignment = ah_right, border=bottom_border)
            self.c_write(cr, 6,  f"{df_w['rs08'].sum():,.0f}", font_A_10, alignment = ah_right, border=bottom_border)
            cr += 1; self.c_write(cr, 1, '-結束- 以下空白', alignment=ah_center_top); self.c_merge(cr,1,cr,6)  

def test1():
    fileName = 'sav04' + '_' + time.strftime("%Y%m%d%H%M%S", time.localtime()) + '.xlsx'
    Report_sav04(fileName, '202208')
    print('ok')

if __name__ == '__main__':
    test1()
    sys.exit() #正式結束程式