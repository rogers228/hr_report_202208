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

class Report_sav09(tool_excel):
    def __init__(self, filename, whereSTR):
        # YM 查詢年月6碼
        self.fileName = filename
        self.whereSTR = whereSTR
        self.report_name = 'sav09' # 人員基本資料
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
        # self.set_page_layout() # 頁面設定layout
        self.set_page_layout_horizontal()

    def output(self):
        caption = '人員基本資料' # 標題
        self.c_write(1, 1, caption, font_A_10) #標題
        dic_ps11 ={0:'無', 1:'就職', 2:'離職', 3:'留職'}
        self.c_column_width([8,12,6,12,6,9,9,28,
            6,6,8,8,9,20,20,12,12,20]) # 設定欄寬
        lis_title = ['人員編號','姓名','性別','身份證號','任職狀態','就職日期','離職日期',
        '通知Email','特休假天數','可特休天數','職稱','職等','生日','聯絡地址','戶籍地址','聯絡電話一','聯絡電話二',
        '班別']
        for i, title in enumerate(lis_title):
            self.c_write(2, i+1, title, font_A_10, alignment = ah_wr, border=bt_border, fillcolor=cf_gray) #標題

        df_ps = self.hr.wGetps_df(self.whereSTR) # 出勤資料
        df_ps[['ps08','ps27','ps35']] = df_ps[['ps08','ps27','ps35']].fillna(value='') #填充
        cr = 3; ci = 0
        for i, r in df_ps.iterrows():
            ci+=1; self.c_write(cr, ci, r['ps02'], font_A_10, border=bottom_border) # 人員編號
            ci+=1; self.c_write(cr, ci, r['ps03'], font_A_10, border=bottom_border) # 姓名
            ci+=1; self.c_write(cr, ci, r['ps05'], font_A_10, border=bottom_border) # 性別
            ci+=1; self.c_write(cr, ci, r['ps06'], font_A_10, border=bottom_border) # 身份證號
            ci+=1; self.c_write(cr, ci, dic_ps11.get(r['ps11'],''), font_A_10, border=bottom_border) # 任職狀態
            ci+=1; fdate = '' if r['ps08'] == '' else r['ps08'][0:8]; self.c_write(cr, ci, fdate, font_A_10, border=bottom_border) # 就職日期
            ci+=1; fdate = '' if r['ps35'] == '' else r['ps35'][0:8]; self.c_write(cr, ci, fdate, font_A_10, border=bottom_border) # 離職日期
            ci+=1; self.c_write(cr, ci, r['ps14'], font_A_10, alignment = ah_wr, border=bottom_border) # 通知Email
            ci+=1; self.c_write(cr, ci, r['ps22'], font_A_10, border=bottom_border) # 特休假天數
            ci+=1; self.c_write(cr, ci, r['ps23'], font_A_10, border=bottom_border) # 可特休天數
            ci+=1; self.c_write(cr, ci, r['ps25'], font_A_10, border=bottom_border) # 職稱
            ci+=1; self.c_write(cr, ci, r['ps26'], font_A_10, border=bottom_border) # 職等
            ci+=1; fdate = '' if r['ps27'] == '' else r['ps27'][0:8]; self.c_write(cr, ci, fdate, font_A_10, border=bottom_border) # 生日
            ci+=1; self.c_write(cr, ci, r['ps28'], font_A_10, alignment = ah_wr, border=bottom_border) # 聯絡地址
            ci+=1; self.c_write(cr, ci, r['ps53'], font_A_10, alignment = ah_wr, border=bottom_border) # 戶籍地址
            ci+=1; self.c_write(cr, ci, r['ps29'], font_A_10, alignment = ah_wr, border=bottom_border) # 聯絡電話一
            ci+=1; self.c_write(cr, ci, r['ps30'], font_A_10, alignment = ah_wr, border=bottom_border) # 聯絡電話二
            ci+=1; self.c_write(cr, ci, r['bn02'], font_A_10, alignment = ah_wr, border=bottom_border) # 班別
            cr += 1; ci = 0
        # self.c_write(cr, 1, '-結束- 以下空白', alignment=ah_center_top)
        # self.c_merge(cr,1,cr, YM_Days+2)

def test1():
    fileName = 'sav09' + '_' + time.strftime("%Y%m%d%H%M%S", time.localtime()) + '.xlsx'
    Report_sav09(fileName, 'ps11>=0')
    print('ok')

if __name__ == '__main__':
    test1()
    sys.exit() #正式結束程式