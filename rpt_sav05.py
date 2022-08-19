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

class Report_sav05(tool_excel):
    def __init__(self, filename, YM):
        # YM 查詢年月6碼
        self.fileName = filename
        self.YM = YM
        self.report_name = 'sav05' # 出勤狀況表(月)
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
        caption = '出勤狀況表(月)' # 標題
        self.c_write(1, 1, caption, font_A_10) #標題

        self.c_write(2, 1, f'查詢時間:{self.YM}', font_A_10) #標題
        yStr = self.YM[:4]
        mStr = self.YM[4:6]
        YM_Days = tool_func.getYMdays(yStr, mStr) #該月天數
        self.c_write(3, 2, '日',   font_A_10, border=bt_border)
        self.c_write(4, 1, '人員', font_A_10, border=bottom_border)
        self.c_write(4, 2, '星期', font_A_10, border=bottom_border)

        #日期
        for d in range(1, YM_Days + 1):
            self.c_write(3, d+2, d, font_A_10, border=bt_border)
            self.c_write(4, d+2, tool_func.getWeekdayStr(yStr, mStr, d), font_A_10, border=bottom_border)

        lis_w = [8,9]
        lis_wd = [7.5]*YM_Days
        lis_w.extend(lis_wd)
        self.c_column_width(lis_w) # 設定欄寬

        df_rd = self.hr.ymGetrd_df(self.YM) # 出勤資料
        lis_ps = list(set(df_rd['rd02'].tolist())) # 人員唯一直
        cr = 5
        for rd02 in lis_ps:
            self.c_write(cr, 1, self.hr.idgetNo(rd02),   font_A_10, border=bottom_border) # 工號
            self.c_write(cr, 2, self.hr.idgetName(rd02), font_A_10, border=bottom_border) # 姓名
            for d in range(1, YM_Days + 1):
                ymd14 = '{0}{1}000000'.format(self.YM, '{:0>2d}'.format(d))
                df_w = df_rd.loc[(df_rd['rd02'] == rd02) & (df_rd['rd03'] == ymd14)] # 篩選
                if len(df_w.index) > 0:
                    rs = df_w.iloc[0]
                    lis_rd11 = rs['rd11'].split(',')
                    rd11 = '\n'.join(lis_rd11)
                    gStr = ''
                    gStr += '遲{:.0f},'.format(rs['rd08']) if rs['rd08'] > 0 else ''
                    gStr += '缺{:.2f},'.format(rs['rd07']) if rs['rd07'] > 0 else ''
                    color =  cf_yellow if rs['rd07'] > 0 else cf_none
                    gStr += '加{0},'.format(rs['rd10'].replace(',','-')) if rs['rd19'] > 0 else ''
                    gStr += '特{0},'.format(rs['rd21']) if rs['rd21'] > 0 else ''
                    gStr += '公{0},'.format(rs['rd22']) if rs['rd22'] > 0 else ''
                    gStr += '婚{0},'.format(rs['rd23']) if rs['rd23'] > 0 else ''
                    gStr += '喪{0},'.format(rs['rd24']) if rs['rd24'] > 0 else ''
                    gStr += '產{0},'.format(rs['rd25']) if rs['rd25'] > 0 else ''
                    gStr += '病{0},'.format(rs['rd26']) if rs['rd26'] > 0 else ''
                    gStr += '事{0},'.format(rs['rd27']) if rs['rd27'] > 0 else ''
                    gStr += '陪{0},'.format(rs['rd28']) if rs['rd28'] > 0 else ''
                    gStr += '檢{0},'.format(rs['rd29']) if rs['rd29'] > 0 else ''
                    gStr += '育{0},'.format(rs['rd30']) if rs['rd30'] > 0 else ''
                    gStr += '留{0},'.format(rs['rd31']) if rs['rd31'] > 0 else ''
                    gStr += '防疫照顧{0},'.format(rs['rd34']) if rs['rd34'] > 0 else ''
                    gStr += '疫苗接種{0},'.format(rs['rd35']) if rs['rd35'] > 0 else ''

                    if '未結算' in rs['rd12']:
                        gStr = '未結算'; color = cf_none

                    if '免刷卡' in rs['rd12']: # 不需要刷卡
                        gStr = gStr.replace('未結算', '')

                    gStr = gStr.rstrip(',')
                    gStr = '\n'.join(gStr.split(','))
                    if len(gStr) > 0:
                        rd11 += '\n' + gStr
                else:
                    rd11 = ''; color = cf_none
                
                self.c_write(cr, d+2, rd11, font_A_10, alignment=ah_wr, border=bottom_border, fillcolor = color)
            cr += 1

        self.c_write(cr, 1, '-結束- 以下空白', alignment=ah_center_top)
        self.c_merge(cr,1,cr, YM_Days+2)

def test1():
    fileName = 'sav05' + '_' + time.strftime("%Y%m%d%H%M%S", time.localtime()) + '.xlsx'
    Report_sav05(fileName, '202208')
    print('ok')

if __name__ == '__main__':
    test1()
    sys.exit() #正式結束程式