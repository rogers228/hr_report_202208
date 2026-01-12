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

class Report_sav10(tool_excel):
    def __init__(self, filename, USERNO, YM):
        # USERNO 工號
        # YM 查詢年月6碼
        self.fileName = filename
        self.USERNO = USERNO
        self.YM = YM
        self.report_name = 'sav10' # 加班確認表(月)
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
        # self.set_page_layout_horizontal()

    def output(self):
        caption = '加班確認表(月)' # 標題
        yStr = self.YM[:4]
        mStr = self.YM[4:6]
        YM_Days = tool_func.getYMdays(yStr, mStr) #該月天數
        row_days = 12   # 每行幾日
        subrows = 5 # 每行有幾個子行
        lis_w = [11.4] # 欄寬第一欄
        lis_wd = [7.4]*row_days # 欄寬第二欄之後
        lis_w.extend(lis_wd)

        self.c_write(1, 1, caption, font_A_12, alignment=ah_center_top) #標題
        self.c_merge(1, 1, 1, row_days+1)
        self.c_write(2, 1, f'年月:{self.YM}', font_A_10) #標題
        self.c_write(3, 1, f'人員:{self.USERNO} {self.hr.nogetName(self.USERNO)}', font_A_10) #標題

        self.c_write(4, 1, '請在超時下班處，確認是否加班，簽章', font_A_10, alignment=ah_center_top, border=bottom_border) #標題
        self.c_merge(4, 1, 4, row_days+1)
        self.c_column_width(lis_w) # 設定欄寬

        df_rd = self.hr.userym_Getrd_df(self.USERNO, self.YM ) # 出勤資料


        head = 3
        for d in range(1, YM_Days + 1):
            sr = (d - 1) // row_days + 1 # 每行 row_days 日，現在是第幾行
            sd = (d - 1) % row_days + 1 # 現在是第幾蘭
            # print(f'd:{d} / sr:{sr} / sd:{sd}')
            cr = (sr-1) * subrows
            #日期
            self.c_write(head+1+sr+cr, 1, '日',   font_A_10, border=bt_border)
            self.c_write(head+1+sr+cr, sd+1, d, font_A_10, border=bt_border)
            self.c_write(head+1+sr+cr+1, 1, '星期', font_A_10, border=bottom_border)
            self.c_write(head+1+sr+cr+1, sd+1, tool_func.getWeekdayStr(yStr, mStr, d), font_A_10, border=bottom_border)
            if df_rd is None:
                pass
            else:

                ymd14 = '{0}{1}000000'.format(self.YM, '{:0>2d}'.format(d))
                df_w = df_rd.loc[(df_rd['rd03'] == ymd14)] # 篩選
                if len(df_w.index) > 0:
                    rs = df_w.iloc[0]
                    lis_rd11 = rs['rd11'].split(',')
                    rd11 = '\n'.join(lis_rd11)
                else:
                    rd11 = ''

                self.c_write(head+1+sr+cr+2, 1, '', font_A_10, alignment=ah_wr, border=bottom_border)
                self.c_write(head+1+sr+cr+2, sd+1, rd11, font_A_10, alignment=ah_wr, border=bottom_border)


                self.c_write(head+1+sr+cr+3, 1, '私事無加班', font_A_10, alignment=ah_wr, border=bottom_border)
                self.c_write(head+1+sr+cr+3, sd+1, '', font_A_10, alignment=ah_wr, border=bottom_border)
                self.c_row_height(head+1+sr+cr+3, 24)
                self.c_write(head+1+sr+cr+4, 1, '加班已申請', font_A_10, alignment=ah_wr, border=bottom_border)
                self.c_write(head+1+sr+cr+4, sd+1, '', font_A_10, alignment=ah_wr, border=bottom_border)
                self.c_row_height(head+1+sr+cr+4, 24)
                if sr <= 2:
                    self.c_row_height(head+1+sr+cr+5, 90)

        cr = head+1+sr+cr+5
        self.c_write(cr, 1, '-結束- 以下空白', alignment=ah_center_top)
        self.c_merge(cr, 1,cr, row_days+1)

def test1():
    fileName = 'sav10' + '_' + time.strftime("%Y%m%d%H%M%S", time.localtime()) + '.xlsx'
    Report_sav10(fileName, 'AA0031', '202601')
    print('ok')

if __name__ == '__main__':
    test1()
    sys.exit() #正式結束程式
