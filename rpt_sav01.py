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

class Report_sav01(tool_excel):
    def __init__(self, filename, YM, userno_arr= ""):
        # YM 查詢年月6碼  起
        # userno_arr 使用者工號 AA0031,AA0094 文字陣列
        # userno_inSTR 使用者工號 "('AA0031','AA0094')"
        self.fileName = filename
        self.YM = YM
        self.userno_arr = userno_arr
        self.report_name = 'sav01' # 薪資單
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
        caption = '薪資單' # 標題
        if True: # seyle
            bk = bottom_border_sk

        self.c_column_width([34,8,8,10,10,30]) # 設定欄寬
        if True: # data
            # 人員
            # df_rs = self.hr.ymGetrs_df(self.YM) 
            # df_rs = self.hr.wuGetrs_df(self.YM, "('AA0031','AA0094')") 
            df_rs = self.hr.wuGetrs_df(self.YM, self.userno_arr) 
            lis_ps = list(set(df_rs['ps02'].tolist())) #人員
            lis_ps.sort()

            # 薪資
            lis_rs = list(set(df_rs['rs01'].tolist())) 
            inStr = "(" + ",".join([str(e) for e in lis_rs]) + ")"
            df_rv = self.hr.Getrv_in_df(inStr) #薪資項目

            # 出勤
            df_rd = self.hr.ymGetrd_df(self.YM)

        cr = 0; page1_rows = 55; page2_rows = 110
        for ps_i, psno in enumerate(lis_ps):
            psid = self.hr.nogetId(psno)

            # page1
            cr+=1; self.c_write(cr, 1, caption, font_A_10) #標題
            self.c_write(cr, 6, ps_i+1, font_A_10, alignment=ah_right) # 人碼
            cr+=1; self.c_write(cr, 1, f'計薪年月: {self.YM}', font_A_10)
            ln = list(self.hr.nogetName(psno)); name_s = f'{ln[0]}**{ln[-1]}'
            cr+=1; self.c_write(cr, 1, f'人員: {psno} {name_s}', font_A_10)

            df_w = df_rs.loc[(df_rs['ps02'] == psno)] #人員
            for i, r in df_w.iterrows():
                cr+=1; self.c_write(cr, 1, '實領薪資總額: ' + '{:,.0f}'.format(r['rs08']), font_A_10)
                cr+=1; self.c_write(cr, 1, '轉帳金額(一): ' + '{:,.0f}'.format(r['rs10']), font_A_10)
                cr+=1; self.c_write(cr, 1, '轉帳金額(二): ' + '{:,.0f}'.format(r['rs11']), font_A_10)
                cr+=1; self.c_write(cr, 1, '轉帳金額(三): ' + '{:,.0f}'.format(r['rs12']), font_A_10)

                cr+=1; lis_column = ['薪資項目','數量','單位','單價','金額','備註']
                for j, col in enumerate(lis_column):
                    ah2 = ah_right if j in [1,3,4] else ah_left
                    self.c_write(cr, j+1, col, font_A_10, alignment=ah2, border=bt_border)

                df_wrv = df_rv.loc[(df_rv['rv01'] == r['rs01'])] #薪資項目
                for j, v in df_wrv.iterrows():
                    cr+=1; 
                    self.c_write(cr, 1, f"{v['pa08']} {v['pa02']}", font_A_10, border=bk)
                    self.c_write(cr, 2, v['rv04'], font_A_10, alignment=ah_right, border=bk)
                    self.c_write(cr, 3, v['rv03'], font_A_10, border=bk)
                    self.c_write(cr, 4, v['rv05'], font_A_10, alignment=ah_right, border=bk)
                    self.c_write(cr, 5, v['rv06'], font_A_10, alignment=ah_right, border=bk)
                    if len(v['rv07']) > 16: #備註過長
                        self.c_write(cr, 6, '', font_A_10, border=bk)
                        cr+=1; self.c_write(cr, 1, '*' + v['rv07'], font_A_10, border=bk); self.c_merge(cr,1,cr,6) # 換行
                    else:
                        self.c_write(cr, 6, v['rv07'], font_A_10, border=bk) #備註

                cr+=1; self.c_write(cr, 1, '-結束- 以下空白', alignment=ah_center_top, border=top_border); self.c_merge(cr,1,cr,6)

                # 特休資訊
                mdays = self.hr.idgetps23(psid) # 人員可特休天數
                tdays = self.hr.Gersw_6_df(psid, self.YM[:4]) # 人員年度特休以請天數
                if tdays is None:
                    tdays = 0
                holidayStr ='{}年度特休假：可休 {:.1f} 天，已請 {:.1f} 天，剩餘 {:.1f} 天。'.format(
                    self.YM[:4], mdays, tdays, float(mdays) -float(tdays))

                cr+=2; self.c_write(cr, 1, holidayStr, font_A_10, alignment=ah_left_center, border=thin_border); self.c_merge(cr,1,cr,6)
                self.c_row_height(cr, 28)

                # Email
                if self.hr.idgetps14(psid) == '': 
                    mailMemo = '***您尚未設定Email !! 請設定以免遺漏重要訊息。'
                    cr+=1; self.c_write(cr, 1, mailMemo, font_A_10, alignment=ah_left_center); self.c_merge(cr,1,cr,6)

                # page2
                cr = ((ps_i+1)*page1_rows) + (ps_i*page1_rows) + 1
                # 遮罩 mask
                m_str = '''▒▊▉▄■▊▄▉▄■▊▒▀▀■▀▒▒▊▒▒▄▊▀▒▄▒■▀▉▀▒▀▉▀▄■▉▉▀▀▊■■▀▒■▒▊▉▒▄▀▀▒▉▒▄▉■▊▀▄■▀▊▊▀▒▒▀▄■▄▒▊▄
                            ▄▉▒▀▒▄▀▊▄▄▀▊▊■▒▀▄▉▉▄▉▒▀▄▊■▊▉▀▒▒▊▊▄▉▄▉▀▒▒▒▊▒▒▊▉▄▄▊▒▄▉▊▄■▒▀▄▄■▀▒▄▉■▄▀▊▀▒▒▀▄■▄▒▊▄▀▄▀▄▒▊▊
                            '''
                m_str = m_str.replace(' ','')
                m_random = [10,5,15,9,8,14,6,0,13,2,7,3,1,4,11,12];  m_lenght = 70
                for i in range(12):
                    mask = m_str[m_random[i]:m_random[i]+m_lenght]
                    if i == 0:
                        mv = f'{mask}   {ps_i+1}' # 人碼
                    elif i == 3:
                        mv = f'{mask}{psno}{name_s}■▊▒▀' # 收件人
                    else:
                        mv = mask
                    self.c_write(cr+i, 6, mv, font_A_10, alignment=ah_right)

                df_wrd = df_rd.loc[(df_rd['rd02'] == psid)] #人員 出勤            
                if len(df_wrd.index) > 0:
                    cr+=12; self.c_write(cr, 1, f'日期年月: {self.YM}', font_A_10)
                    cr+=1 ; self.c_write(cr, 1, '日期                刷卡時間', font_A_10, border=bt_border); self.c_merge(cr,1,cr,4)
                    self.c_write(cr, 5, '出勤狀況', font_A_10, border=bt_border); self.c_merge(cr,5,cr,6)
                    for rdi, r in df_wrd.iterrows():
                        cr+=1
                        dic_r = self.rd2value_dic(r)
                        self.c_write(cr, 1, self.format_date(dic_r), font_A_10, border=bk); self.c_merge(cr,1,cr,4)
                        self.c_write(cr, 5, self.format_state(dic_r), font_A_10, border=bk); self.c_merge(cr,5,cr,6)
                else:
                    cr+=12; self.c_write(cr, 1, '無刷卡紀錄', font_A_10, border=bk)


                cr+=1; self.c_write(cr, 1, '-結束- 以下空白', alignment=ah_center_top, border=top_border); self.c_merge(cr,1,cr,6)

            cr += ((ps_i+1)*page2_rows)-cr # change user

    def rd2value_dic(self, pandas_row):
        # pandas row to dic
        dic = pandas_row.to_dict()
        dic_v = {}
        for key in dic:
            if dic[key] != 0:
                dic_v[key] = dic[key]

        if dic_v['rd10'] == '': # 移除 無加班
            del dic_v['rd10']

        # 移除 無異常的項目
        dic_remove = {
            'rd06': '實際出勤(天)',
            'rd14': '公休天數',
            'rd15': '周休天數'}
        for key in dic_remove:
            if key in list(dic_v.keys()):
                if dic_v[key] == 1: # 移除 正常 1天
                    del dic_v[key]
        return dic_v

    def format_date(self, dic):
        dic_rd04 = {0:'未設定', 1:'出勤日', 2:'公休日', 3:'周休日', 4:'國定假日', 5:'無薪公休日', 6:'不計'}
        result = '{0}   {1}  {2}  {3}'.format(
            dic['rd03'][6:8],
            tool_func.getWeekdayStr(dic['rd03'][0:4], dic['rd03'][4:6], dic['rd03'][6:8]),
            dic_rd04[dic.get('rd04', 0)],
            dic['rd11']
            )
        return result

    def format_state(self, dic):
        dic_rd = {
        'rd06': '出', # 實際出勤(天)
        'rd07': '缺', # 實際缺勤(天)
        'rd08': '遲', # 核定遲到(分鐘)
        'rd10': '加', # 加班分段
        'rd14': '公休', # 公休天數
        'rd15': '周休', # 周休天數
        'rd16': '國定假日', # 國定假日天數
        'rd17': '無薪公休', # 無薪公休天數
        'rd18': '不計天數', # 不計天數
        'rd21': '特休', # 特休假天數
        'rd22': '公假', # 公假天數
        'rd23': '婚假', # 婚假天數
        'rd24': '喪假', # 喪假天數
        'rd25': '產假', # 產假天數
        'rd26': '病假', # 病假天數
        'rd27': '事假', # 事假天數
        'rd28': '陪產假', # 陪產假天數
        'rd29': '產檢假', # 產檢假天數
        'rd30': '育嬰假', # 育嬰假天數
        'rd31': '留職停薪', # 留職停薪天數
        'rd33': '補刷卡次數', # 補刷卡次數
        'rd34': '防疫照顧假', # 防疫照顧假
        'rd35': '疫苗接種假' # 疫苗接種假
        }
        t = ''
        for key in dic_rd:
            if key in dic:
                t += '{0}{1}, '.format(
                    dic_rd[key], dic[key])
        t = t.rstrip(' ')
        t = t.rstrip(',')
        return t

def test1():
    fileName = 'sav01' + '_' + time.strftime("%Y%m%d%H%M%S", time.localtime()) + '.xlsx'
    Report_sav01(fileName, '202207', 'AA0031, AA0094')
    # Report_sav01(fileName, '202207')
    print('ok')

if __name__ == '__main__':
    test1()
    sys.exit() #正式結束程式