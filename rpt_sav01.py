if True:
    import sys, custom_path
    config_path = custom_path.custom_path['hr_report_202208'] # 取得專案引用路徑
    sys.path.append(config_path)

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
        if True:
            # stlye
            a10 = font_A_10; a8 = font_A_8; bk = bottom_border_sk # seyle
            ahr = ah_right
            # func, method
            write=self.c_write; img=self.c_image; merge = self.c_merge; height = self.c_row_height
            # image
            img_sg = os.path.join(config_path,'image','acg4.jpg') # 裝訂符號
            img_n1 = os.path.join(config_path,'image','acg1.jpg') # 折線1
            img_n2 = os.path.join(config_path,'image','acg2.jpg') # 折線1
            img_n3 = os.path.join(config_path,'image','acg3.jpg') # 折線1
            mask = self.mask_tolist() # 保密遮罩

        if True: # data
            # 人員
            if self.userno_arr is None:
                df_rs = self.hr.wuGetrs_df(self.YM, '') 
            else:
                df_rs = self.hr.wuGetrs_df(self.YM, self.userno_arr) 
            lis_ps = list(set(df_rs['ps02'].tolist())) #人員
            lis_ps.sort()
            # 薪資
            lis_rs = list(set(df_rs['rs01'].tolist())) 
            inStr = "(" + ",".join([str(e) for e in lis_rs]) + ")"
            df_rv = self.hr.Getrv_in_df(inStr) #薪資項目
            # 出勤
            df_rd = self.hr.ymGetrd_df(self.YM)

        self.c_column_width([34,8,8,10,10,30]) # 設定欄寬
        cr = 0; page1_rows = 53; page2_rows = 108
        for ps_i, psno in enumerate(lis_ps):
            psid = self.hr.nogetId(psno)
            ln = list(self.hr.nogetName(psno)); name_s = f'{ln[0]}**{ln[-1]}' # 姓名遮罩

            # page1
            cr+=1; write(cr, 1, caption, a10) #標題
            write(cr, 6, ps_i+1, a10, alignment=ahr) # 人碼
            cr+=1; write(cr, 1, f'薪資: {self.YM[:4]}年 {self.YM[4:6]}月', a10)
            cr+=1; write(cr, 1, f'人員: {psno} {name_s}', a10)

            df_w = df_rs.loc[(df_rs['ps02'] == psno)] #人員
            for i, r in df_w.iterrows():
                cr+=1; write(cr, 1, f"實領薪資總額: {r['rs08']:,.0f}", a10)
                cr+=1; write(cr, 1, f"轉帳金額(一): {r['rs10']:,.0f}", a10)
                cr+=1; write(cr, 1, f"轉帳金額(二): {r['rs11']:,.0f}", a10)
                cr+=1; write(cr, 1, f"轉帳金額(三): {r['rs12']:,.0f}", a10)
                cr+=1; lis_column = ['薪資項目','數量','單位','單價','金額','備註']
                for j, col in enumerate(lis_column):
                    ah2 = ahr if j in [1,3,4] else ah_left
                    write(cr, j+1, col, a10, alignment=ah2, border=bt_border)

                df_wrv = df_rv.loc[(df_rv['rv01'] == r['rs01'])] #薪資項目
                for j, v in df_wrv.iterrows():
                    cr+=1; write(cr, 1, f"{v['pa08']} {v['pa02']}", a10, border=bk)
                    write(cr, 2, v['rv04'], a10, alignment=ahr, border=bk)
                    write(cr, 3, v['rv03'], a10, border=bk)
                    write(cr, 4, v['rv05'], a10, alignment=ahr, border=bk)
                    write(cr, 5, v['rv06'], a10, alignment=ahr, border=bk)
                    if len(v['rv07']) > 16: #備註過長直接換行顯示
                        write(cr, 6, '', a10, border=bk)
                        cr+=1; write(cr, 1, '*' + v['rv07'], a10, border=bk); merge(cr,1,cr,6) # 換行
                    else:
                        write(cr, 6, v['rv07'], a10, border=bk) #備註

                cr+=1; write(cr, 1, '-結束- 以下空白', a8, alignment=ah_center, border=top_border); merge(cr,1,cr,6)

                # 特休資訊
                mdays = self.hr.idgetps23(psid) # 人員可特休天數
                tdays = self.hr.Gersw_6_df(psid, self.YM[:4]) # 人員年度特休以請天數
                if tdays is None:
                    tdays = 0
                holidayStr =f'{self.YM[:4]}年度特休假：可休 {mdays:.1f} 天，已請 {tdays:.1f} 天，剩餘 {float(mdays) -float(tdays):.1f} 天。'
                cr+=2; write(cr, 1, holidayStr, a10, alignment=ah_left_center, border=thin_border); merge(cr,1,cr,6); height(cr, 28)
                
                # 勞退資訊
                laps_a = self.hr.idgetps56(psid) # 月提繳工資
                laps_s = self.hr.idgetps57(psid) # 雇主提繳金額
                cr+=2; height(cr, 28)
                if all([laps_a>0, laps_s>0]):
                    lapsStr = f'新制勞工退休金：月提繳工資 {laps_a:.0f} 雇主提繳率 {laps_s/laps_a:.0%} 提繳金額 {laps_s:.0f}。'
                    write(cr, 1, lapsStr, a10, alignment=ah_left_center, border=thin_border); merge(cr,1,cr,6)

                # Email資訊
                cr+=2; height(cr, 28)
                if self.hr.idgetps14(psid) == '': 
                    mailMemo = '***您尚未設定Email !! 請設定以免遺漏重要訊息。'
                    write(cr, 1, mailMemo, a10, alignment=ah_left_center, border=thin_border); merge(cr,1,cr,6)

                # page2
                cr = page1_rows+1 if ps_i==0 else (ps_i*page2_rows) + page1_rows+1
                # 遮罩 mask 頁首
                for mi in range(11):
                    dic_case = {
                        0: f'{mask[mi]}—   {ps_i+1}', # 人碼
                        10: f'——      {psno} {name_s}   敬啟  ——'} # 收件人
                    write(cr+mi, 6, dic_case.get(mi, mask[mi]), a10, alignment=ahr)

                img(cr+1,  6, img_sg,20, 85, 0.1,2.95) # 裝訂符號
                img(cr,    4, img_n3,20,199, 0.1,0)    # 折線3
                img(cr+12, 1, img_n2,362,19, 0.1,2.8) # 折線2

                for mi in range(4): # 頁尾
                    write(cr+51+mi, 6, mask[mi], a10, alignment=ahr)

                #打卡出勤紀錄
                df_wrd = df_rd.loc[(df_rd['rd02'] == psid)]
                df_wrd.reset_index(inplace=True) #重置索引
                # print(df_wrd)
                if len(df_wrd.index) > 0:
                    cr+=14; write(cr, 1, f'{self.YM[:4]}年 {self.YM[4:6]}月', a8)
                    cr+=1 ; write(cr, 1, '日期                    打卡時間', a8, border=bt_border); merge(cr,1,cr,4)
                    write(cr, 5, '出勤狀況', a8, border=bt_border); merge(cr,5,cr,6)
                    for rdi, r in df_wrd.iterrows():
                        dic_r = self.rd2value_dic(r)
                        cr+=1; write(cr, 1, self.format_date(dic_r), a8, border=bk); merge(cr,1,cr,4)
                        write(cr, 5, self.format_state(dic_r), a8, border=bk); merge(cr,5,cr,6)
                        if rdi == 10:
                            cr +=1; img(cr, 1, img_n1,362,19, 0.1,2.8) # 折線1 位置實際列印調整

                    cr+=1; write(cr, 1, '-結束- 以下空白', a8, alignment=ah_center, border=top_border); merge(cr,1,cr,6)
                else:
                    cr+=14; write(cr, 1, '無打卡紀錄', a8)
                    cr+=13; img(cr, 1, img_n1,362,19, 0.1,2.8) # 折線1 位置實際列印調整

            cr += ((ps_i+1)*page2_rows)-cr # change user

    def rd2value_dic(self, pandas_row):
        # pandas row to dic
        dic = pandas_row.to_dict()

        lis_key = list(filter(lambda e: dic[e]!= 0, list(dic.keys()))) # 篩選出不為0的key
        lis_value = [dic[e] for e in lis_key]
        dic_v = dict(zip(lis_key, lis_value)) # 重新建構

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
        # 日   星期  類型  打卡
        # 20   三  出勤日  0721,1215,1733
        dic_rd04 = {0:'未設定', 1:'出勤日', 2:'公休日', 3:'周休日', 4:'國定假日', 5:'無薪公休日', 6:'不計'}
        d = dic['rd03']
        return f"{d[6:8]}   {tool_func.getWeekdayStr(d[0:4], d[4:6], d[6:8])}  {dic_rd04[dic.get('rd04', 0)]}  {dic['rd11']}"

    def format_state(self, dic):
        # 欲顯示的項目
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
        'rd35': '疫苗接種假', # 疫苗接種假
        'rd36': '返鄉探親假' # 返鄉探親假
        }

        lis = list(filter(lambda e: e in list(dic_rd.keys()), list(dic.keys()))) # 有出現在欲顯示項目中的項目
        return ' '.join([f'{dic_rd[e]}{dic[e]}' for e in lis])

    def mask_tolist(self):
        return [
            '▀▀▄▒▄▊▒▄■▀▓▒▄▓▒▊▄▄▒▀▓♕▒▊▊▒▊▊♞■▊▒▀■▊■▊■▒▒▒▒♕▊▄▄▒▒♞▊▒▄♞♕♞▒▊▄▒▊▄▄▒▄■▊▀▄▄♕',
            '♕▓♕▄▒▄▒▒▒▊▊▀▄▊▒▓■■▒▓♞▊▒▓▓▄▄▄▊■▒▊▀▊▒▄▊▀▊♕▊▒▒▄▒▒▊▊▒▄▒▊▄▄▒▊▊▄♞■▓▄▊■▊■▒▊▀▊',
            '■▊▓▓▒▒▒▊▊▒■■♞■▄▊▓▄▊♕▓♕▒▒▊▓▒▓■▊■▄▊▒▀▄▒▊▊■▓▒▊▊■♞▀▄▄▒▓▀■▒▒♞▄▒▓▓▓▀▊▓▓♞▀▊♞▓',
            '▊▊▓▀▒▄■■▓▒▊▒■▒▊▊▊♞▊▀▀▄■▒▒▒▊▊▊■▄▊▒▒▓▄▊▓▀▊▀▀▒■▊▓▊■♞▄▄▒▀▊♞▒▓▒▒▀♞▄■▊▓▊▒▒▒▊',
            '▓▊▊▄▒▀▒▓■▊▒▀▒▊▓▒▒▒▄■▒▊▊▀▒▓▓▒▓▊♕▒▓▒▊▓▒▒■♕■▓▊▓▒▊▀▊▓■▄■■■▒▒▄▓▊▊■▊■▒■♕▒♞▊▀',
            '▊♕▊▒▄▒▒♞▒▊■▀♕▊▀■♞▊▄▄▊▓▒▊♕▊♞▒▊▒▊▒▊▄▊▒▒▊■▄▀▄▄■▊▀▒▊▀▊▒▄▊▒▒■▊▒▓▒♞▓♕▒▒▒▀▄▒■',
            '▒▊▒■♞▒▊▒▓▓▊■■▄▒▄▊▒▒▓▀▒▊▄■▊▒▊▊▒▄▀♞▒▊▒▄♞▒▊▓▊♕▊▒▒▀▒▊▊♞▊▒▊▄▒▀▒▊▊▒▒▊■▒▒▒▊▊▒',
            '▊▒▄▀▓▀▓▒▊■■■▊▊♕▊▓▀▓♞▀♕▒▄♞♕▒▊■▊▒▒■▀▊▒▀▊▒▊■▊▄▒▒▊▊▊▊▒▊▄♞▊■▊▒▊▄▄▒▒■▒▊▀■▒▄▒',
            '▊▊▒▀▊▊▊▊▊▒▒▊▄▊▀■♞■▀▓▀▊▄▊▊♞▒■▄▊■▒▒▒▊▀■▊▊▄■♞▊▊▓▄▊▒▄■♕■▒▒▄▊▀▒▒▊■■▊▒♕▒▊▄♞▊',
            '▄▊▒▊▊▒▀▊▒▓▒▊▒▄▊▒▒▓▊▀▒■■▓▀▊▒♕▄▄▓▒▒▒▀▄■♞▄▊▄▒▀■▀▒■▒▊▒▀▒♞▊▀▓▊▄▓▒▊▊▒▒■▊■▓▄▄',
            '▒♞▀▊▊■■▒▒▓▀▄▓▒▀♞▒▊▊▊▀■▓▊▄▊▒▓▒▒♕▒■▀♞▊▊▀▀♕♕■▒▊▓▒■▒▊▊■■▊▀▄♕▒▊▓▄▒▀■▒♞▀■▄▓▊',
            '■▀▒▊▀♕▒▊▀♞♞▊▒▄▒▒▊♕▒▒▊▄▓▄■▒▊▒▓▒▓▓▊▊▊▊▊▄▊▊▊■▓▒▒▒▒▊▒■▓▒♞♞■■▒♞▒▊▒▊▒▒■▄▒▀♞■',
            '▒▒▊▊▊▒▒▀▒▒▊▄■▊▊▓▊▄▒■▒▒▀▀▊▊♕▒♞▓▓▊▒▄▀▄▊▒▀■▊▄▊▒▓▄▊▀▒▒▀▒▊▄▊▓■▒▊♞▒▒▒■▊▒■▊▒▀',
            '▒▊■▊▄▒■▄▓▓▊▓■▊▒▒▊▄▊■▄▒▓▒♞▄▓▊▒▓■♕▓■▒▀▀▀▓▒▒▀▀♕▓▒▀▊♞▊▒▊▒▒▊▀▒■▒▀▄♞▊♞▄▊▊▀▒■',
            '▒■▊■▊▊▒■▀▒▊▊▒▊♕▊▄▊♕▊♞▊▄■♕▓▀▊▒▓▄▊▒▀▒▊■▊▊▄▊▊▊▊▒■▒▊▊▀■▊▀▀▒▊▒▀■▓♞▊▊♕▓▓♞▓▒▊',
            '▊▊▀▊▒▊▓▒▒▄▊▓▓▄■▊▄▒▓▒▒▒▀▀▊▒▊▊▊♞▒▒■♞▒▒▓▓▀▊■♕▀▒▒▀■▒▒▒▒■▊▓■▒■■▓▒▓▀♕▊▒▄▀■■▊'
            ]

def test1():
    timer1 = time.perf_counter()
    fileName = 'sav01' + '_' + time.strftime("%Y%m%d%H%M%S", time.localtime()) + '.xlsx'
    Report_sav01(fileName, '202208', 'AA0031,AA0094')
    # Report_sav01(fileName, '202208', 'AA0031')
    # Report_sav01(fileName, '202208')
    print('運算時間:',time.perf_counter()-timer1)

if __name__ == '__main__':
    test1()
    sys.exit() #正式結束程式