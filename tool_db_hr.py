if True:
    import sys, custom_path
    config_path = custom_path.custom_path['hr_report_202208'] # 取得專案引用路徑
    sys.path.append(config_path) # 載入專案路徑

import pandas as pd
# import pyodbc
from sqlalchemy.engine import URL
from sqlalchemy import create_engine
from config import *

class db_hr(): #讀取excel 單一零件
    def __init__(self):
        # self.cn = pyodbc.connect(config_conn_HR) # connect str 連接字串
        # self.rpt = pyodbc.connect(config_conn_RPT) # connect str 連接字串
        self.cn = create_engine(URL.create('mssql+pyodbc', query={'odbc_connect': config_conn_HR})).connect()
        self.rpt = create_engine(URL.create('mssql+pyodbc', query={'odbc_connect': config_conn_RPT})).connect()
        self.dbps = self.get_database_ps() # 建議一次性基本資料檔，避免多次存取db

    def runsql(self, SQL):
        try:
            cur = self.cn.cursor()
            cur.execute(SQL) #執行
            cur.commit() #更新
            cur.close() #關閉
        except:
            print(SQL)
            print('error class db_ab().def runsql()! 無法執行SQL!')

    def runsql_rpt(self, SQL):
        try:
            cur = self.rpt.cursor()
            cur.execute(SQL) #執行
            cur.commit() #更新
            cur.close() #關閉
        except:
            print(SQL)
            print('error class db_ab().def runsql()! 無法執行SQL!')

    def get_database_ps(self):
        s = "SELECT ps01,ps02,ps03,ps11,ps12,ps14,ps23,ps31,ps32,ps33,ps34,ps52,ps56,ps57 FROM rec_ps ORDER BY ps01"
        df = pd.read_sql(s, self.cn) #轉pd
        return df

    def nogetName(self, myno): #人員編號取得姓名
        ps = self.dbps
        df = ps.loc[ps['ps02'] == myno] # 篩選
        return df.iloc[0]['ps03'] if len(df.index) > 0 else ''

    def nogetId(self, myno): #依人員列表，人員編號取得id
        ps = self.dbps
        df = ps.loc[ps['ps02'] == myno] # 篩選
        return df.iloc[0]['ps01'] if len(df.index) > 0 else ''

    def idgetNo(self, myid): #依人員列表，人員ID取得人員編號
        ps = self.dbps
        df = ps.loc[ps['ps01'] == myid] # 篩選
        return df.iloc[0]['ps02'] if len(df.index) > 0 else ''

    def idgetName(self, myid): #依人員列表，人員ID取得人員姓名
        ps = self.dbps
        df = ps.loc[ps['ps01'] == myid] # 篩選
        return df.iloc[0]['ps03'] if len(df.index) > 0 else ''

    def idgetps14(self, myid): #人員ID取得 ps14通知Email
        ps = self.dbps
        df = ps.loc[ps['ps01'] == myid] # 篩選
        return df.iloc[0]['ps14'] if len(df.index) > 0 else ''

    def idgetps23(self, myid): #人員ID取得 ps23可特休天數
        ps = self.dbps
        df = ps.loc[ps['ps01'] == myid] # 篩選
        return df.iloc[0]['ps23'] if len(df.index) > 0 else 0

    def idgetps32(self, myid): #人員ID取得 ps32訂餐
        ps = self.dbps
        df = ps.loc[ps['ps01'] == myid] # 篩選
        return df.iloc[0]['ps32'] if len(df.index) > 0 else 0

    def idgetps33(self, myid): #人員ID取得 ps33訂餐備註
        ps = self.dbps
        df = ps.loc[ps['ps01'] == myid] # 篩選
        return df.iloc[0]['ps33'] if len(df.index) > 0 else ''

    def idgetps34(self, myid): #人員ID取得是否需要刷卡ps34
        ps = self.dbps
        df = ps.loc[ps['ps01'] == myid] # 篩選
        return df.iloc[0]['ps34'] if len(df.index) > 0 else 0

    def idgetps56(self, myid): #人員ID取得月提繳工資ps56
        ps = self.dbps
        df = ps.loc[ps['ps01'] == myid] # 篩選
        return df.iloc[0]['ps56'] if len(df.index) > 0 else 0

    def idgetps57(self, myid): #人員ID取得雇主提繳金額ps57
        ps = self.dbps
        df = ps.loc[ps['ps01'] == myid] # 篩選
        return df.iloc[0]['ps57'] if len(df.index) > 0 else 0

    def ps_atwork_df(self): #在職人員列表
        s = "SELECT ps01,ps02,ps03,ps12,ps52 FROM rec_ps WHERE ps11 = 1 ORDER BY ps02"
        df = pd.read_sql(s, self.cn) #轉pd
        return df if len(df.index) > 0 else None

    def ps02Getdku_lis(self, ps02): #代理誰職務，等同 職務代理人的反查
        ps = self.dbps
        df = ps.loc[(ps['ps12'].str.contains(ps02)) & (ps['ps11'] == 1)] # 篩選
        return df['ps02'].tolist() if len(df.index) > 0 else []

    def ps02Getdbt_lis(self, ps02): #誰請假通知我，等同 請假通知人的反查
        ps = self.dbps
        df = ps.loc[(ps['ps52'].str.contains(ps02)) & (ps['ps11'] == 1)] # 篩選
        return df['ps02'].tolist() if len(df.index) > 0 else []

    def ps18Getps01(self, ps18): #使用者電腦名稱 查詢 使用者代號
        s = "SELECT TOP 1 ps01 FROM rec_ps WHERE ps18 = '{0}'"
        s = s.format(ps18)
        df = pd.read_sql(s, self.cn) #轉pd
        return df.iloc[0]['ps01'] if len(df.index) > 0 else ''

    def pc02Getpc01(self, pc02): #設備 電腦名稱 查詢 設備代號
        s = "SELECT TOP 1 pc01 FROM rec_pc WHERE pc02 = '{0}'"
        s = s.format(pc02)
        df = pd.read_sql(s, self.rpt) #轉pd
        return df.iloc[0]['pc01'] if len(df.index) > 0 else ''

    def pa08Getpa02(self, pa08): #項目編號 查詢 項目名稱
        s = "SELECT TOP 1 pa02 FROM rec_pa WHERE pa08 LIKE '{0}'"
        s = s.format(pa08)
        df = pd.read_sql(s, self.cn) #轉pd
        return df.iloc[0]['pa02'] if len(df.index) > 0 else ''

    def ca01Getca03(self, ca01): #設備 電腦名稱 查詢 設備代號
        s = "SELECT TOP 1 ca03 FROM rec_ca WHERE ca01 = {0}"
        s = s.format(ca01)
        df = pd.read_sql(s, self.cn) #轉pd
        return df.iloc[0]['ca03'] if len(df.index) > 0 else ''

    def cpGer_qs_lis(self, qs01): # qs01使用者代號或設備代號 查詢權限 list
        s = "SELECT qs02 FROM rec_qs WHERE qs01 ='{0}'"
        s = s.format(qs01)
        df = pd.read_sql(s, self.rpt) #轉pd
        return df['qs02'].tolist() if len(df.index) > 0 else []

    def Getps_inwork_lis(self): #所有在職人員
        ps = self.dbps
        df = ps.loc[ps['ps11'] == 1] # 篩選 在職
        return df['ps02'].tolist() if len(df.index) > 0 else []

    def wGetps_df(self, whereSTR=''):
        # whereSTR: 不包含 WHERE的 where SQL語句
        s = """
            SELECT ps02,ps03,ps05,ps06,ps07,ps08,ps09,ps10,ps11,
                    ps12,ps13,ps14,ps22,ps23,ps25,ps26,ps27,ps28,ps29,
                    ps30,ps31,ps34,ps35,ps52,ps53,
                    bn02,ca02
            FROM rec_ps
                LEFT JOIN rec_bn ON ps31=bn01
                LEFT JOIN rec_ca ON ps40=ca01
            WHEREPLACESTR
            ORDER BY ps02 ASC
            """
        s = s.replace('WHEREPLACESTR','' if whereSTR =='' else f'WHERE {whereSTR}')
        df = pd.read_sql(s, self.cn) #轉pd
        return df if len(df.index) > 0 else None

    def Getca_df(self):
        s = "SELECT ca01,ca02,ca03,ca04 FROM rec_ca ORDER BY ca02 ASC"
        df = pd.read_sql(s, self.cn) #轉pd
        return df if len(df.index) > 0 else None

    def ymGetrd_df(self, ym):
        # ym 年月日6碼
        s = """
            SELECT rd02,rd03,rd04,rd06,rd07,rd08,rd10,rd11,rd12,rd14,rd15,rd16,rd17,rd18,rd19,
                    rd21,rd22,rd23,rd24,rd25,rd26,rd27,rd28,rd29,rd30,rd31,rd33,rd34,rd35
            FROM rec_rd
            WHERE rd03 LIKE '{0}%'
            ORDER BY rd03 ASC
            """
        s = s.format(ym)
        df = pd.read_sql(s, self.cn) #轉pd
        return df if len(df.index) > 0 else None

    def userym_Getrd_df(self, userno, ym):
        # ym 年月日6碼
        # userno 工號
        s = """
            SELECT rd02,ps02,ps03,rd03,rd11
            FROM rec_rd LEFT JOIN rec_ps ON rd02 = ps01
            WHERE
                ps02 LIKE '{0}' AND
                rd03 LIKE '{1}%'
            ORDER BY rd03 ASC
            """
        s = s.format(userno, ym)
        df = pd.read_sql(s, self.cn) #轉pd
        return df if len(df.index) > 0 else None

    def userGetrd_df(self, userno, year):
        # userno 工號
        # year 4碼
        s = """
            SELECT rd02,rd03,rd04,rd06,rd07,rd08,rd10,rd11,rd12,rd19,
                    rd21,rd22,rd23,rd24,rd25,rd26,rd27,rd28,rd29,rd30,rd31,rd34,rd35
            FROM rec_rd LEFT JOIN rec_ps ON rd02=ps01
            WHERE
                ps02 LIKE '{0}%' AND
                rd03 LIKE '{1}%'
            ORDER BY rd03 DESC
            """
        s = s.format(userno, year)
        df = pd.read_sql(s, self.cn) #轉pd
        return df if len(df.index) > 0 else None

    def ymGetrs_df(self, ym):
        # ym 年月日6碼
        s = """
            SELECT rs01,ps02,ps03,ps40,rs02,rs08,rs10,rs11,rs12
            FROM rec_ps
                LEFT JOIN rec_rs ON ps01=rs02
            WHERE rs03 LIKE '{0}%'
            ORDER BY ps02 ASC
            """
        s = s.format(ym)
        df = pd.read_sql(s, self.cn) #轉pd
        return df if len(df.index) > 0 else None

    def wuGetrs_df(self, ym, userno_arr=''):
        # ym 年月日6碼
        # userno_arr 使用者工號 AA0031,AA0094 文字陣列

        # 轉 userno_inSTR 使用者工號 "('AA0031','AA0094')"
        # print('ym:',ym)
        # print('userno_arr:', userno_arr)
        if userno_arr == "":
            userno_inSTR = ""
        else:
            userno_arr = str(userno_arr).replace(' ','') # 去除空格
            userno_inSTR = "('" + "','".join(userno_arr.split(',')) + "')"
        s = """
            SELECT rs01,ps02,ps03,ps40,rs02,rs08,rs10,rs11,rs12
            FROM rec_ps
                LEFT JOIN rec_rs ON ps01=rs02
            WHERE rs03 LIKE '{0}%'
            WHEREPLACESTR
            ORDER BY ps02 ASC
            """
        s = s.format(ym)
        s = s.replace('WHEREPLACESTR','' if userno_inSTR =='' else f' AND ps02 IN {userno_inSTR}')
        df = pd.read_sql(s, self.cn) #轉pd
        return df if len(df.index) > 0 else None

    def Getrv_in_df(self, inSTR=''):
        # 薪資項目明細
        s = """
            SELECT rv01,rv02,rv03,rv04,rv05,rv06,rv07,
                pa02,pa08
            FROM rec_rv
                LEFT JOIN rec_pa ON rv02=pa01
            WHERE
                rv01 in {0}
            ORDER BY pa08 ASC
            """
        s = s.format(inSTR)
        df = pd.read_sql(s, self.cn) #轉pd
        return df if len(df.index) > 0 else None

    def ymGetrd_sum_df(self, ym1, ym2):
        # ym 年月日6碼
        s = """
            SELECT rd02, ps02, ps03,
                Sum(rd06) AS Srd06,
                Sum(rd07) AS Srd07,
                Sum(rd08) AS Srd08,
                Sum(rd09) AS Srd09,
                Sum(rd14) AS Srd14,
                Sum(rd15) AS Srd15,
                Sum(rd16) AS Srd16,
                Sum(rd17) AS Srd17,
                Sum(rd18) AS Srd18,
                Sum(rd19) AS Srd19,
                Sum(rd20) AS Srd20,
                Sum(rd21) AS Srd21,
                Sum(rd22) AS Srd22,
                Sum(rd23) AS Srd23,
                Sum(rd24) AS Srd24,
                Sum(rd25) AS Srd25,
                Sum(rd26) AS Srd26,
                Sum(rd27) AS Srd27,
                Sum(rd28) AS Srd28,
                Sum(rd29) AS Srd29,
                Sum(rd30) AS Srd30,
                Sum(rd31) AS Srd31,
                Sum(rd32) AS Srd32,
                Sum(rd33) AS Srd33,
                Sum(rd34) AS Srd34,
                Sum(rd35) AS Srd35,
                Sum(rd36) AS Srd36
            FROM rec_rd LEFT JOIN rec_ps ON rec_rd.rd02 = rec_ps.ps01
            WHERE
                rd03 > '{0}' AND
                rd03 < '{1}'
            GROUP BY rd02,ps02,ps03
            ORDER BY ps02 ASC
            """
        s = s.format(ym1, ym2)
        df = pd.read_sql(s, self.cn) #轉pd
        return df if len(df.index) > 0 else None

    def ymdGersv_df(self, ymdh1, ymdh2):
        # ymdh1 年月日時10碼 起
        # ymdh2 年月日時10碼 迄
        s = """
            SELECT sv02,ps02,ps03,sv03
            FROM rec_sv LEFT JOIN rec_ps ON sv02=ps01
            WHERE
                sv03 >= '{0}'AND
                sv03 <= '{1}'AND
                (sv04 = 1 OR sv04 = 2 OR sv04 = 3 OR sv04 = 6 OR sv04 = 7)
            ORDER BY ps02, sv03 ASC
            """
        s = s.format(ymdh1, ymdh2)
        df = pd.read_sql(s, self.cn) #轉pd
        return df if len(df.index) > 0 else None

    def ymdGersv2_df(self, ymdh1, ymdh2):
        # ymd1 年月日時8碼 起
        # ymd2 年月日時8碼 迄
        s = """
            SELECT sv01,ps02,sv03,sv04
            FROM rec_sv LEFT JOIN rec_ps ON sv02=ps01
            WHERE
                sv03 >= '{0}'AND
                sv03 <= '{1}'AND
                (sv04 = 1 OR sv04 = 2 OR sv04 = 3 OR sv04 = 6 OR sv04 = 7)
            ORDER BY sv03 ASC
            """
        s = s.format(ymdh1, ymdh2)
        df = pd.read_sql(s, self.cn) #轉pd
        return df if len(df.index) > 0 else None

    def wGerpf_df(self, whereSTR=''):
        # 薪資項目明細
        s = """
            SELECT ps02,pa08,pa02,pf05
            FROM rec_pf
                LEFT JOIN rec_ps ON pf01=ps01
                LEFT JOIN rec_pa ON pf02=pa01
            WHERE
                ps11 = 1
            WHEREPLACESTR
            ORDER BY pf02, ps02 ASC
            """
        s = s.replace('WHEREPLACESTR','' if whereSTR =='' else f' AND {whereSTR}')
        df = pd.read_sql(s, self.cn) #轉pd
        return df if len(df.index) > 0 else None

    def Gersw_6_df(self, psid, year): # 人員年度特休以請天數
        # psid : ps01
        # year 4碼
        s = """
            SELECT SUM(sw06) AS TDays
            FROM rec_sw
                LEFT JOIN rec_sg ON rec_sw.sw01 = rec_sg.sg01
            WHERE
                sw02 = {0} AND
                sw04 = 1 AND
                (sg10 = 0 OR sg10 = 1) AND
                LEN(sw05) = 6 AND
                sw05 LIKE '{1}%'
            """
        s = s.format(psid, year)
        df = pd.read_sql(s, self.cn) #轉pd
        return df.iloc[0]['TDays'] if len(df.index) > 0 else 0

    def wGerpa_df(self, whereSTR=''):
        # 基本薪資項目明細
        s = """
            SELECT pa01,pa02,pa03,pa04,pa05,pa06,pa07,pa08,pa09
            FROM rec_pa
            ORDER BY pa08 ASC
            """
        df = pd.read_sql(s, self.cn) #轉pd
        return df if len(df.index) > 0 else None

    def test1(self):
        # 薪資項目明細
        s = "SELECT TOP 1 * FROM rec_ps"
        df = pd.read_sql(s, self.cn) #轉pd
        return df if len(df.index) > 0 else None

    def test(self):
        # s = """
        #     SELECT TOP 100 rd01,rd02,ps02,ps03
        #     FROM rec_rd LEFT JOIN rec_ps ON rd02 = ps01
        #     WHERE ps02 LIKE '%AA0385%'
        #     ORDER BY ps02 ASC
        #     """
        s = """
            SELECT br01,br02,br03
            FROM rec_br
            WHERE br02 LIKE '%202310%' AND br01 = 7
            ORDER BY br02
            """

        df = pd.read_sql(s, self.cn) #轉pd
        print(df)
        # return df if len(df.index) > 0 else None

def test3():
    hr = db_hr()
    # df = hr.userGetrd_df('AA0031', '2025')
    df = hr.userym_Getrd_df('AA0031', '202512')
    pd.set_option('display.max_rows', df.shape[0]+1) # 顯示最多列
    pd.set_option('display.max_columns', None) #顯示最多欄位
    print(df)

def test2(): #添加欄位
    pass
    # hr = db_hr()
    # # 慎重使用
    # s = "ALTER TABLE rec_ps ADD ps56 decimal(12,3), ps57 decimal(12,3)"
    # hr.runsql(s)

    # s = "UPDATE rec_ps SET rp056 = 0, rp057 = 0 WHERE rp01 = 12"
    # s = "UPDATE rec_ps SET ps55 = ''"
    # hr.runsql(s)

def test1():
    # new id
    hr = db_hr()
    df = hr.test()


if __name__ == '__main__':
    test3()
    print('ok')