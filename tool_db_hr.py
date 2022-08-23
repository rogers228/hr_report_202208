if True:
    import sys, custom_path
    config_path = custom_path.custom_path['hr_report_202208'] # 取得專案引用路徑
    sys.path.append(config_path) # 載入專案路徑
    
import pandas as pd
import pyodbc
from config import *

class db_hr(): #讀取excel 單一零件
    def __init__(self):
        self.cn = pyodbc.connect(config_conn_HR) # connect str 連接字串
        self.rpt = pyodbc.connect(config_conn_RPT) # connect str 連接字串
        self.dbps = self.get_database_ps() # 建議一次性基本資料檔，避免多次存取db

    def runsql(self, SQL):
        try:
            cur = self.cn.cursor()
            cur.execute(SQL) #執行
            cur.commit() #更新
            cur.close() #關閉
        except:
            print(SQL)
            logging.warning('error class db_ab().def runsql()! 無法執行SQL!')

    def runsql_rpt(self, SQL):
        try:
            cur = self.rpt.cursor()
            cur.execute(SQL) #執行
            cur.commit() #更新
            cur.close() #關閉
        except:
            print(SQL)
            logging.warning('error class db_ab().def runsql()! 無法執行SQL!')

    def get_database_ps(self):
        s = "SELECT ps01,ps02,ps03,ps11,ps12,ps31,ps32,ps33,ps34,ps52 FROM rec_ps ORDER BY ps01"
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

    def df_ps_atwork(self): #在職人員列表
        ps = self.dbps
        df = ps.loc[ps['ps11'] == 1] # 篩選
        df = df.fillna('') # 填充NaN為空白
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

    def wGetps_df(self, whereSTR=''):
        # whereSTR: 不包含 WHERE的 where SQL語句
        s = """
            SELECT ps02,ps03,ps05,ps06,ps07,ps08,ps09,ps10,ps11,
                    ps12,ps13,ps14,ps22,ps23,ps25,ps26,ps27,ps28,ps29,
                    ps30,ps31,ps34,ps35,ps52,ps53,
                    bn02
            FROM rec_ps
            LEFT JOIN rec_bn ON ps31=bn01
            WHEREPLACESTR
            ORDER BY ps02 ASC
            """
        s = s.replace('WHEREPLACESTR','' if whereSTR =='' else f'WHERE {whereSTR}')
        df = pd.read_sql(s, self.cn) #轉pd
        return df if len(df.index) > 0 else None

    def ymGetrd_df(self, ym):
        # ym 年月日6碼
        s = """
            SELECT rd02,rd03,rd06,rd07,rd08,rd10,rd11,rd12,rd19,
                    rd21,rd22,rd23,rd24,rd25,rd26,rd27,rd28,rd29,rd30,rd31,rd34,rd35
            FROM rec_rd
            WHERE rd03 LIKE '{0}%'
            ORDER BY rd03 ASC
            """
        s = s.format(ym)
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
            SELECT ps02,ps03,ps40,rs02,rs08,rs10,rs11,rs12
            FROM rec_ps 
                LEFT JOIN rec_rs ON ps01=rs02
            WHERE rs03 LIKE '{0}%'
            ORDER BY ps02 ASC
            """
        s = s.format(ym)
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
                Sum(rd35) AS Srd35
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

    def test(self):
        s = "SELECT TOP 5 * FROM rec_ps"
        # s= "SELECT ps01,ps02,ps03 FROM rec_ps"
        # s ="""SELECT rd01,rd02,rd03  FROM rec_rd
        #     WHERE
        #         rd02 = 32 AND
        #         rd03 LIKE '202006%'
        #     ORDER BY rd03"""

        # s = """SELECT COUNT(*) FROM rec_sv
        #     WHERE
        #         sv02 = 219 AND
        #         sv03 LIKE '202109%' AND
        #         sv04 = 1
        #         """
        print(s)
        df = pd.read_sql(s, self.cn) #轉pd
        pd.set_option('display.max_rows', df.shape[0]+1) # 顯示最多列
        pd.set_option('display.max_columns', None) #顯示最多欄位
        print(df)

def test2(): #添加欄位
    pass
    hr = db_hr()
    # # 慎重使用
    # s = "ALTER TABLE rec_ps ADD ps53 text"
    # hr.runsql(s)

    # rpt
    s = "UPDATE rec_rpt SET rp07 = 0 WHERE rp01 = 12"
    hr.runsql_rpt(s)

def test1():
    # new id
    hr = db_hr()
    # df = hr.ymGetrs_df('202207')
    # pd.set_option('display.max_rows', df.shape[0]+1) # 顯示最多列
    # pd.set_option('display.max_columns', None) #顯示最多欄位    
    # print(df)

    print(hr.pa08Getpa02('A002'))
    # whereSTR = "pa08 IN ('0010','0020','0030','0040','0050','0060','0070','0210','0220','A001')"
    # df_rd = hr.wGerpf_df(whereSTR)
    # print(df_rd)

if __name__ == '__main__':
    test1()        
    print('ok')