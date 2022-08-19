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
        s = "SELECT ps01,ps02,ps03,ps11,ps12,ps31,ps34,ps52 FROM rec_ps ORDER BY ps01"
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

    def cpGer_qs_lis(self, qs01): # qs01使用者代號或設備代號 查詢權限 list 
        s = "SELECT qs02 FROM rec_qs WHERE qs01 ='{0}'"
        s = s.format(qs01)
        df = pd.read_sql(s, self.rpt) #轉pd
        return df['qs02'].tolist() if len(df.index) > 0 else []

    def ymGetrd_df(self, ym):
        # rd02 工號
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
    df = hr.ymGetrd_sum_df('202207','202209')
    pd.set_option('display.max_rows', df.shape[0]+1) # 顯示最多列
    pd.set_option('display.max_columns', None) #顯示最多欄位    
    print(df)
    print(df.dtypes)
if __name__ == '__main__':
    test1()        
    print('ok')