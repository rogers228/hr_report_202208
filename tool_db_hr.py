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
        s = "SELECT ps01,ps02,ps03,ps11,ps12,ps31,ps52 FROM rec_ps ORDER BY ps01"
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
    # no = hr.ps18Getps01('DESKTOP-CFANU1B')
    # qs = hr.cpGer_qs_lis(no)
    # print(no)
    # print(qs)

    no = hr.pc02Getpc01('DESKTOP-0LGQBL4')
    print(no)
    qs = hr.cpGer_qs_lis(no)
    print(qs)
if __name__ == '__main__':
    test2()        
    print('ok')