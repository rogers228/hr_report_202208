if True:
    import sys, custom_path
    config_path = custom_path.custom_path['hr_report_202208'] # 取得專案引用路徑
    sys.path.append(config_path) # 載入專案路徑

import os, time
import click
import tool_auth
import rpt_sav03, rpt_sav04, rpt_sav05, rpt_sav06, rpt_sav07, rpt_sav08

@click.command() # 命令行入口
@click.option('-report_name', help='report name', required=True, type=str) # required 必要的
@click.option('-userno', help='user no Like AA0031', type=str)
@click.option('-y1', help='year 4 char', type=str)
@click.option('-ym', help='year and month 6 char', type=str)
@click.option('-ym1', help='year and month 6 char', type=str)
@click.option('-ym2', help='year and month 6 char', type=str)
@click.option('-ymd', help='year, month, day 8 char', type=str)
@click.option('-h1', help='hour 2 char', type=str)
@click.option('-h2', help='hour 2 char', type=str)
def main(report_name, userno='',
        y1='', ym='', ym1='', ym2='', ymd='',
        h1='', h2=''):
    au = tool_auth.Authorization()
    if not au.isqs(701): # 檢查 701 權限
        return # 無權限 退出

    global usernoStr; usernoStr = userno
    global y1Str;     y1Str = y1
    global ymStr;     ymStr = ym
    global ym1Str;    ym1Str = ym1
    global ym2Str;    ym2Str = ym2
    global ymdStr;    ymdStr = ymd
    global h1Str;     h1Str = h1
    global h2Str;     h2Str = h2
    global fileName; fileName = report_name + '_' + time.strftime("%Y%m%d%H%M%S", time.localtime()) + '.xlsx'
    dic = {'sav03': sav03,
           'sav04': sav04,
           'sav05': sav05,
           'sav06': sav06,
           'sav07': sav07,
           'sav08': sav08,
          }

    func = dic.get(report_name, None)
    if func is not None:
        func()

def sav03(): # 出勤統計表
    rpt_sav03.Report_sav03(fileName, ym1Str, ym2Str)

def sav04(): # 薪資轉帳明細表
    rpt_sav04.Report_sav04(fileName, ymStr)

def sav05(): # 出勤狀況表(月)
    rpt_sav05.Report_sav05(fileName, ymStr)

def sav06(): # 出勤狀況表(個人年度)
    rpt_sav06.Report_sav06(fileName, usernoStr, y1Str)

def sav07(): # 職務代理人清冊
    rpt_sav07.Report_sav07(fileName)

def sav08(): # 出勤狀況表(當日)  看訂便當人數
    rpt_sav08.Report_sav08(fileName, ymdStr, h1Str, h2Str)

if __name__ == '__main__':
    # main('sav07') # debug
    main()
    # cmd
    # C:\python_venv\python.exe \\220.168.100.104\pdm\python_program\hr_report_202208\rpt_main.py -report_name sav07
