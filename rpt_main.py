if True:
    import sys, custom_path
    config_path = custom_path.custom_path['hr_report_202208'] # 取得專案引用路徑
    sys.path.append(config_path) # 載入專案路徑

import os, time
import click
import tool_auth
import rpt_sav05, rpt_sav07

@click.command() # 命令行入口
@click.option('-report_name', help='report name', required=True, type=str) # required 必要的
@click.option('-ym', help='year and month 6 integer', type=str) # required 必要的
def main(report_name, ym = ''):
    au = tool_auth.Authorization()
    if not au.isqs(701): # 檢查 701 權限
        return # 無權限 退出

    global ymStr;    ymStr = ym
    global fileName; fileName = report_name + '_' + time.strftime("%Y%m%d%H%M%S", time.localtime()) + '.xlsx'
    dic = {'sav05': sav05,
           'sav07': sav07
          }

    func = dic.get(report_name, None)
    if func is not None:
        func()

def sav05(): # 出勤狀況表(月)
    rpt_sav05.Report_sav05(fileName, ymStr) #製表後開啟

def sav07(): # 職務代理人清冊
    rpt_sav07.Report_sav07(fileName) #製表後開啟

if __name__ == '__main__':
    # main('sav07') # debug
    main()
    # cmd
    # C:\python_venv\python.exe \\220.168.100.104\pdm\python_program\hr_report_202208\rpt_main.py -report_name sav07
