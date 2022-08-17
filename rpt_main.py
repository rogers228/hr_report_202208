if True:
    import sys, custom_path
    config_path = custom_path.custom_path['hr_report_202208'] # 取得專案引用路徑
    sys.path.append(config_path) # 載入專案路徑

import os, time
import click
import tool_auth
import rpt_sav07

@click.command() # 命令行入口
@click.option('-report_name', help='report name', required=True, type=str) # required 必要的
def main(report_name):
    au = tool_auth.Authorization()
    if not au.isqs(701): # 檢查 701 權限
        return # 無權限 退出

    dic = {'sav07': sav07}
    func = dic.get(report_name, None)
    if func is not None:
        func()

def sav07():
    # print('sav07')
    report_name = 'sav07'
    fileName = report_name + '_' + time.strftime("%Y%m%d%H%M%S", time.localtime()) + '.xlsx'
    rpt_sav07.Report_sav07(fileName) #製表後開啟

if __name__ == '__main__':
    # main('sav07') # debug
    main()
    # cmd
    # C:\python_venv\python.exe \\220.168.100.104\pdm\python_program\hr_report_202208\rpt_main.py -report_name sav07
