import os, sys
sys.path.append(os.path.dirname(__file__)) #加入資料夾路徑   使用絕對路徑  cmd執行才會正確

import click
import gui_sav07

@click.command() # 命令行入口
@click.option('-name', help='report name', required=True) # required 必要的
def main(name):
    dic = {'sav07': sav07}
    func = dic.get(name, None)
    if func is not None:
        func()

def sav07():
    gui_sav07.Gui_sav07()

if __name__ == '__main__':
    main('sav07') # debug
    # main()
