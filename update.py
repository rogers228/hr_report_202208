# update to server
from config import *
import dirsync

def main():
    # 開發資料夾 同步至server 以利工作站執行
    dirsync.sync(config_develop_program, config_servr_program, action='sync')
    print('finish')

if __name__ == '__main__':
    main()