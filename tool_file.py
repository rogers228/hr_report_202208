if True:
    import sys, custom_path
    config_path = custom_path.custom_path['hr_report_202208'] # 取得專案引用路徑
    sys.path.append(config_path) # 載入專案路徑
    
import os
import config

class File_tool():
    def __init__(self):
        self.init_cs()  # 初始化

    def init_cs(self):
        # 初始化
        self.report_dir = config.config_report_dir # 資料夾名稱
        self.report_path = os.path.join(os.path.expanduser(r'~\Documents'), self.report_dir) #資料夾路徑
        if not os.path.isdir(self.report_path): #建立資料夾
            os.mkdir(self.report_path)

    def clear(self, key): # 清除特定報表
        for f in os.listdir(self.report_path):
            if os.path.isfile(os.path.join(self.report_path, f)): # 僅針對檔案
                if f.find(key) == 0: # 該檔案是否為key開頭
                    try:
                        os.remove(os.path.join(self.report_path, f))
                    except:
                        pass

def test1():
    ftl = File_tool()
    ftl.clear('sav07')
    # ftl.ini_write('sav07','456')
    # print(ftl.ini_get('sav07'))
    # ftl.clear()
if __name__ == '__main__':
    test1()



