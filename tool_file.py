import os
# import configparser
import config

class File_tool():
    def __init__(self):
        self.init_cs()  # 初始化

    def init_cs(self):
        # 初始化
        # 在電腦文件夾建立資料夾及ini檔
        # ini紀錄當下製表紀錄，作為下次製表時的清除依據

        self.report_dir = config.config_report_dir # 資料夾名稱
        self.report_path = os.path.join(os.path.expanduser(r'~\Documents'), self.report_dir) #資料夾路徑
        self.ini = 'report.ini' # ini檔案名稱
        self.ini_path = os.path.join(self.report_path, self.ini) # ini路徑

        if not os.path.isdir(self.report_path): #建立資料夾
            os.mkdir(self.report_path)

            # # 建立ini
            # cf = configparser.ConfigParser()
            # cf['report'] = {}
            # with open(self.ini_path, 'w') as f:
            #     cf.write(f)

    # def ini_write(self, key, value): #寫入ini
    #     cf = configparser.ConfigParser()
    #     cf.read(self.ini_path)
    #     cf.set('report', key, str(value))
    #     with open(self.ini_path, 'w') as f:
    #         cf.write(f)

    # def ini_get(self, key): #寫入ini
    #     result = ''
    #     cf = configparser.ConfigParser()
    #     cf.read(self.ini_path)
    #     try:
    #         result = cf['report'][key]
    #     except:
    #         pass
    #     return result

    # def clear(self, key): # 清除特定報表
    #     # 製表sav07 前 清除上一個sav07
    #     old_file = os.path.join(self.report_path, self.ini_get(key))
    #     if os.path.exists(old_file): #若檔案存在則刪除
    #         try:
    #             os.remove(old_file)
    #         except Exception as e:
    #             print('Failed to delete %s. Reason: %s' % (old_file, e))

    def clear(self, key): # 清除特定報表
        for f in os.listdir(self.report_path):
            if os.path.isfile(os.path.join(self.report_path, f)): # 僅針對檔案
                if f.find(key) == 0: # 該檔案是否為key開頭
                    try:
                        os.remove(os.path.join(self.report_path, f))
                    except Exception as e:
                        print('Failed to delete %s. Reason: %s' % (os.path.join(self.report_path, e), e))

def test1():
    ftl = File_tool()
    ftl.clear('sav07')
    # ftl.ini_write('sav07','456')
    # print(ftl.ini_get('sav07'))
    # ftl.clear()
if __name__ == '__main__':
    test1()



