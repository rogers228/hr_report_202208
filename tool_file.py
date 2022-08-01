import os
from config import *

class File_tool():
    def __init__(self):
        self.report_path = config_report_path.get(os.getenv('COMPUTERNAME'), None)
        if self.report_path is None:
            print('找不到路徑')
            raise SystemExit  #結束

    def clear(self):
        for filename in os.listdir(self.report_path):
            try:
                os.remove(os.path.join(self.report_path, filename))
            except Exception as e:
                print('Failed to delete %s. Reason: %s' % (filename, e))

def test1():
    ftl = File_tool()
    ftl.clear()
if __name__ == '__main__':
    test1()



