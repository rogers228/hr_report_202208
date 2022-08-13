import os, time
import tool_db_hr
import PySimpleGUI as sg
import rpt_sav07

sg.theme('SystemDefault')
class Gui_sav07():
    def __init__(self):
        self.qsno = 701 #權限
        self.gui_name = '職務代理人清冊'
        self.report_name = 'sav07'
        self.hr = tool_db_hr.db_hr()
        self.check_qs()
        self.gui()

    def check_qs(self): # 檢查權限
        computer_name = os.getenv('COMPUTERNAME')
        if computer_name in ['VM-TESTER']:
            return # 開發環境

        no = self.hr.ps18Getps01(computer_name) #電腦名稱 搜尋使用者
        if no == '':
            no = self.hr.pc02Getpc01(computer_name) # 設備名稱搜尋使用者
        lis_qs = self.hr.cpGer_qs_lis(no) # 全限列表
        # print(lis_qs)
        if self.qsno not in lis_qs:
            sg.popup('尚未設定權限!\n\n請洽系統管理員\n')
            raise SystemExit  #正式結束程式 而不需要導入sys

    def gui(self):
        layout = [  [sg.Text(self.gui_name + '\n\n')],
                    [sg.Button('ok'), sg.Button('Cancel')] ]
        w = sg.Window('Yeoshe HR', layout, 
                        size=(350, 120),
                        resizable=True)
        while True:
            event, values = w.read()
            if event == sg.WIN_CLOSED or event == 'Cancel': # if user closes window or clicks cancel
                break

            if event == 'ok':
                fileName = self.report_name + '_' + time.strftime("%Y%m%d%H%M%S", time.localtime()) + '.xlsx'
                rpt_sav07.Report_sav07(fileName) #製表後開啟
                break

def test1():
    win = Gui_sav07()

if __name__ == '__main__':
    test1()