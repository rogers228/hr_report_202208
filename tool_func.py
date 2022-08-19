import datetime

def getYMdays(year_s, month_s): #該月天數
    try:
        mydate = datetime.datetime(int(year_s), int(month_s), 1)
        old_Month = mydate.month
        new_Month = mydate.month

        i = 0
        while new_Month == old_Month:
            mydate = mydate + datetime.timedelta(days = 1)
            new_Month = mydate.month
            i = i + 1
            if i > 31:
                return 31
        return i
    except:
        return 31        

def getWeekdayStr(year_s, month_s, day_s): #計算星期幾
    try:
        mydate = datetime.datetime(int(year_s), int(month_s), int(day_s))
        index = mydate.weekday()
        wStr = ('一','二','三','四','五','六','日')
        return wStr[index]
    except:
        return ''

def test1():
    print(getYMdays('2022','02'))
    print(getWeekdayStr('2022','08','18'))

if __name__ == '__main__':
    test1()
