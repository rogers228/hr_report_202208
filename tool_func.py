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

def getNextymStr(ym): #下一個月的年月6碼
    # ym 年月6碼
    y = int(ym[:4])
    m = int(ym[4:6])
    if m < 12:
        m += 1
    else:
        m =1
        y += 1

    return '{:0>4d}'.format(y) + '{:0>2d}'.format(m)


def test1():
    print(getNextymStr('202201'))


if __name__ == '__main__':
    test1()
