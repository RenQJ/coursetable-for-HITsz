import pandas as pd
from datetime import datetime, timedelta
import numpy as np
from ics import Calendar, Event

course = []
'''
解析每一条目，从字符串中分割出有用的信息
'''
def sp(str, week, time):
    arrays = str.split('\n')
    for i in range(0, len(arrays), 3):
        item = []
        item.append(arrays[i])
        item.append(arrays[i+1][1:-1])
        third = arrays[i+2]
        s = third.split('][')
        s[0] = s[0][1:-1]
        s[1] = s[1][:-1]
        item.append(s[0])
        item.append(s[1])
        item.append(week-1)
        item.append(time)
        weeks = s[0].split(',')
        for i in range(len(weeks)):
            interval = weeks[i].split('-')
            if(len(interval) == 1):
                tmp = item.copy()
                tmp[2] = int(interval[0])
                course.append(tmp)
            else:
                for j in range(int(interval[0]), int(interval[-1])+1):
                    tmp = item.copy()
                    tmp[2] = j
                    course.append(tmp)

'''
读表格
'''
df = pd.read_excel('1.xlsx')
for index, row in df.iterrows():
    for col in range(1,8):
        if not pd.isna(row[col]):
            sp(row[col], col, index)

#自定义上课时间，和开学第一天
time = {0:[8,0,9,45],1:[10,0,11,45],2:[14,0,15,45],3:[16,0,17,45],4:[18,45,20,30],5:[20,45,22,30]}
first_day = datetime(2021, 2, 22)
#由周信息转换成日期
for i in range(len(course)):
    item = course[i]
    week = item[2]-1
    day = item[4]
    clas = item[5]
    dtstart = first_day + timedelta(days=week*7+day, hours=time[clas][0]-8, minutes=time[clas][1])#我不知道咋改时区，就这么凑活用吧
    dtend = first_day + timedelta(days=week*7+day, hours=time[clas][2]-8, minutes=time[clas][3])
    item.append(dtstart)
    item.append(dtend)

#添加每节课
c = Calendar()
for i in range(len(course)):
    item = course[i]
    e = Event()
    e.name = item[0]
    e.begin = item[-2].strftime("%Y-%m-%d %H:%M:%S")
    e.end   = item[-1].strftime("%Y-%m-%d %H:%M:%S")
    e.location = item[3]
    e.description = item[1]
    c.events.add(e)
#在每周一标记一下第几周
for i in range(18):
    dtstart = (first_day + timedelta(days=7*i, hours=-8)).strftime("%Y-%m-%d %H:%M:%S")
    dtend = (first_day + timedelta(days=7*i, hours=-7)).strftime("%Y-%m-%d %H:%M:%S")
    e = Event()
    e = Event()
    e.name = "第" + str(i+1) + "周"
    e.begin = dtstart
    e.end = dtend
    c.events.add(e)
with open('course_table.ics', 'w') as my_file:
    my_file.writelines(c)