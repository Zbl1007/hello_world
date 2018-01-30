#导包

import sys
import calendar
import importlib
import time, datetime
from datetime import datetime,date
import json
importlib.reload(sys)


#每周第一天为星期天
calendar.setfirstweekday(6)

#保存为 .ics 文件
def save(string):
     f = open("class.ics", 'wb')
     f.write(string.encode("utf-8"))
     f.close()
     print("课表保存成功!")


#.ics文件头
icsTitle = "BEGIN:VCALENDAR\n"+"PRODID:-//Microsoft Corporation//Outlook 15.0 MIMEDIR//EN\n"+"VERSION:2.0\n"+"X-WR-CALNAME:2017年夏秋季学期\n"+"BEGIN:VTIMEZONE\n"+"TZID:China Standard Time\n"+"BEGIN:STANDARD\n"+"DTSTART:16010101T000000\n"+"TZOFFSETFROM:+0800\n"+"TZOFFSETTO:+0800\n"+"END:STANDARD\n"+"END:VTIMEZONE\n"
BEGIN = "BEGIN:VEVENT\n"
END = "END:VEVENT\n"
DTSTART = "DTSTART;TZID=\"China Standard Time\":"
DTEND = "DTEND;TZID=\"China Standard Time\":"
RRULE = "RRULE:FREQ=WEEKLY;COUNT="
SUMMARY = "SUMMARY;LANGUAGE=zh-cn:"
END_END = "END:VCALENDAR\n"

#计算课程开始和结束时间
START_TIME_OF_CLASS = ["T082000\n","T102000\n","T143000\n","T163000\n","T193000\n"]
END_TIME_OF_CLASS = ["T100000\n","T120000\n","T161000\n","T181000\n","T211000\n"]



start_day = input("请设置第一周的星期一日期(如：20160905):")
start_day = time.strptime(start_day,'%Y%m%d')
#开学那天为几号星期几
wday = [1,2,3,4,5,6,0]
start_weekday = wday[start_day.tm_wday]

start_day = time.mktime(start_day)

#打开json文件
global data
with open('classInfo.json','r') as f:
    data = json.load(f)
    #字符串转字典
    data = eval(data)
 
classInfo = data['classInfo']

#that_day = start_day + ((class_start_week - 1) * 7 + f[:weekday] - start_weekday) * 24 * 3600

 #写ics文件
global icsString
icsString = ""
icsString +=icsTitle

#记录第几节课
classTime = 0
WEEK = ''
for i in classInfo:
    #单双周的记录
    Flag=0
    for j in i:
        weekday = int(j['weekday'])
        week = j['week'].split('.')
        for k in week:
            icsString +=BEGIN
            #上课周不连续
            flag = k.find('-')
            if flag>0:
                week = k.split('-')
                WEEK = week[0]
                COUNT = int(week[1])-int(week[0])+1
                that_day = start_day + ((int(week[0])-1)*7 + weekday - start_weekday)*24*3600

            else:
                COUNT = 1
                WEEK = week[Flag]
                that_day = start_day + ((int(week[Flag])-1)*7 + weekday - start_weekday)*24*3600
                Flag +=1
            
            #计算开课日期
            
            that_day = time.localtime(that_day)
            that_day = time.strftime("%Y%m%d",that_day)
            
            
            icsString +=DTEND + that_day + START_TIME_OF_CLASS[classTime]
            icsString +=DTSTART + that_day + END_TIME_OF_CLASS[classTime]
            icsString +=RRULE+str(COUNT)+'\n'
            icsString +=SUMMARY+j['SUMMARY']+'\n'
            icsString +=END+'\n'
        
    #课程开始和结束时间
    classTime +=1

icsString +=END_END

print("课表写入成功！")
save(icsString)



