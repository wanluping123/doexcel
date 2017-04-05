#!/usr/bin/env python
#coding:utf-8
#__author__="ybh"
import xlrd, xlwt
import re
import os
import time,datetime
import settings

from backend.dh_func import daohang_path,dianshang_path,ruanjian_path
def yuming():
    dir=settings.yuming_dir
    files=os.listdir(dir)
    wbk = xlwt.Workbook()
    sheet = wbk.add_sheet(u"域名汇总")
    sheet.write(0, 0, u"时间")
    sheet.write(0, 1, u"域名")
    sheet.write(0, 2, u"IP")
    temp = 1
    for file in files:
        path=os.path.join(dir,file)
        print(path)
        wb = xlrd.open_workbook(path,encoding_override='gb2312')  # 打开文件
        sh = wb.sheet_by_index(0)  # 第一个表

        atime=sh.cell(3,0).value
        if not atime:
            atime=sh.cell(4,0).value

        #time=sh.cell(4,0).value
        atime=re.findall(r"[0-9]{4}-[0-9]{2}-[0-9]{2}",atime)[0]
        if "昨日报表" not in file:
            domainList = sh.col_values(start_rowx=6,colx=0)
            ipList = sh.col_values(start_rowx=6,colx=3)
        else:
            domainList = sh.col_values(start_rowx=6, colx=1)
            ipList = sh.col_values(start_rowx=6, colx=4)

        for i in range(0,len(domainList)-2):
            sheet.write(i+temp,0,atime)
            sheet.write(i+temp,1,domainList[i])
            sheet.write(i+temp,2,ipList[i])
        temp=temp+i+1
    ctime=time.strftime("%Y-%m-%d",time.localtime())
    wbk.save(u"d:\python\域名汇总-%s.xls" % (ctime))

def daohang():
    dir=settings.daohang_dir
    files=os.listdir(dir)
    wbk = xlwt.Workbook()
    sheet = wbk.add_sheet(u"导航汇总")
    sheet.write(0, 0, u"时间")
    sheet.write(0, 1, u"渠道子ID")
    sheet.write(0, 2, u"通过2345浏览器")
    sheet.write(0, 3, u"通过其他浏览器")
    temp = 1
    style1 = xlwt.XFStyle()
    style1.num_format_str = 'YYYY-MM-DD'
    for file in files:
        path = os.path.join(dir, file)
        print(path)
        try:
            wb = xlrd.open_workbook(path, encoding_override='gb2312')  # 打开文件
        except Exception as e:
            newpath=daohang_path(path)
            wb = xlrd.open_workbook(newpath, encoding_override='gb2312')
        sh = wb.sheet_by_index(0)  # 第一个表

        #time=sh.cell(4,0).value
        check=sh.cell(0,1).value
        timeList = sh.col_values(start_rowx=1, colx=0)
        #print(type(timeList[0]))

        if check=="通过2345浏览器":
            try:
                timeList = [xlrd.xldate.xldate_as_datetime(x, 0) for x in timeList[:-1]]
            except Exception as e:
                timeList = [time.strptime(x, "%Y-%m-%d") for x in timeList[:-1]]
                timeList = [datetime.datetime(*x[:3]) for x in timeList]
            browser1List = sh.col_values(start_rowx=1,colx=1)
            browser2List = sh.col_values(start_rowx=1, colx=3)
            qid=re.findall(r"[0-9]{5}",file)
            for i in range(0,len(timeList)):
                sheet.write(i+temp,0,timeList[i].strftime('%Y-%m-%d'))
                sheet.write(i+temp,1,qid)
                sheet.write(i+temp,2,browser1List[i])
                sheet.write(i+temp,3,browser2List[i])
            temp=temp+i+1
        if check=="子账户号":
            try:
                timeList = [xlrd.xldate.xldate_as_datetime(x, 0) for x in timeList]
            except Exception as e:
                timeList = [time.strptime(x, "%Y-%m-%d") for x in timeList]
                timeList = [datetime.datetime(*x[:3]) for x in timeList]
            browser1List = sh.col_values(start_rowx=1, colx=2)
            qidList = sh.col_values(start_rowx=1, colx=1)
            for i in range(0,len(timeList)):
                sheet.write(i+temp,0,timeList[i].strftime('%Y-%m-%d'))
                sheet.write(i+temp,1,qidList[i])
                sheet.write(i+temp,2,browser1List[i])
                sheet.write(i+temp,3,"")
            temp=temp+i+1
        if check=="渠道代码":
            try:
                timeList = [xlrd.xldate.xldate_as_datetime(x, 0) for x in timeList[:-1]]
            except Exception as e:
                timeList = [time.strptime(x, "%Y-%m-%d") for x in timeList[:-1]]
                timeList = [datetime.datetime(*x[:3]) for x in timeList]
            browser2List = sh.col_values(start_rowx=1, colx=2)
            qidList = sh.col_values(start_rowx=1, colx=1)
            for i in range(0,len(timeList)):
                sheet.write(i+temp,0,timeList[i].strftime('%Y-%m-%d'))
                sheet.write(i+temp,1,re.findall(r"[0-9]{5}",qidList[i])[0])
                sheet.write(i+temp,2,"")
                sheet.write(i+temp,3,browser2List[i])
            temp=temp+i+1


    ctime=time.strftime("%Y-%m-%d",time.localtime())
    wbk.save(u"d:\python\导航汇总-%s.xls" % (ctime))

def dianshang():
    dir=settings.dianshang_dir
    files=os.listdir(dir)
    wbk = xlwt.Workbook()
    sheet = wbk.add_sheet(u"电商汇总")
    sheet.write(0, 0, u"时间")
    sheet.write(0, 1, u"渠道子ID")
    sheet.write(0, 2, u"预估收入")
    temp = 1
    style1 = xlwt.XFStyle()
    style1.num_format_str = 'YYYY-MM-DD'
    for file in files:
        path = os.path.join(dir, file)
        print(path)
        try:
            wb = xlrd.open_workbook(path, encoding_override='gb2312')  # 打开文件
        except Exception as e:
            path=dianshang_path(path)
            wb = xlrd.open_workbook(path, encoding_override='gb2312')
        sh = wb.sheet_by_index(0)  # 第一个表

        #time=sh.cell(4,0).value
        check=sh.cell(0,1).value
        qid = re.findall(r"[0-9]{5}", file)

        if check=="广告方案":
            timeList = sh.col_values(start_rowx=1, colx=11)
            countList = sh.col_values(start_rowx=1, colx=5)
            #try:
             #   timeList = [xlrd.xldate.xldate_as_datetime(x, 0) for x in timeList[:-1]]
            #except Exception as e:
            timeList = [time.strptime(x, "%Y-%m-%d %H:%M:%S") for x in timeList]
            timeList = [datetime.datetime(*x[:3]) for x in timeList]
            for i in range(0,len(timeList)):
                sheet.write(i+temp,0,timeList[i].strftime('%Y-%m-%d'))
                sheet.write(i+temp,1,qid)
                sheet.write(i+temp,2,countList[i])
            temp=temp+i+1
        if check=="平台-设备":
            timeList = sh.col_values(start_rowx=1, colx=0)
            countList = sh.col_values(start_rowx=1, colx=3)
            #try:
             #   timeList = [xlrd.xldate.xldate_as_datetime(x, 0) for x in timeList[:-1]]
            #except Exception as e:
            timeList = [time.strptime(x, "%Y%m%d") for x in timeList]
            timeList = [datetime.datetime(*x[:3]) for x in timeList]
            for i in range(0, len(timeList)):
                sheet.write(i + temp, 0, timeList[i].strftime('%Y-%m-%d'))
                sheet.write(i + temp, 1, qid)
                sheet.write(i + temp, 2, countList[i])
            temp=temp+i+1
        if check=="推广位名称":
            timeList = sh.col_values(start_rowx=1, colx=0)
            countList = sh.col_values(start_rowx=1, colx=7)
            #try:
            #    timeList = [xlrd.xldate.xldate_as_datetime(x, 0) for x in timeList[:-1]]
            #except Exception as e:
            timeList = [time.strptime(x, "%Y%m%d ") for x in timeList]
            timeList = [datetime.datetime(*x[:3]) for x in timeList]
            for i in range(0, len(timeList)):
                sheet.write(i + temp, 0, timeList[i].strftime('%Y-%m-%d'))
                sheet.write(i + temp, 1, qid)
                sheet.write(i + temp, 2, countList[i])
            temp=temp+i+1


    ctime=time.strftime("%Y-%m-%d",time.localtime())
    wbk.save(u"d:\python\电商汇总-%s.xls" % (ctime))


def ruanjian():
    dir=settings.ruanjian_dir
    files = os.listdir(dir)
    wbk = xlwt.Workbook()
    sheet = wbk.add_sheet(u"软件汇总")
    sheet.write(0, 0, u"时间")
    sheet.write(0, 1, u"渠道子ID")
    sheet.write(0, 2, u"安装量")
    temp = 1
    style1 = xlwt.XFStyle()
    style1.num_format_str = 'YYYY-MM-DD'
    for file in files:
        path = os.path.join(dir, file)
        print(path)
        try:
            wb = xlrd.open_workbook(path, encoding_override='gb2312')  # 打开文件
        except Exception as e:
            path = ruanjian_path(path)
            wb = xlrd.open_workbook(path, encoding_override='gb2312')
        sh = wb.sheet_by_index(0)  # 第一个表

        # time=sh.cell(4,0).value
        check = sh.cell(0, 0).value


        if check == "账号":
            timeList = sh.col_values(start_rowx=2, colx=1)
            countList = sh.col_values(start_rowx=2, colx=2)
            try:
                timeList = [xlrd.xldate.xldate_as_datetime(x, 0) for x in timeList]
            except Exception as e:
                timeList = [time.strptime(x, "%Y-%m-%d") for x in timeList]
                timeList = [datetime.datetime(*x[:3]) for x in timeList]
            qid=re.findall(r'[0-9]{5}',file)
            for i in range(0, len(timeList)):
                sheet.write(i + temp, 0, timeList[i].strftime('%Y-%m-%d'))
                sheet.write(i + temp, 1, qid)
                sheet.write(i + temp, 2, countList[i])
            temp = temp + i + 1
        if check == "SoftStat":
            timeList = sh.col_values(start_rowx=2, colx=1)
            #print(timeList)
            countList = sh.col_values(start_rowx=2, colx=2)
            countList=[ x.strip() for x in countList ]
            qid = re.findall(r'[0-9]{5}', file)
            try:
                #timeList = [xlrd.xldate.xldate_as_datetime(x, 0) for x in timeList[:-1]]
                timeList = [x.strip() for x in timeList[:-1]]
                #print(timeList)
                #timeList = [time.strptime(x.strip('\r\n').strip(" "), "%Y-%m-%d") for x in timeList]

                #timeList = [datetime.datetime(*x[:3]) for x in timeList]
            except Exception as e:
                print(e)
            for i in range(0, len(timeList)):
                sheet.write(i + temp, 0, timeList[i])
                sheet.write(i + temp, 1, qid)
                sheet.write(i + temp, 2, countList[i])
            temp = temp + i + 1
        if check == "序号":
            timeList = sh.col_values(start_rowx=1, colx=1)
            countList = sh.col_values(start_rowx=1, colx=5)
            qidList= sh.col_values(start_rowx=1, colx=4)
            try:
                timeList = [xlrd.xldate.xldate_as_datetime(x, 0) for x in timeList]
            except Exception as e:
                timeList = [time.strptime(x, "%Y-%m-%d") for x in timeList]
                timeList = [datetime.datetime(*x[:3]) for x in timeList]
            for i in range(0, len(timeList)):
                sheet.write(i + temp, 0, timeList[i].strftime('%Y-%m-%d'))
                sheet.write(i + temp, 1, qidList[i])
                sheet.write(i + temp, 2, countList[i])
            temp = temp + i + 1
        if check == "日期":
            timeList = sh.col_values(start_rowx=1, colx=0)
            countList = sh.col_values(start_rowx=1, colx=1)
            qid = re.findall(r'[0-9]{5}', file)
            try:
                timeList = [xlrd.xldate.xldate_as_datetime(x, 0) for x in timeList]
            except Exception as e:
                timeList = [time.strptime(x, "%Y-%m-%d") for x in timeList]
                timeList = [datetime.datetime(*x[:3]) for x in timeList]
            for i in range(0, len(timeList)):
                sheet.write(i + temp, 0, timeList[i].strftime('%Y-%m-%d'))
                sheet.write(i + temp, 1, qid)
                sheet.write(i + temp, 2, countList[i])
            temp = temp + i + 1

    ctime = time.strftime("%Y-%m-%d", time.localtime())
    wbk.save(u"d:\python\软件汇总-%s.xls" % (ctime))

if __name__=="__main__":
    yuming()
    #daohang()
    #dianshang()
    #ruanjian()