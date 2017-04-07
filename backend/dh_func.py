#!/usr/bin/env python
#coding:utf-8
#__author__="ybh"
import xlwt
import os,time
#file=open("D:\余斌宏\数据模板\导航\\2345导航33883子渠道导出.xls",'r')
#result=file.readlines()
#for i in range(len(result)-1):

#    print(result[i])


def daohang_path(path="D:\余斌宏\数据模板\导航\\2345导航33883子渠道导出.xls"):
    file=open(path,'r')
    result=file.readlines()
    file.close()
    wbk = xlwt.Workbook()
    sheet = wbk.add_sheet("sheet1")
    for i in range(len(result)):
        list=result[i].strip('\n').split()
        #print(list)

        for y in range(len(list)):
            sheet.write(i,y,list[y])
            #print(list[y])
    file_name=os.path.basename(path)
    dir_name=os.path.dirname(path)
    newpath="%s%snew%s" %(dir_name,os.sep,file_name)
    wbk.save(u"%s" % (newpath))
    return newpath


def dianshang_path(path='D:\python\data\dianshang\唯品会2月数据下载.csv'):
    try:
        f = open(path,'r')
        lines = f.readlines()
        f.close()
        wbk = xlwt.Workbook()
        sheet = wbk.add_sheet("sheet1")
        for i in range(len(lines)):
            list = lines[i].strip('\n').split(',')
            for y in range(len(list)):
                sheet.write(i, y, list[y].strip("\"").strip("￥"))
    except Exception as e:
        f = open(path, 'r', encoding='utf_16_le')
        lines=f.readlines()
        f.close()
        wbk = xlwt.Workbook()
        sheet = wbk.add_sheet("sheet1")
        for i in range(len(lines)):
            list=lines[i].strip().split('\t')
            for y in range(len(list)):
                sheet.write(i, y, list[y].strip("\ufeff"))
                #print(list[y].strip("\"").strip("￥"))
    file_name = os.path.basename(path)
    file_name = file_name.split('.')[0]
    dir_name = os.path.dirname(path)
    newpath = "%s%snew%s.xlsx" % (dir_name, os.sep, file_name)
    wbk.save(u"%s" % (newpath))
    return newpath

def ruanjian_path(path='D:\python\data\ruanjian\东方输入法多日.csv'):
    f = open(path, 'r',encoding='utf-8')

    lines = f.readlines()
    f.close()
    wbk = xlwt.Workbook()
    sheet = wbk.add_sheet("sheet1")
    for i in range(len(lines)):
        list = lines[i].strip('\n').split(',')
        for y in range(len(list)):
            sheet.write(i, y, list[y].strip("\"").strip("\r\n").strip(" "))
            # print(list[y].strip("\"").strip("￥"))
    file_name = os.path.basename(path)
    file_name = file_name.split('.')[0]
    dir_name = os.path.dirname(path)
    newpath = "%s%snew%s.xlsx" % (dir_name, os.sep, file_name)
    wbk.save(u"%s" % (newpath))
    return newpath

if __name__=="__main__":
    #get_newpath()
    dianshang_path()

