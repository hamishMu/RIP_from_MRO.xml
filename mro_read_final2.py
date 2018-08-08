#!user/bin/env python3
#_*_ coding :utf-8 _*_
import os
import gzip
import time
from lxml import etree
from os.path import join, splitext
import xlwt
import xlrd
import sys
import csv
import glob
def read_MRO(SourceFile):
    f = glob.iglob(SourceFile+"\*.xml")
    names =[]
    for x in f:
        names.append(x)
    #name = SourceFile.split('.')
    lst = []
    i = 0	#rowcount
    start_time = time.time()
    for i in range(len(names)):
        root = etree.parse(names[i])
        name = names[i].split('.')
        objs = root.xpath('/bulkPmMrDataFile/eNB/measurement[last()]/object')
        lst.extend('eNb')
        lst.extend(',')
        lst.extend('id')
        lst.extend(',')
        lst.extend('rip_value')
        lst.append('\n')  
        for obj in objs:
            for v in obj.xpath('./v/node()'):
                i += 1
                if i%10000 == 0:
                    print('------>%d' % i)  
            lst.extend(root.xpath('/bulkPmMrDataFile/eNB/@id'))
            lst.extend(',')
            id_name =obj.xpath('@id')
            id_v = id_name[0].split(':')[0]
            #print(type(id_name))
	    #lst_id=id_name.split(':')
	    #lst.extend(lst_id[0])
            #lst.extend(obj.xpath('@id'))
            lst.extend(id_v)
            lst.extend(',')
            lst.extend(obj.xpath('@v'))
            lst.append(v+'\n')
        #生成csv文件以写入数据
        #print(len(lst))
        with open(name[0]+'.csv','w') as f:
            #t.writelines(a)
            #写入解析后内容
            f.writelines(lst)
        print('文件行计数：%d，处理用时：%f.' % (i,time.time()-start_time))
        lst.clear()
      
    ENBs,IDs,Values =[],[],[]
    #读取csv数据
    for i in range(len(names)):
        #fileName = name[i]+'.csv'
        fileName = names[i].split('.')
        with open(fileName[0]+".csv")as f:
            reader = csv.reader(f)
            for row in reader:
                ENB = row[0]
                ENBs.append(ENB)
                ID  = row[1]
                IDs.append(ID)
                Value = row[2]
                Values.append(Value)
                #创建一个Workbook对象，相当于创建了一个EXCEL文件
                book = xlwt.Workbook(encoding='utf-8',style_compression=0)
                #style_compression表示是否压缩
                #创建一个sheet对象，一个sheet对象对应Excel文件中的一张表格
                sheet = book.add_sheet('MRO解析',cell_overwrite_ok = True)
                #添加数据
                sheet.write(0,0,'ENB') #其中的O行，0列代表单元格,
                sheet.write(0,1,'ID')
                sheet.write(0,2,'RIP_Value')
                #name = SourceFile.split('.')
                for i in range(1,len(IDs)):
                    sheet.write(i,0,ENBs[i])
                    sheet.write(i,1,IDs[i])
                    sheet.write(i,2,Values[i])
                book.save(fileName[0]+'.xls')
                
                del book
                del sheet
        ENBs.clear()
        IDs.clear()
        Values.clear()

def remove_csv_File(path):
    f = glob.iglob(path+"\*.csv")
    for x in f:
        os.remove (x)
    print('ok!')
                
if __name__ == "__main__":
    print('*'*20)
    filePath = input("需要解析的MRO.xml文件路径(例如:D:\文件夹)：")
    print("文件解析文件生成在程序所在的目录中！")
    print('生成中>>>>>>！')
    print('*'*20)
    read_MRO(filePath)
    print("请稍等.....程序正在努力跑-。-")
    remove_csv_File(filePath)
    os.system('pause')
