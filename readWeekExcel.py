#coding=utf-8

'''
Create 2018-05-18
Update 2018-05-18
Author: cking
Github: https://github.com/cking0821

'''

import xlrd,xlwt
import xdrlib,sys
import os,re


# 列出目录下的所有文件和目录
def listDir():
    list = os.listdir()
    listRe = []
    for i in list:
        if (str(i[-4:]).__contains__("xls") and bool(re.search(r'\d', i))): #
            listRe.append(i)
    # print (listRe)
    return listRe

#open excel
def open_excel(file):
    try:
        data = xlrd.open_workbook(file)
        return data
    except Exception as e:
        print(str(e))

#columnName  colnameindex：表头列名所在行的所以
def getColumnIndex(table, columnName):
    columnIndex = None
    for i in range(table.ncols):
        if (table.cell_value(0, i).__contains__(columnName)):
            columnIndex = i
            break
    return columnIndex

#根据名称获取Excel表格中的数据
def excel_table_byname(columnName,colnameindex=0,by_name=u'Sheet1'):
    fileList = listDir()
    list = []
    f = xlwt.Workbook()  # 创建工作簿
    sheet1 = f.add_sheet(u'sheet1', cell_overwrite_ok=True)  # 创建sheet
    row0 = [u'文件名称',u'姓名时间', u'周报内容']

    # 生成第一行
    for i in range(0, len(row0)):
        sheet1.write(0, i, row0[i])

    i = 1
    for file in fileList:
        data = open_excel(file)
        table = data.sheet_by_name(by_name)
        print("****")
        fileName = file #[:-5]    #文件名
        print (fileName)
        try:
            fileContent = table.cell_value(1,getColumnIndex(table,columnName))  #文件内容
            list.append(fileName)
            list.append(columnName)
            list.append(fileContent)

            for j, p in enumerate(list):
                sheet1.write(i, j, p)
            i +=1
            f.save('weekBook.xls')
            list = []
        except Exception as e :
            print(e)


def main(columnName):
    excel_table_byname(columnName)

if __name__ == '__main__':
    columnName = '04.23-04.27'  #列名的情况
    main(columnName)




