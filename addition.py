# -*- coding: utf-8 -*-
import xlrd
import xlwt
from datetime import date, datetime
import pymysql
from xlutils.copy import copy
import os

class Shoot():
    def get_start(self):
        #初始化读取标准文件
        pass

    def make_device(self):
        #生成设备名称
        pass

    def make_hub(self):
        #生成光缆名称
        pass

    def make_duan(self):
        #生成端子信息
        pass

    def make_secodelight(self):
        #生成二级分光器名称（二级分光器）
        pass

    def main(self):
        #拷贝模板

        #生成各列表

        #粘贴

        #填充固定项

        #保存
        pass


def getaddress():
    #实现地址分裂
    workbook = xlwt.Workbook(encoding='ascii')
    worksheet = workbook.add_sheet('address')
    workbook2 = xlrd.open_workbook(r'标准文件.xls')
    sheet = workbook2.sheet_by_index(0)
    row_number = 0
    write_number = 0
    GF_number = 3
    while 1:
        try:
            row = sheet.row_values(row_number)
            count = len(row[6].split('..'))
            for i in range(count):
                worksheet.write(write_number, 0, label='闽侯-{}-GF{}'.format(row[4],str(GF_number).zfill(3)))
                write_number += 1
            row_number += 1
            GF_number += 1
        except IndexError:
            break
        except:
            row = sheet.row_values(row_number)
            worksheet.write(row_number, 0, label='闽侯-{}-GF{}'.format(row[4],str(GF_number).zfill(3)))
            write_number += 1
            GF_number += 1
            row_number += 1
    workbook.save('Excel_Address.xls')

def other2():
    #实现其他内容填充
    workbook = xlrd.open_workbook(r'标准文件.xls')
    sheet = workbook.sheet_by_index(0)
    i = 0
    info = []
    while 1:
        try:
            rows = sheet.row_values(i)
            i += 1
            info.append(rows)
        except:
            break
    workbook = xlwt.Workbook(encoding='ascii')
    worksheet = workbook.add_sheet('My Worksheet')
    row_number = 0
    GFnumber_star = int(info[0][0][1:])
    GFnumber = GFnumber_star
    for i in info:
        if i[1] == 1:
            for j in range(6):
                worksheet.write(row_number, 0, label='{}-{}-{}'.format(i[0][0],i[0][1:],j+1))
                worksheet.write(row_number, 1, label=i[2])
                worksheet.write(row_number, 2, label=i[3])
                worksheet.write(row_number, 3, label=i[4])
                worksheet.write(row_number, 4, label='闽侯-{}-GF{}'.format(i[4],str(GFnumber).zfill(3)))
                row_number += 1
            GFnumber += 1
            continue
        if i[1] == 7:
            for j in range(6,12):
                worksheet.write(row_number, 0, label='{}-{}-{}'.format(i[0][0], i[0][1:], j + 1))
                worksheet.write(row_number, 1, label=i[2])
                worksheet.write(row_number, 2, label=i[3])
                worksheet.write(row_number, 3, label=i[4])
                worksheet.write(row_number, 4, label='闽侯-{}-GF{}'.format(i[4],str(GFnumber).zfill(3)))
                row_number += 1
            GFnumber += 1
            continue
        if i[1] == 12:
            for j in range(12):
                worksheet.write(row_number, 0, label='{}-{}-{}'.format(i[0][0], i[0][1:], j + 1))
                worksheet.write(row_number, 1, label=i[2])
                worksheet.write(row_number, 2, label=i[3])
                worksheet.write(row_number, 3, label=i[4])
                worksheet.write(row_number, 4, label='闽侯-{}-GF{}'.format(i[4],str(GFnumber).zfill(3)))
                row_number += 1
            GFnumber += 1
            continue
        else:

            start = int(i[1][-1])*12-12
            for j in range(12):
                worksheet.write(row_number, 0, label='A-01-{}'.format(start+j+1))
                worksheet.write(row_number, 1, label=i[2])
                worksheet.write(row_number, 2, label=i[3])
                worksheet.write(row_number, 3, label=i[4])
                worksheet.write(row_number, 4, label='闽侯-{}-GF{}'.format(i[4], str(GFnumber).zfill(3)))
                row_number += 1
            GFnumber += 1

    worksheet2 = workbook.add_sheet('My second')
    row_number = 0
    GFnumber = GFnumber_star
    for i in info:
        if type(i[5]) == float:

            worksheet2.write(row_number, 0, label='闽侯-{}-GF{}'.format(i[4],str(GFnumber).zfill(3)))
            worksheet2.write(row_number, 1, label='闽侯-{}-GF{}-二级POS001'.format(i[4],str(GFnumber).zfill(3)))
            row_number += 1
            GFnumber += 1
            continue
        else:
            worksheet2.write(row_number, 0, label='闽侯-{}-GF{}'.format(i[4], str(GFnumber).zfill(3)))
            worksheet2.write(row_number, 1, label='闽侯-{}-GF{}-二级POS001'.format(i[4], str(GFnumber).zfill(3)))
            row_number += 1
            worksheet2.write(row_number, 0, label='闽侯-{}-GF{}'.format(i[4], str(GFnumber).zfill(3)))
            worksheet2.write(row_number, 1, label='闽侯-{}-GF{}-二级POS002'.format(i[4], str(GFnumber).zfill(3)))
            row_number += 1
            GFnumber += 1

    workbook.save('Excel_Workbook.xls')


def main():
    #creat new excel file
    #put address in sheet1
    addresses = getaddress('坎水村(1).xlsx')
    #put other in sheet2
    pass
if __name__ == '__main__':
    other2()
    getaddress()
