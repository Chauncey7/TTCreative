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
        workbook = xlwt.Workbook(encoding='ascii')
        worksheet = workbook.add_sheet('address')
        workbook2 = xlrd.open_workbook(r'标准文件.xls')
        sheet1 = workbook2.sheet_by_index(0)
        row = sheet1.row_values(2)#从第一个小区信息表中获取小区基本信息

        self.level1 = row[4]#一级地址
        self.level2 = row[5]#二级地址
        self.level3 = row[6]#三级地址
        self.level4 = row[7]#四级地址
        self.adm_coding = row[8]#单位执行编码
        self.level5 = row[9]#五级地址
        self.level6 = row[10]#六级
        self.manager = row[15]#社区经理
        self.copyman = row[19]#录入人
        self.upnet = row[21]#上级网络

        sheet2 = workbook2.sheet_by_index(2)
        row_number = 1

        self.address = []
        while 1:
            try:
                row = sheet2.row_values(row_number)
                self.address.append(row[0])
                row_number += 1
            except Exception as e:
                print(e)
                break


    def make_staticadd(self):
        #生成第一表
        book = xlrd.open_workbook('标准文件.xls')
        book2 = xlrd.open_workbook('模板.xls')
        sheet2 = book.sheet_by_index(1)
        row_number = 1

        address = []
        while 1:
            try:
                row = sheet2.row_values(row_number)
                address.append((row[5],row[7].count('..')+1))
                row_number += 1
            except Exception as e:
                print(e)
                break

        final_add = []
        GFnumber = 3
        for i in address:
            for j in range(i[1]):
                final_add.append(i[0]+'-GF{}'.format(str(GFnumber).zfill(3)))
            GFnumber += 1
        print(set(final_add))

        # 复制一个excel
        new_book = copy(book2)  # 复制了一份原来的excel
        sheet = new_book.get_sheet(1)  # 获取到第一个sheet页

        new_rownumber = 1 #从第一行开始写入新表
        while 1:
            try:
                sheet.write(new_rownumber, 0, new_rownumber)  # 写入excel，第一个值是行，第二个值是列
                sheet.write(new_rownumber, 1, self.level1)
                sheet.write(new_rownumber, 2, self.level2)
                sheet.write(new_rownumber, 3, self.level3)
                sheet.write(new_rownumber, 4, self.level4)
                sheet.write(new_rownumber, 5, self.adm_coding)
                sheet.write(new_rownumber, 6, self.level5)
                sheet.write(new_rownumber, 7, self.level6)
                sheet.write(new_rownumber, 8, '/')
                sheet.write(new_rownumber, 9, '/')
                sheet.write(new_rownumber, 10, self.address[new_rownumber-1])
                sheet.write(new_rownumber, 11, '')
                sheet.write(new_rownumber, 12, '家庭场景')
                sheet.write(new_rownumber, 13, '{}-{}'.format(self.upnet,final_add[new_rownumber-1]))
                new_rownumber += 1
            except Exception as e:
                print(e)
                break
        new_book.save('stu_new.xls')

    def make_box(self):
        #生成箱体表
        book = xlrd.open_workbook('标准文件.xls')
        book2 = xlrd.open_workbook('模板.xls')
        sheet2 = book.sheet_by_index(1)
        row_number = 1

        self.GJ = sheet2.row_values(1)[0]

        info = []
        while 1:
            #读取数据
            try:
                row = sheet2.row_values(row_number)
                info.append((row[1],row[2],row[3],row[4],row[5]))
                row_number += 1
            except Exception as e:
                print(e)
                break

        box_name = []#分纤箱名称
        jd = []#经度
        wd = []#纬度
        add6 = []#6级地址
        gl_name = []#光缆名称
        duanzi = []#上联光交对应光子信息
        box_sh = []#分纤箱芯序号
        zsl = []#R列
        kxsl = []#s列

        GFnumber = 3
        for i in info:
            if i[1] != 12:

                for j in range(6):
                    box_name.append('{}-{}-GF{}'.format(self.upnet,i[4],str(GFnumber).zfill(3)))
                    jd.append(i[3])
                    wd.append(i[2])
                    add6.append(i[4])
                    gl_name.append('{}{}{}-{}{}GF{}'.format(self.upnet,self.level6,self.GJ,self.upnet,i[4],str(GFnumber).zfill(3)))

                    duanzi.append('{}-{}-{}'.format(i[0][0],str(i[0][1:]).zfill(2),j+i[1]))
                    box_sh.append(j+1)
                    zsl.append(6)
                    kxsl.append(6)
                    GFnumber += 1

            else:
                for j in range(12):
                    box_name.append('{}-{}-GF{}'.format(self.upnet, i[4], str(GFnumber).zfill(3)))
                    jd.append(i[3])
                    wd.append(i[2])
                    add6.append(i[4])
                    gl_name.append('{}{}{}-{}{}GF{}'.format(self.upnet, self.level6, self.GJ, self.upnet, i[4],
                                                            str(GFnumber).zfill(3)))

                    duanzi.append('{}-{}-{}'.format(i[0][0], str(i[0][1:]).zfill(2), j+1))
                    box_sh.append(j + 1)
                    zsl.append(12)
                    kxsl.append(12)
                    GFnumber += 1

        table = []
        number = 0
        while 1:
            try:
                table.append((number,
                              '法兰盘分纤箱',
                              box_name[number],
                              jd[number],
                              wd[number],
                              self.adm_coding,
                              '/',
                              self.level6,
                              '',
                              '',
                              add6[number],
                              '{}-{}-{}'.format(self.upnet,self.level6,self.GJ),
                              gl_name[number],
                              ))



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
    # other2()
    # getaddress()
    shoot = Shoot()
    shoot.get_start()
    shoot.make_staticadd()
    # for root,dirs,files in os.walk('test'):
    #     print(files)



