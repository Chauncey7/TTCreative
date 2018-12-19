# -*- coding: utf-8 -*-
import xlrd
import xlwt
from xlutils.copy import copy
from config import *
from add_split import add_spliter


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
            except IndexError:
                break


    def make_staticadd(self,new_book):
        #生成第标准地址
        book = xlrd.open_workbook('标准文件.xls')
        sheet2 = book.sheet_by_index(1)
        row_number = 1

        info = []
        while 1:
            try:
                row = sheet2.row_values(row_number)
                info.append((row[5],row[7]))
                row_number += 1
            except IndexError:
                break

        GFnumber = GF_START
        add_list, GF_list = add_spliter(self.upnet,GFnumber,info)

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
                sheet.write(new_rownumber, 10, add_list[new_rownumber-1])
                sheet.write(new_rownumber, 11, '')
                sheet.write(new_rownumber, 12, '家庭场景')
                sheet.write(new_rownumber, 13, '{}-{}'.format(self.upnet,GF_list[new_rownumber-1]))
                new_rownumber += 1
            except IndexError:
                break

    def make_box(self,new_book):
        #生成箱体表
        book = xlrd.open_workbook('标准文件.xls')
        sheet2 = book.sheet_by_index(1)
        row_number = 1



        info = []
        while 1:
            #读取数据

            # row = sheet2.row_values(row_number)
            # info.append((row[1], row[2], row[3], row[4], row[5], row[10], row[0]))
            # row_number += 1


            try:
                row = sheet2.row_values(row_number)
                info.append((row[1],row[2],row[3],row[4],row[5],row[10],row[0]))
                row_number += 1
            except IndexError:
                break


        GJ = []#光交编号
        box_name = []#分纤箱名称
        jd = []#经度
        wd = []#纬度
        add6 = []#6级地址
        upGJname = []#上联光交箱名称
        gl_name = []#光缆名称
        duanzi = []#上联光交对应端子信息
        box_sh = []#分纤箱芯序号
        zsl = []#R列
        kxsl = []#s列


        GFnumber = GF_START
        for i in info:
            if type(i[1]) == str:

                a,b = i[1].split(',')
                core_number = 1#分纤箱芯序号每次都是从1开始
                for j in range(eval(a),eval(b)+1):
                    GJ.append(i[6])#光交列表变化，需添加  标准表0列
                    upGJname.append(i[5])#光交名称10列
                    box_name.append('{}-{}{}-GF{}'.format(self.upnet, self.level6, i[4], str(GFnumber).zfill(3)))
                    jd.append(i[3])
                    wd.append(i[2])
                    add6.append(i[4])
                    gl_name.append('{}{}{}-{}{}GF{}'.format(self.upnet, self.level6, i[6], self.upnet, i[4],
                                                            str(GFnumber).zfill(3)))
                    if i[0]:
                        '''
                        96芯不同于普通芯
                        端子命名不同
                        '''
                        a = i[0][0]#取出纤盘信息第一位A或者B
                        b = i[0][1:]
                        duanzi.append('{}-{}-{}'.format(a,str(b).zfill(2),j))
                    else:
                        duanzi.append('A-01-{}'.format(j))
                    box_sh.append(core_number)
                    zsl.append(12)#其中一个有变化
                    kxsl.append(12)
                    core_number += 1
                GFnumber += 1
                continue


            # if (i[1] == 1)|(i[1] == 7):
            #
            #     for j in range(6):
            #         GJ.append(i[6])
            #         upGJname.append(i[5])
            #         box_name.append('{}-{}-GF{}'.format(self.upnet,i[4],str(GFnumber).zfill(3)))
            #         jd.append(i[3])
            #         wd.append(i[2])
            #         add6.append(i[4])
            #         gl_name.append('{}{}{}-{}{}GF{}'.format(self.upnet,self.level6,i[6],self.upnet,i[4],str(GFnumber).zfill(3)))
            #         if i[0]:
            #             duanzi.append('{}-{}-{}'.format(i[0][0],str(i[0][1:]).zfill(2),j+int(i[1])))
            #         else:
            #             duanzi.append('******')
            #         box_sh.append(j+1)
            #         zsl.append(6)
            #         kxsl.append(6)
            #     GFnumber += 1
            #     continue
            #
            # else:
            #     for j in range(12):
            #         GJ.append(i[6])
            #         upGJname.append(i[5])
            #         box_name.append('{}-{}-GF{}'.format(self.upnet, i[4], str(GFnumber).zfill(3)))
            #         jd.append(i[3])
            #         wd.append(i[2])
            #         add6.append(i[4])
            #         gl_name.append('{}{}{}-{}{}GF{}'.format(self.upnet, self.level6, i[6], self.upnet, i[4],
            #                                                 str(GFnumber).zfill(3)))
            #         if i[0]:
            #             duanzi.append('{}-{}-{}'.format(i[0][0], str(i[0][1:]).zfill(2), j+1))
            #         else:
            #             duanzi.append('******')
            #         box_sh.append(j + 1)
            #         zsl.append(12)
            #         kxsl.append(12)
            #     GFnumber += 1
            #     continue
        table = []
        number = 0
        while 1:
            try:
                table.append((number+1,
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
                              '{}-{}-{}'.format(self.upnet,upGJname[number],GJ[number]),
                              gl_name[number],
                              12,
                              12,
                              duanzi[number],
                              box_sh[number],
                              zsl[number],
                              kxsl[number],
                              '在网',
                              '',
                              '0000',
                              '',
                              '',
                              '飞线',
                              '陈浩',
                              '',
                              '',
                              '自建',
                              '低压电力线',
                              '3.5M以外4M以内',
                              '2米及以上3米以内',
                              ''
                              ))
                number += 1

            except IndexError:
                break
        sheet = new_book.get_sheet(2)  # 获取到第一个sheet页
        for i,j in enumerate(table):
            for k,l in enumerate(j):
                sheet.write(i+1, k, l)

    def make_first(self,new_book):
        # 生成1级分光器名称
        jl1j_o = []#级联一级（旧）
        jl1j_n = []#级联一级（新）
        upOLT = []#上联OLT
        updk = []#上联端口
        slzymc = []#F列
        location = []#安装位置
        yjxl = []#一级下联

        book = xlrd.open_workbook('标准文件.xls')
        sheet2 = book.sheet_by_index(1)

        yj_firstname = sheet2.row_values(0)[9]
        row_number = 1
        info = []
        while 1:
            # 读取数据
            try:
                row = sheet2.row_values(row_number)
                item = (row[9],row[10],row[0])
                hash(item)
                if item not in info:
                    info.append(item)
                row_number += 1
            except IndexError:
                break
        for i in info:
            posnum = 1
            '''普通处理方法如下'''
            # if len(i) < 6:
            #     for j in range(8):
            #         jl1j_o.append('FZHM{}-1-1-{}-虚拟分光器'.format(yj_firstname,i.replace(',','-')))
            #         upOLT.append('FZMH{}'.format(yj_firstname))
            #         updk.append('1-1-{}'.format(i.replace(',','-')))
            # else:
            #     print(i)
            #     a,b = i.split(',')
            #     for j in range(8):
            #         jl1j_o.append('FZHM{}-1-1-{}-虚拟分光器'.format(a,b))
            #         upOLT.append('FZMH{}'.format(a))
            #         updk.append('1-1-{}'.format(b))
            # for j in range(8):
            #     jl1j_n.append('{}-{}-{}-一级POS{}'.format(self.upnet,self.level6,self.GJ,posnum))
            #     slzymc.append('F1')
            #     location.append('{}-{}-{}'.format(self.upnet,self.level6,self.GJ))
            #     yjxl.append(j+1)
            # posnum += 1
            '''96芯处理情况'''
            for j in i[0].split(','):
                for k in range(8):
                    a,b = j.split('-',1)
                    '''拼接第一列'''
                    jl1j_o.append('{}{}{}-{}{}{}'.format(sheet4_first,a,sheet4_middle1,sheet4_middle2,b,sheet4_end))
                    '''拼接第二列'''
                    jl1j_n.append('{}-{}-{}-一级POS{}'.format(self.upnet,i[1],i[2],str(posnum).zfill(3)))
                    upOLT.append('{}{}{}'.format(sheet4_first,a,sheet4_middle1))
                    updk.append('{}{}'.format(sheet4_middle2,b))
                    slzymc.append('F1')
                    location.append('{}-{}-{}'.format(self.upnet,i[1],i[2]))
                    yjxl.append(k+1)
                posnum += 1

        table = []
        number = 0
        while 1:
            try:
                table.append((jl1j_o[number],
                              jl1j_n[number],
                              '1:8',
                              upOLT[number],
                              updk[number],
                              slzymc[number],
                              '',
                              '',
                              location[number],
                              yjxl[number],
                              '',
                              '在网',
                              '级联',
                              '盒式分光器',
                              '',
                              '',
                              '0000',
                              '陈浩'))
                number += 1
            except IndexError:
                break
        sheet = new_book.get_sheet(3)  # 获取到第4个sheet页
        for i, j in enumerate(table):
            for k, l in enumerate(j):
                sheet.write(i + 1, k, l)


    def make_secodelight(self,new_book):
        #生成2级分光器表
        second_name = []#二级分光器名称
        locations = []#安装位置
        up_fgq = []#上联分光器名称
        zldk = []#G列
        ejdz = []#H列

        book = xlrd.open_workbook('标准文件.xls')
        sheet2 = book.sheet_by_index(1)
        row_number = 1
        info = []
        while 1:
            #读取数据
            try:
                row = sheet2.row_values(row_number)
                info.append((row[5],row[6],row[8]))
                row_number += 1
            except IndexError:
                break
        GFnumber = GF_START
        for i in info:
            if type(i[1]) != str:
                second_name.append('{}-{}-GF{}-二级POS001'.format(self.upnet,i[0],str(GFnumber).zfill(3)))
                locations.append('{}-{}-GF{}'.format(self.upnet,i[0],str(GFnumber).zfill(3)))
                up_fgq.append(', ')
                zldk.append(i[2])
                ejdz.append('A-1-1-{}'.format(int(i[1])))
            else:
                second_name.append('{}-{}-GF{}-二级POS001'.format(self.upnet, i[0], str(GFnumber).zfill(3)))
                second_name.append('{}-{}-GF{}-二级POS002'.format(self.upnet, i[0], str(GFnumber).zfill(3)))
                locations.append('{}-{}-GF{}'.format(self.upnet, i[0], str(GFnumber).zfill(3)))
                locations.append('{}-{}-GF{}'.format(self.upnet, i[0], str(GFnumber).zfill(3)))
                up_fgq.append('zzzz')
                up_fgq.append('zzzz')
                a,b = i[2].split(',')
                zldk.append(a)
                zldk.append(b)
                a, b = i[1].split(',')
                ejdz.append('A-1-1-{}'.format(a))
                ejdz.append('A-1-1-{}'.format(b))
            GFnumber += 1
        table = []
        number = 0
        while 1:
            try:
                table.append(('',
                             '家庭个人宽带',
                             second_name[number],
                             locations[number],
                             '1:8',
                             up_fgq[number],
                             zldk[number],
                             ejdz[number],
                             '',
                             '在网',
                             'FTTH',
                             '卡片式分光器',
                             '',
                             '',
                             '0000',
                             '陈浩'))
                number += 1
            except IndexError:
                break
        # new_book = copy(book2)
        sheet = new_book.get_sheet(4)  # 获取到第4个sheet页
        for i, j in enumerate(table):
            for k, l in enumerate(j):
                sheet.write(i + 1, k, l)




    def main(self):
        #拷贝模板
        book2 = xlrd.open_workbook('模板.xls')
        new_book = copy(book2)
        #生成各列表
        self.get_start()
        for i in WANT_TO.split(','):
            if i == '1':
                self.make_staticadd(new_book)
                continue
            if i == '2':
                self.make_box(new_book)
                continue
            if i == '3':
                self.make_first(new_book)
                continue
            if i == '4':
                self.make_secodelight(new_book)
                continue
            else:
                self.make_staticadd(new_book)
                self.make_box(new_book)
                self.make_first(new_book)
                self.make_secodelight(new_book)

        #粘贴

        #填充固定项

        #保存
        new_book.save('stu_new.xls')




if __name__ == '__main__':
    shoot = Shoot()
    shoot.main()



