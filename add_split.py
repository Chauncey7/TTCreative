#地址分裂器
import xlrd
import xlwt
import re

upnet = '连江'
def add_spliter():
    #读取文件
    workbook = xlwt.Workbook(encoding='ascii')
    worksheet = workbook.add_sheet('address')
    workbook2 = xlrd.open_workbook(r'12-14test1.xlsx')
    sheet1 = workbook2.sheet_by_index(0)
    rownumber = 0
    add_list = []
    GF_list = []
    GFnumber = 14

    temp = []

    while 1:
        try:
            print(rownumber+1)
            row = sheet1.row_values(rownumber)
            try:
                front = re.findall(r'(.*?)\d+-?\d*号.*', row[0])[0]
            except:
                print(row, rownumber + 1)
                front = re.findall(r'(.*?)右侧墙上|左侧墙上|墙上', row[0])[0]

            adds = row[1].split('/')

            for i in adds:
                if len(i) > 4:
                    add_list.append('{}号'.format(i))
                else:
                    add_list.append('{}{}号'.format(front, i))
                GF_list.append('{}-{}-GF{}'.format(upnet,row[0],str(GFnumber).zfill(3)))
                temp.append(rownumber+1)
            rownumber += 1
            GFnumber += 1
        except Exception as e:
            print(e)
            break

    num = 0
    while 1:
        try:
            worksheet.write(num, 0, temp[num])
            worksheet.write(num, 2, add_list[num])
            worksheet.write(num, 1, GF_list[num])
            num += 1
        except Exception as e:
            print(e)
            break
    workbook.save('newadds.xls')
if __name__ == '__main__':
    add_spliter()