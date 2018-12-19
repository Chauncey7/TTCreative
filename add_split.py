#地址分裂器
import re

upnet = '连江'
def add_spliter(upnet,gf_star,info):
    '''
    info为外部传入参数
    info[0]为箱体地址
    info[1]为箱体覆盖
    '''
    rownumber = 0
    add_list = []
    GF_list = []
    GFnumber = gf_star

    temp = []

    while 1:
        try:
            row = info[rownumber]
            try:
                front = re.findall(r'(.*?)\d+-?\d*-?\d*号.*', row[0])[0]
            except:
                print(row, rownumber + 1)
                front = re.findall(r'(.*?)右侧墙上|左侧墙上|墙上', row[0])[0]

            adds = row[1].split('/')

            for i in adds:
                if not re.match(r'\d+-?\d*',i):
                    add_list.append('{}号'.format(i))
                else:
                    add_list.append('{}{}号'.format(front, i))
                GF_list.append('{}-{}-GF{}'.format(upnet,row[0],str(GFnumber).zfill(3)))
                temp.append(rownumber+1)
            rownumber += 1
            GFnumber += 1
        except IndexError:
            break

    return add_list,GF_list
if __name__ == '__main__':

    a = []
    print(a[1])