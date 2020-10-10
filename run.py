import pickle
import random
import time

import xlwt

from __init__ import *


def format(num):
    """
    格式化编号
    :param num:
    :return:
    """
    if num < 10:
        fnum = "00" + str(num)
    elif num < 100:
        fnum = "0" + str(num)
    elif num < 1000:
        fnum = str(num)
    else:
        fnum = ""
    return fnum

class PersonList():

    def __init__(self):
        self.person = max_person_each_dp
        self.departments = int(max_person / max_person_each_dp)

    def readTemplats(self, filename):
        file = INPUTPATH + "\\" + filename
        with open(file, 'rb') as f:
            title = f.readlines()
            for t in title:
                t = t.decode('big5')
                print(t)
            return title

    def makePerson(self):
        all_person = {}

        for dp in range(1, self.departments + 1):
            dp_person = []
            dp = format(dp)
            for num in range(1, self.person + 1):
                num = format(num)
                t = time.localtime()
                _person_info = []
                _person_info.append("jy%ss0000000%s" % (dp, num))
                _person_info.append("jy%ss0000000%s" % (dp, num))
                _person_info.append(random.randint(0, 2))
                _person_info.append("jy%ss00IDNO0000000%s" % (dp, num))
                _person_info.append("jy%ss00ICNO0000000%s" % (dp, num))
                _person_info.append("%s/%s/%s" % (t.tm_year, t.tm_mon, t.tm_mday))
                _person_info.append(random.randint(11100000000, 19000000000))
                _person_info.append("ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz!@#$%^&*()-_")

                dp_person.append(_person_info)
            all_person["dp:" + dp] = dp_person

        return all_person

    def writePersonList(self, all_person):
        language = input("语言")
        if "cn" == language:
            title = title_cn
        else:
            title = title_en

        for dp in range(1, self.departments + 1):
            # print(all_person)
            dp = format(dp)
            dp_person = all_person["dp:" + dp]
            file = OUTPUTPATH + "\\" + language + dp + ".xls"

            f = xlwt.Workbook()
            sheet1 = f.add_sheet(u'sheet1', cell_overwrite_ok=True)  # 创建sheet
            # 写入表头
            for row in range(len(title)):
                sheet1.write(0, row, title[row])
            # 将数据写入第 i 行，第 j 列
            i = 1
            for person in dp_person:
                for row in range(len(person)):
                    sheet1.write(i, row, person[row])
                i = i + 1

            f.save(file)  # 保存文件


if __name__ == '__main__':
    pl = PersonList()
    # title = pl.readTemplats("Cn01.xls")
    title = title_cn
    all_person = pl.makePerson()
    pl.writePersonList(all_person)
