import xlrd
import os
import time
import re


def bumen_rename(path, year):
    rename_form = str(year) + '年部门比较数据'
    names = os.listdir(path)
    for i in names:
        file = path + '/' + i
        data = xlrd.open_workbook(file)
        table = data.sheet_by_index(0)
        cell = table.cell(1, 1).value
        rename_new = path + '/' + cell + rename_form + '.xls'
        os.rename(file, rename_new)


def diqu_rename(path, year):
    rename_form = str(year) + '年地区比较数据'
    names = os.listdir(path)
    for i in names:
        file = path + '/' + i
        data = xlrd.open_workbook(file)
        table = data.sheet_by_index(0)
        cell = table.cell(1, 0).value
        if cell.find('省') >= 0:
            index = cell.find('省') + 1
        elif cell.find('自治区') >= 0:
            index = cell.find('自治区') + 3
        elif cell.find('兵团') >= 0:
            index = cell.find('兵团') + 2
        else:
            index = cell.find('市') + 1
        prov = cell[:index]
        print(prov)
        rename_new = path + '/' + prov + rename_form + '.xls'
        os.rename(file, rename_new)


def diquxian_rename(path, year):
    rename_form = str(year) + '年地区比较数据'
    names = os.listdir(path)
    for i in names:
        file = path + '/' + i
        data = xlrd.open_workbook(file)
        table = data.sheet_by_index(0)
        cell = table.cell(1, 0).value
        zhixia = re.compile('北京|上海|重庆|天津')
        if re.search(zhixia, cell) is not None:
            if cell.find('市辖区') >= 0:
                index = cell.find('市辖区') + 3
            else:
                index = cell.find('市县') + 2
        else:
            if cell.find('盟') >= 0:
                index = cell.find('盟') + 1
            elif cell.find('自治州') >= 0:
                index = cell.find('自治州') + 3
            elif cell.find('行政区划') >= 0:
                index = cell.find('行政区划') + 4
            elif cell.find('行政单位') >= 0:
                index = cell.find('行政单位') + 4
            elif cell.find('地区') >= 0:
                index = cell.find('地区') + 2
            elif cell.find('市辖区') >= 0:
                 if cell[:cell.find('市辖区')].find('市') >= 0:
                     index = cell.find('市辖区')
                 else:
                     index = cell.find('市辖区') + 3
            elif cell.find('市') >= 0:
                index = cell.find('市') + 1
            else:
                index = -3
        prov = cell[:index]
        print(prov)
        rename_new = path + '/' + prov + rename_form + '.xls'
        os.rename(file, rename_new)


def check_sheng(path):
    names = os.listdir(path)
    for i in names:
        count = 1
        file = path + '/' + i
        data = xlrd.open_workbook(file)
        table = data.sheet_by_index(0)
        col = table.col(0)
        cell = table.cell(1, 0).value
        for j in col[1:]:
            if j.value[:2] != cell[:2]:
                rename_new = '错误' + str(time.time())
                os.rename(file, rename_new)
                count += 1
        print(count)



path = 'C:/Users/51694/Desktop/数据/2006/部门'



# check(path)
bumen_rename(path, 2006)