import win32com.client
import os
import shutil
import sys

SOURCE_PATH = r'C:\Users\51694\Desktop\test'
RESTORE_PATH = r'C:\Users\51694\restore'
ONLY_RENAME_TABLE_NUM = 5
TOTAL_TABLE_NUM = 8
ONLY_RENAME_TABLE_NAME = ['发行股票', '政府补助', '董事会与高管_现任管理层', '董事会与高管_离任高管', '高管薪酬和持股']
RESTORE_TABLE_NAME = ['公司介绍', '资金流量', '员工构成']


def check():
    ''' check the tables' number and name '''
    file_id = 1
    table_id = 1
    if len(tables) % TOTAL_TABLE_NUM != 0:
        raise Exception("某股票对应表可能存在漏下，请确认表数量是%d的倍数" % TOTAL_TABLE_NUM)
    for i, name in enumerate(tables):
        if name[:name.find(".")] != str(file_id) + '-' + str(table_id):
            raise Exception("下载表可能存在表名错误或漏下")
        table_id += 1
        if (i + 1) % TOTAL_TABLE_NUM == 0:
            table_id = 1
            file_id += 1


def excel_create(visible):
    xl = win32com.client.DispatchEx('Excel.Application')
    xl.visible = visible
    xl.DisplayAlerts = False
    return xl

if __name__ == '__main__':
    tables = os.listdir(SOURCE_PATH)
    print(tables)
    stock_id = 60001
    xl = excel_create(0)
    check()
    new_dir_path = ''
    for i, name in enumerate(tables):
        table_id = i % 8 + 1
        file_id = i // 8 + 1
        source_table_path = SOURCE_PATH + '\\' + name
        if table_id == 1:
            new_dir_path = RESTORE_PATH + '\\' + str(stock_id)
            os.mkdir(new_dir_path)
            stock_id += 1
            print("正在复制第%d个股票代码文件" % file_id)

        if table_id <= ONLY_RENAME_TABLE_NUM:
            shutil.copy(source_table_path, new_dir_path + '\\' + ONLY_RENAME_TABLE_NAME[table_id - 1] + '.xlsx')
        else:
            xl_workbook = xl.Workbooks.Open(source_table_path)
            xl_workbook.SaveAs(Filename=new_dir_path + '\\' + RESTORE_TABLE_NAME[table_id - 1 - ONLY_RENAME_TABLE_NUM] +'.xlsx')
            xl_workbook.Close()

    xl.Application.Quit()
    print("处理成功，共处理%d个股票代码" % (len(tables) // TOTAL_TABLE_NUM))






