import win32com.client
import os
import shutil
import sys

SOURCE_PATH = r'C:\Users\51694\Desktop\test'
RESTORE_PATH = r'C:\Users\51694\restore'
STOCK_ID_PATH = 'stock_id.txt'
ONLY_RENAME_TABLE_NUM = 5
TOTAL_TABLE_NUM = 8
TABLE_NAME_PATH = 'wind_table_name.txt'


class ExcelMethod:

    @staticmethod
    def create_empty(table_name):
        xl_workbook = xl.Workbooks.Add()
        xl_workbook.SaveAs(SOURCE_PATH + '\\' + table_name + '.xlsx')
        xl_workbook.Close()

    @staticmethod
    def excel_create(visible):
        xl = win32com.client.DispatchEx('Excel.Application')
        xl.visible = visible
        xl.DisplayAlerts = False
        return xl

def get_tables():
    stocks_table = []
    source_tables = os.listdir(SOURCE_PATH)
    start = 0
    start_stock = str(stock_line)

    for i, name in enumerate(source_tables):
        stock_id = name[:name.find('-')]
        if stock_id != start_stock:
            temp = source_tables[start:i]
            temp.sort(key=lambda x: int(x[x.find('-') + 1:x.find('.')]))
            stocks_table.append(temp)
            start_stock = stock_id
            start = i
    temp = source_tables[start:]
    temp.sort(key=lambda x: int(x[x.find('-') + 1:x.find('.')]))
    stocks_table.append(temp)
    return stocks_table


def init():
    tables = get_tables()
    with open(TABLE_NAME_PATH, 'r', encoding='utf-8') as f:
        table_names = f.readlines()
    with open(STOCK_ID_PATH, 'r') as f:
        stocks_id = f.readlines()
    xl = ExcelMethod.excel_create(0)
    return tables, table_names, stocks_id, xl


def complete_the_empty(stock_id):
    for stock_tables in tables:
        if len(stock_tables) == 12:
            continue
        else:
            for i in range(1, 13):
                t_n = str(stock_id) + '-' + str(i)
                if t_n + '.xlsx' not in stock_tables and t_n + '.xls' not in stock_tables:
                    ExcelMethod.create_empty(t_n)
        stock_id += 1


def restore_table(cp_tables):
    for ts in cp_tables:
        stock_id = int(ts[0][:ts[0].find('-')])
        stock_num = str(stocks_id[stock_id - 1].strip())
        print("正在复制股票代码%s的文件" % stock_num)
        new_dir_path = RESTORE_PATH + '\\' + stock_num
        os.mkdir(new_dir_path)
        table_id = 0
        for table in ts:
            source_table_path = SOURCE_PATH + '\\' + table
            target_path = new_dir_path + '\\' + table_names[table_id].strip() + '.xlsx'
            if table[table.find('.') + 1:] == 'xlsx':
                shutil.copy(source_table_path, target_path)
            elif table[table.find('.') + 1:] == 'xls':
                xl_workbook = xl.Workbooks.Open(source_table_path)
                xl_workbook.SaveAs(Filename=target_path)
                xl_workbook.Close()
            else:
                raise TypeError('Wrong excel type')
            table_id += 1
    xl.Application.Quit()
    print("处理完成")

if __name__ == '__main__':
    stock_line = 1
    tables, table_names, stocks_id, xl = init()
    complete_the_empty(stock_line)
    cp_tables = get_tables()
    restore_table(cp_tables)







