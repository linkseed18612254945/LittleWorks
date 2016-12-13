import xlrd, xlwt
import os
import win32com.client
import shutil


class convert_xlsx:

    def __init__(self):
        self.__excel_names = ['发行股票', '政府补助', '董事会与高管_现任管理层', '董事会与高管_离任高管', '高管薪酬和持股']
        self.__base_path = r'C:\Users\51694\PycharmProjects\LittleWorks\年报数据\整合数据'
        self.__xl = self.__excel_create(0)
        self.__file_ids = os.listdir(self.__base_path)
        self._count = 0
        self.__file_num = 19

    def __get_path(self, stock_id, stock_name):
        xls_flag = False
        id_dir_path = self.__base_path + '\\' + stock_id
        excel_list = os.listdir(id_dir_path)
        if name + '.xlsx' in excel_list:
            file_name = stock_name + '.xlsx'
        else:
            xls_flag = True
            file_name = stock_name + '.xls'
        file_path = id_dir_path + '\\' + file_name
        return file_path, xls_flag

    def __get_num(self, stock_id):
        excel_num = len(os.listdir(self.__base_path + '\\' + stock_id))
        return excel_num

    def check(self, check_num=0):
        num_error_list = []
        for stock_id in self.__file_ids:
            excel_num = self.__get_num(stock_id)
            if check_num == 0:
                if excel_num != self.__file_num:
                    num_error = stock_id + '  文件数为: ' + str(excel_num) + '\n'
                    num_error_list.append(num_error)
            for stock_name in self.__excel_names:
                file_path, xls_flag = self.__get_path(stock_id, stock_name)
                if check_num == 0:
                    xlrd.open_workbook()





    @staticmethod
    def __excel_create(visible):
        xl = win32com.client.DispatchEx('Excel.Application')
        xl.visible = visible
        xl.DisplayAlerts = False
        return xl

if __name__ == '__main__':
    count = 0
    excel_name = ['发行股票', '政府补助', '董事会与高管_现任管理层', '董事会与高管_离任高管', '高管薪酬和持股']
    base_path = r'C:\Users\51694\PycharmProjects\LittleWorks\年报数据\整合数据'
    # error_path = r'C:\Users\51694\PycharmProjects\LittleWorks\error_num.txt'
    xl = excel_create(0)
    file_ids = os.listdir(base_path)
    for id in file_ids:
        count += 1
        for name in excel_name:
            xls_flag = False
            id_dir_path = base_path + '\\' + id
            excel_list = os.listdir(id_dir_path)
            if name+'.xlsx' in excel_list:
                file_name = name + '.xlsx'
            else:
                xls_flag = True
                file_name = name + '.xls'
            file_path = id_dir_path + '\\' + file_name
            workbook = xl.Workbooks.Open(file_path)
            workbook.SaveAs(Filename=id_dir_path + '\\' + name + '.xlsx', FileFormat=51)
            workbook.Close()
            if xls_flag:
                os.remove(file_path)
        print('Processed ' + str(count) + ' files')
        if count == 2:
            break
    xl.Application.Quit()