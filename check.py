import os
import xlrd


if __name__ == '__main__':
    count = 0
    excel_name = ['发行股票', '政府补助', '董事会与高管_现任管理层', '董事会与高管_离任高管', '高管薪酬和持股']
    base_path = r'C:\Users\Administrator\PycharmProjects\LittleWorks\tee\data'
    # error_path = r'C:\Users\51694\PycharmProjects\LittleWorks\error_num.txt'
    file_ids = os.listdir(base_path)
    for id in file_ids:

        count += 1
        for name in excel_name:
            xls_flag = False
            id_dir_path = base_path + '\\' + id
            excel_list = os.listdir(id_dir_path)
            if len(excel_list) != 19:
                print('---------------' + id + '----------------')
                print(excel_list)
                # if name + '.xls' in excel_list:
                #     file_name = name + '.xls'
                #     file_path = id_dir_path + '\\' + file_name
                #     os.remove(file_path)
            # if name+'.xlsx' in excel_list:
            #     file_name = name + '.xlsx'
            # else:
            #     xls_flag = True
            #     file_name = name + '.xls'
            # file_path = id_dir_path + '\\' + file_name
            # try:
            #     xl = xlrd.open_workbook(file_path)
            # except:
            #     with open('convert_id.txt', 'a', encoding='utf-8') as f:
            #         f.write(id+' ' + file_name + '\n')
            #     print(id, file_name)

        # if count <= 336:
        #     continue