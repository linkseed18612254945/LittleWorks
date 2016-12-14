import time
import os
import win32com.client




def excel_create(visible):
    xl = win32com.client.DispatchEx('Excel.Application')
    xl.visible = visible
    xl.DisplayAlerts = False
    return xl

if __name__ == '__main__':
    start_time = time.time()
    base_path = r'C:\Users\Administrator\PycharmProjects\LittleWorks\data'
    error_path = r'C:\Users\Administrator\PycharmProjects\LittleWorks\error_num.txt'
    file_ids = os.listdir(base_path)
    count = 0
    write_row = 1
    output_num = 0
    bottom = 1
    sheet_num = 1
    form_title = ['报告期', '姓名', '职务', '薪酬（元）', '股票代码']
    sheet = 'Sheet' + str(sheet_num)
    w_application = excel_create(1)
    w_workbook = w_application.Workbooks.Add()
    w_worksheet = w_workbook.Worksheets(sheet)
    w_worksheet.Range(w_worksheet.Cells(1, 1), w_worksheet.Cells(1, 5)).Value = form_title
    for id in file_ids:
        id_col = []
        top = bottom + 1
        count += 1
        xl_path = base_path + '\\' + id
        with open(error_path) as f:
            error_list = f.readlines()
        if id[:-4]+'\n' in error_list:
            continue
        r_application = excel_create(0)
        r_workbook = r_application.Workbooks.Open(xl_path)
        r_worksheet = r_workbook.Worksheets(1)
        nrow = 0
        while r_worksheet.Cells(nrow+1, 1).Value not in [None, '']:
            nrow += 1
        bottom = top + nrow - 3
        r_value = r_worksheet.Range(r_worksheet.Cells(3, 1), r_worksheet.Cells(nrow, 4)).Value
        w_worksheet.Range(w_worksheet.Cells(top, 1), w_worksheet.Cells(bottom, 4)).Value = r_value
        for i in range(nrow):
            id_col.append(id[:-4])
        w_worksheet.Range(w_worksheet.Cells(top, 5), w_worksheet.Cells(bottom, 5)).NumberFormat = '@'
        w_worksheet.Range(w_worksheet.Cells(top, 5), w_worksheet.Cells(bottom, 5)).Value = id_col
        r_application.Application.Quit()
        print('Processing: ' + str(count) + ' files' + ' ,' + id)
        if count % 20 == 0:
            runtime = time.time() - start_time
            print('Program has run %.2f s' % runtime)
        if count == 500:
            print('--------------------Processed 500files------------------------')
            count = 0
            bottom = 1
            sheet_num += 1
            w_workbook.Worksheets.Add()
            sheet = 'Sheet' + str(sheet_num)
            print(sheet)
            w_worksheet = w_workbook.Worksheets(sheet)
            w_worksheet.Range(w_worksheet.Cells(1, 1), w_worksheet.Cells(1, 5)).Value = form_title
    runtime = time.time() - start_time
    print('Program total run %.2f s' % runtime)
    w_workbook.SaveAs(Filename=r'C:\Users\Administrator\PycharmProjects\LittleWorks\output.xlsx')
    # w_application.Application.Quit()



