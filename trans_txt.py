import os
import win32com.client


def save_as(load_path, save_path, wd):
    count = 0
    class_names = os.listdir(load_path)
    for class_name in class_names:
        class_path = load_path + '\\' + class_name
        file_names = os.listdir(class_path)
        print(class_name)
        for file_name in file_names:
            file_path = class_path + '\\' + file_name
            save_file_path = save_path + '\\' + class_name + '\\' + file_name[:-5]
            print(save_file_path)
            doc = wd.Documents.Open(file_path)
            doc.SaveAs(save_file_path, 2)
            doc.Close()
        print(DIV)

if __name__ == '__main__':
    load_path = r'C:\Users\51694\PycharmProjects\LittleWorks\data\data'
    save_path = r'C:\Users\51694\PycharmProjects\LittleWorks\savetxt'
    DIV = '------------------------------------------------------------------------'

    wd = win32com.client.DispatchEx('Word.Application')
    wd.visible = 1
    wd.DisplayAlerts = False
    save_as(load_path, save_path, wd)
    wd.Quit()
