import os
#from bs4 import BeautifulSoup
import win32com.client


def save_as(load_path, save_path, wd):
    DIV = '------------------------------------------------------------------------'
    count = 0
    class_names = os.listdir(load_path)
    for class_name in class_names:
        class_path = load_path + '\\' + class_name
        file_names = os.listdir(class_path)
        print(class_name)
        for file_name in file_names:
            file_path = class_path + '\\' + file_name
            file_end = file_name[file_name.find('.')+1:]
            if file_end not in ['doc', 'docx']:
                count += 1
                save_file_path = save_path + '\\' + class_name + '\\' + file_name[:file_name.find('.')]
                with open(file_path, 'r', encoding='utf-8') as f:
                    content = BeautifulSoup(f.read(), 'html.parser')
                print(content.get_text)
            if count == 2:
                return 0
        print(DIV)

if __name__ == '__main__':
    load_path = r'C:\Users\51694\PycharmProjects\LittleWorks\data\data'
    save_path = r'C:\Users\51694\PycharmProjects\LittleWorks\savetxt'


    wd = win32com.client.DispatchEx('Word.Application')
    wd.visible = 1
    wd.DisplayAlerts = False
    save_as(load_path, save_path, wd)

