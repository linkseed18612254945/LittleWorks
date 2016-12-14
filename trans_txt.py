import os
import win32com.client


def save_as(load_path, save_path, wd):
    DIV = '------------------------------------------------------------------------'
    class_names = os.listdir(load_path)
    for class_name in class_names:
        class_path = load_path + '\\' + class_name
        file_names = os.listdir(class_path)
        # os.mkdir(save_path + '\\' + class_name) 生成存入文件夹下的各类文件夹
        print(class_name)
        for file_name in file_names:
            file_path = class_path + '\\' + file_name
            save_file_names = os.listdir(save_path + '\\' + class_name)
            name = file_name[:file_name.find('.')]
            # 跳过已处理的文件
            if name + '.txt' in save_file_names:
                print('processed pass')
                continue
            # 跳过js,css文件夹
            if file_name[-6:] == '_files':
                print('file pass')
                continue
            # 找出后缀为url的空网页跳过处理，记录到wrong中
            if file_name[-3:] == 'url':
                print('url pass')
                with open(r'C:\Users\Administrator\PycharmProjects\LittleWorks\wrong.txt', 'a', encoding='utf-8') as f:
                    f.write(file_name[:-4]+'\n')
                continue
            save_file_path = save_path + '\\' + class_name + '\\' + name
            print(file_path)
            # 调用API打开doc保存为txt
            doc = wd.Documents.Open(file_path)
            doc.SaveAs(save_file_path, 2)
            doc.Close()
        print(DIV)

if __name__ == '__main__':
    # 待处理文件夹
    load_path = r'C:\Users\Administrator\PycharmProjects\LittleWorks\给李琨\给李琨'
    # 保存的文件夹
    save_path = r'C:\Users\Administrator\PycharmProjects\LittleWorks\savetxt'

    wd = win32com.client.DispatchEx('Word.Application')
    wd.visible = 1
    wd.DisplayAlerts = False
    save_as(load_path, save_path, wd)
    wd.Quit()
