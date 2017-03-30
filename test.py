import pandas as pd



class ExcelPanda:
    """Processing Excel by pandas"""
    def __init__(self, excel_path):
        self.df = pd.read_excel(excel_path)
        self.column_num = len(self.column_names())
        self.row_num = len(self.df.index)

    def column_names(self):
        """ 返回表中的所以列名 """
        return self.df.columns

    def get_example(self, num=5):
        """ 返回表中的前n行数据作为样例，默认为5 """
        return self.df.head(num)

    def get_columns(self, *args):
        """ 返回指定的某些列数据，参数为列名称 """
        try:
            return self.df[list(args)]
        except KeyError:
            raise KeyError('请输入正确的列名')

    def get_rows(self, start=2, end=(), df=False):
        """ 返回指定的某些行数据,默认从第一行开始"""
        if start > 1 and end <= self.row_num + 1:
            if isinstance(end, tuple):
                end = self.row_num + 1
            if isinstance(df, pd.DataFrame):
                return df[start - 2: end - 1]
            return self.df[start - 2: end - 1]
        else:
            raise IndexError('请输入表中对应的正确的行号')


    def get_range(self, start_row, end_row, *args):
        """返回指定行和列的数据"""
        return self.get_rows(start_row, end_row, self.df[list(args)])

    def set_value(self, row, column, value):
        """设定指定位置的值"""
        self.df.at[row - 2, column] = value
        return self

    def column_iter(self, col_name):
        """返回某列的迭代器"""
        return iter(self.df[col_name].values)

    def __repr__(self):
        return self.df.__repr__()

if __name__ == '__main__':
    df = pd.read_excel(r'C:\Users\51694\PycharmProjects\LittleWorks\服务区域表.xlsx')
    index = 1
    col_name = '服务区域'
    all_parts = []
    for i in df[col_name].values:
        if index > 10:
            break
        areas = i.split('；')
        if len(areas) > 1:
            part = df[:index].copy()
            part.at[index, col_name] = areas[0]
            for area in areas[1:]:
                new_row = part.tail(1)
                new_row.at[0, col_name] = area
                print(new_row)
            # change_row = ex_part.tail(1)
            # ex_part.at[index, col_name] = areas[0]
            # print(ex_part)
            # print(change_row)
        index += 1
