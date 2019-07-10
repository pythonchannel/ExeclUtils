# coding:utf-8

import xlwt
import xlrd

'''
这里是操作execl的工具类,以后也可以直接复用
'''


class ExeclUtils(object):

    def __init__(self, execl_name, sheet_name, row_titles):
        '''
        :param execl_name:  文件名，不需要后缀xlsx
        :param sheet_name:  execl中的sheet名
        :param row_titles:  execl中每一列的名称
        '''
        self.execl_name = u'{}.xlsx'.format(execl_name)
        self.execl_file = xlwt.Workbook()
        self.execl_sheet = self.execl_file.add_sheet(sheet_name, cell_overwrite_ok=True)
        for i in range(0, len(row_titles)):
            self.execl_sheet.write(0, i, row_titles[i])

    def write_execl(self, count, data):
        '''
        :param count:  execl文件的行数
        :param data:  要传入的一条数据
        :return: None
        '''
        for j in range(len(data)):
            self.execl_sheet.write(count, j, data[j])

        self.execl_file.save(self.execl_name)

    def read_execl(self):
        '''
        :return:  返回一个execl的二维集合
        '''
        all_data = []  # 所有的数据
        row_data = []  # 每一行数据
        data = xlrd.open_workbook(self.execl_name)  # 打开execl文件
        table = data.sheets()[0]  # 通过索引顺序获取table, 一个execl文件一般都至少有一个table
        for a in range(1, table.nrows):  # 行数据，正好要去掉第1行标题 所以从1开始
            for b in range(table.ncols):  # 列数据
                row_data.append(table.cell(a, b).value)  # 根据行与列，可以获取到每一格数据

            all_data.append(row_data)
            row_data = []  # 清空数据

        return all_data
