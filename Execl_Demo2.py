
from ExeclUtils import ExeclUtils


class Execl_Demo(object):

    def __init__(self):
        rows_title = [u'标题', u'时间', u'作者']
        self.execl_grid_value = []  # 存放每一格的各元素，
        self.count = 0  # 数据插入从1开始的
        self.execl_util = ExeclUtils('Python的Execl工具', '工具', rows_title)

    def write_execl(self):
        titles = ['a', 'b', 'c', 'd']
        times = ['111', '222', '333', '444']
        author = ['AAA', 'BBB', 'CCC', 'DDD']

        for i in range(0, 4):
            self.execl_grid_value.append(titles[i])
            self.execl_grid_value.append(times[i])
            self.execl_grid_value.append(author[i])

            self.count = self.count + 1
            self.execl_util.write_execl(self.count, self.execl_grid_value)
            self.execl_grid_value = []

        self.read_execl()
        pass

    def read_execl(self):
        print(self.execl_util.read_execl())
        pass


if __name__ == '__main__':
    e = Execl_Demo()
    e.write_execl()
