### 封装了一个Execl的工具类， 手把手教你学会封装

在程序设计中永远有一个思想就是 **write once run anywhere！**

封装思想在我们编程工作中是非常重要的，有的人工作了好多年，还不会如何封装代码，写出来的代码可读性与可维护性极差，跟他们一个做项目是非常累的，但跟大牛合作，他们写的工具类会写得非常好，你只需要按工具类的要求传入数据，它就会给你返回结果. 今天我们就来学习一把如何封装.


今天用一个小的案例来教大家学会封装！

封装的精髓我是这样理解的: 

**这里有一个黑匣子，按黑匣子的指令，输入它需要的东西，然后得到我们想要的东西，具体黑匣子做了什么，不用关心。 我们封装任何东西都是这个道理，把握好输入输出这两个点，一切就好办了!**

输入：指我们外部需要提供什么样的参数给这个工具类
输出：我们通过工具类得到想返回的东西.
封装，但不要过度封装，区分那些是可变的，那些是不变的，把可变把放在外部作为输入参数传入.

理解了上面这些话，你以后封装任何东西都不在话下了.

今天我们来做一个Execl的封装，其实在Python中操作Execl还是比较频繁的，所以如果能把这些execl功能封装一下，就比较好办了，

**创建一个Execl对象**

```python
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

```

上面我创建了一个execl,在execl中，文件名，sheet名，与row_titles都是可变的，所以我把这些东西作为参数输入进来.

**把数据写入到execl中**

我只需要把行号，以及每行的数据传入进来，然后保存就行了. 代码很简单，如下

```python

    def write_execl(self, count, data):
        '''
        :param count:  execl文件的行数
        :param data:  要传入的一条数据
        :return: None
        '''
        for j in range(len(data)):
            self.execl_sheet.write(count, j, data[j])

        self.execl_file.save(self.execl_name)


```
数据写入到execl中，那些是变化的呢, 每行，每列，以及每个单元格等数据都是变化的，我传入count就是让每一行变化,data是每一列的数据，这样就好办了，于是所有的数据都可以对号入座.

**读取execl文件: **

我读取execl文件，我只需要输入文件名称，就给我返回数据，这里我把每行数据打包成一个集合，再把所有的集合组成一个新的集合返回.然后我们就可以直接到数据

```python

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

```

这个返回的集合是集合中的集合，大家直接取数据就行了.


这个execl工具已经封装好了，那么我们如何用呢？ 代码很简单

```python

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

~~~~~~~~~~~~~~运行结果~~~~~~~~~~~~~~~~~~~~~~~~~~
[['a', '111', 'AAA'], ['b', '222', 'BBB'], ['c', '333', 'CCC'], ['d', '444', 'DDD']]

```

是不是很简单，自己也可以尝试着去封装一些工具类.

### [点击进入技术交流群](http://https://mp.weixin.qq.com/s/3WVnQTOgu66FDg8X-65VvQ)
