"""
写xls公共类
想要实现，自定义列名、格式，etc：
class wb:
    xm=column(string)
    kh=column(number)
data=[('a',1), ('b',2)]
x=wb(data)
x.write()
"""
import os
import xlwt
from xlwt.Style import XFStyle


class Field:
    def __init__(self, name, column_type):
        self.name = name
        self.column_type = column_type

    def __str__(self):
        return '<%s:%s>' % (self.__class__.__name__, self.name)
    
    
class SimpleField(Field):
    def __init__(self, name):
        super(SimpleField, self).__init__(name, XFStyle())


class DataField(Field):
    def __init__(self, name):
        style = XFStyle()
        style.num_format_str = 'YYYY-MM-DD h:mm:ss'
        super(DataField, self).__init__(name, style)


class XLSMetaClass(type):
    def __new__(cls, name, bases, attrs):
        if name == "Xls":
            return super(XLSMetaClass, cls).__new__(cls, name, bases, attrs)
        mappings = dict()
        for k, v in attrs.items():
            if isinstance(v, Field):
                print("Found mapping: %s ==> %s" % (k, v))
                mappings[k] = v
        for k in mappings.keys():
            attrs.pop(k)
        if name.endswith('_t'):
            attrs['__nontitle__'] = True
        else:
            attrs['__nontitle__'] = False
        attrs['__mappings__'] = mappings
        return super(XLSMetaClass, cls).__new__(cls, name, bases, attrs)


class Xls(metaclass=XLSMetaClass):
    def __init__(self, data_iter):
        self.data_iter = data_iter
        self.wb = xlwt.Workbook()
        self.ws = self.wb.add_sheet("sheet1")
        self.row_id = 0
        self._write()

    def _write(self):
        if not self.__nontitle__:
            title = [v.name for v in self.__mappings__.values()]
            for col in range(len(title)):
                self.ws.write(self.row_id, col, title[col])
            self.row_id += 1
        column_type = [v.column_type for v in self.__mappings__.values()]
        for cols in self.data_iter:
            for col in range(len(cols)):
                self.ws.write(self.row_id, col, cols[col], column_type[col])
            self.row_id += 1

    def save(self, name, path=""):
        file_path = os.path.join(path, name)
        self.wb.save(file_path)

    def __str__(self):
        return "<%s(%s)>" % (self.__class__.__name__, str(self.data_iter))


if __name__ == '__main__':
    class test(Xls):
        xm = SimpleField('xm')
        data = DataField('data')


    import datetime
    now = datetime.datetime.now()
    now2 = datetime.datetime(2019, 7, 1, 21, 11, 46, 267838)
    x = test([('xsha', now), ('xzhang', now2)])
    print(x.__nontitle__)
    x.save('666.xls')
