import os
import datetime
import struct


class Field:
    def __init__(self, name, column_type, size):
        self.name = name
        self.column_type = column_type
        self.size = size

    def __str__(self):
        return '<%s:%s(%s)>' % (self.__class__.__name__, self.name, str(self.size))


class StringField(Field):
    def __init__(self, name, size):
        super(StringField, self).__init__(name, 'C', size)


class DateField(Field):
    def __init__(self, name, fmt="%Y%m%d%H%M%S"):
        self.fmt = fmt
        super(DateField, self).__init__(name, 'C', len(self.fmt)+2)


class DBFMetaClass(type):
    def __new__(cls, name, bases, attrs):
        if name == "Dbf":
            return super(DBFMetaClass, cls).__new__(cls, name, bases, attrs)
        mappings = dict()
        for k, v in attrs.items():
            if isinstance(v, Field):
                print("Found mapping: %s ==> %s" % (k, v))
                mappings[k] = v
        for k in mappings.keys():
            attrs.pop(k)
        attrs['__mappings__'] = mappings
        return super(DBFMetaClass, cls).__new__(cls, name, bases, attrs)


class Dbf(metaclass=DBFMetaClass):
    def __init__(self, data_iter):
        self.fields = [(v.name, v.column_type, v.size) for v in self.__mappings__.values()]
        self.fields_name = [field[0] for field in self.fields]
        self.actions={
            'C': lambda val, size: bytes(str(val), 'gbk')[:size].ljust(size) or b' '*size,
            'D': lambda val, fmt: bytes(val.strftime(fmt), 'gbk'),
        }
        self.recbody = b''
        self.data_iter = data_iter
        self.now = datetime.date.today()
        self._set_header()
        self._set_record()

    def _set_header(self):
        # HEADER_FORMAT = b'<B3BLHH17xB2x'
        HEADER_FORMAT = b'<B3BLHH20x'
        # FIELD_DESCRIPTION_FORMAT = b'<11sc4xBB14x'
        FIELD_DESCRIPTION_FORMAT = b'<11sc4xB15x'
        year, month, day = self.now.year - 2000, self.now.month, self.now.day
        numrec = len(self.data_iter)
        lenheader = len(self.fields) * 32 + 33
        recsize = sum([field[2] for field in self.fields]) + 1
        self.header = struct.pack(HEADER_FORMAT, 0x03, year, month, day,
                                  numrec, lenheader, recsize)
        for name, typ, size in self.fields:
            name = bytes(name, 'gbk').ljust(11, b'\x00')
            fld = struct.pack(FIELD_DESCRIPTION_FORMAT, name, bytes(typ, 'gbk'), size)
            self.header += fld
        self.header += b'\x0d'

    def _set_record(self):
        converters = [self.actions['D'] if hasattr(v, 'fmt') else self.actions['C']
                       for v in self.__mappings__.values()]
        for record in self.data_iter:
            rec = list(map(lambda a, b: (a, b.fmt) if hasattr(b, 'fmt') else (a, b.size),
                           record, [v for v in self.__mappings__.values()]))
            raw_rec = b' '+b''.join(map(lambda a, b: a(*b), converters, rec))
            self.recbody += raw_rec
        self.recbody += b'\x1a'

    def save(self, name, path=""):
        file_path = os.path.join(path, name)
        with open(file_path, 'wb') as f:
            f.write(self.header)
            f.write(self.recbody)

    def __str__(self):
        return "<%s(%s)>" % (self.__class__.__name__, str(self.data_iter))


if __name__ == '__main__':
    class test(Dbf):
        a = StringField('测试1', 8)
        b = DateField('测试2','%Y-%m-%d %H:%M:%S')

    records=[('123',datetime.datetime(2019,7,8,13,14,15)), ('789', datetime.datetime(1234,5,6,7,8,9))]
    x=test(records)
    x.save('test.dbf', '/home/xsha/桌面')
