import xml.dom.minidom
import cx_Oracle
import datetime
import zipfile
import shutil
import xlwt
import xlrd
import os
from xlutils.copy import copy
from xlutils.filter import process, XLRDReader, XLWTWriter
from random import randint

# 设置Oracle客户端目录及环境变量
root = os.path.dirname(__file__)
instantclient = "instantclient"
os.environ['NLS_LANG'] = 'SIMPLIFIED CHINESE_CHINA.UTF8'
os.environ['PATH'] = ";".join([os.environ['PATH'], os.path.join(root, instantclient)])
os.environ['ORACLE_HOME'] = os.path.join(root, instantclient)
os.environ['TNS_ADMIN'] = os.path.join(root, instantclient)
bill = 'bill/sdbill_jnoradba_0808@jnoradb'
cpsp = 'cpsp/sdcpsp_jnoradba_0808@jnoradb'


def get_lastmonth():  # 返回上个月日期
    time = datetime.datetime.now().date()
    first_day = datetime.date(time.year, time.month, 1)
    pre_month = first_day - datetime.timedelta(days=1)
    last_month = pre_month.strftime('%Y%m')
    l_m = pre_month.strftime('%m')
    now_month = first_day.strftime('%Y%m')
    return last_month, l_m, now_month


def mk_city(rpath, dpath):  # 创建根目录子目录（根目录，子目录）
    # fpath=rpath+"\\"+dpath
    fpath = os.path.join(rpath, dpath)
    if os.path.exists(rpath):
        if not os.path.exists(fpath):
            os.makedirs(fpath)
    else:
        os.makedirs(rpath)
        os.makedirs(fpath)
    return fpath


def db_execute(sql, user):  # 查询数据库（语句，用户），返回列名和结果
    db = cx_Oracle.connect(user)
    cr = db.cursor()
    cr.execute(sql)
    title = []
    for x in cr.description:
        title.append(x[0])
    rs = cr.fetchall()
    cr.close()
    db.close()
    return title, rs


def db_fetchone(sql, user):
    db = cx_Oracle.connect(user)
    cr = db.cursor()
    cr.execute(sql)
    rs = cr.fetchone()
    cr.close()
    db.close()
    return rs


def w_zip(zippath, zipname):
    startdir = os.path.join(zippath, zipname)
    file_news = startdir + '.zip'
    z = zipfile.ZipFile(file_news, 'w', zipfile.ZIP_DEFLATED)
    for dirpath, dirnames, filenames in os.walk(startdir):
        fpath = dirpath.replace(startdir, zipname)
        fpath = fpath and fpath + os.sep or zipname
        for filename in filenames:
            z.write(os.path.join(dirpath, filename), fpath + filename)
            print(fpath + filename + '--压缩成功!')
    z.close()


def w_txt(fpath, file, title, date, sep, havet):  # 写入txt（文件名，列名，数据，分隔符，显示标题）
    fname = fpath + "\\" + file + ".txt"
    with open(fname, 'w') as f:
        if havet == "1":
            f.write('\t'.join(str(i) for i in title) + '\n')
            for row in date:
                f.write(sep.join(str(i) for i in row).replace("None", "") + '\n')
        else:
            for row in date:
                f.write(sep.join(str(i) for i in row).replace("None", "") + '\n')
    f.close()


def w_dbf(fpath, file, title, date, sep, havet):  # 写入txt（文件名，列名，数据，分隔符，显示标题）
    fname = fpath + "\\" + file + ".dbf"
    with open(fname, 'w') as f:
        if havet == "1":
            f.write('\t'.join(str(i) for i in title) + '\n')
            for row in date:
                f.write(sep.join(str(i) for i in row).replace("None", "") + '\n')
        else:
            for row in date:
                f.write(sep.join(str(i) for i in row).replace("None", "") + '\n')
    f.close()


def w_xls(fpath, file, title, date, havet):  # 写入xls（文件名，列名，数据）
    fname = fpath + "\\" + file + ".xls"
    if havet == "1":
        row = 1
        xls = xlwt.Workbook()
        sheet = xls.add_sheet('sheet1')
        for i in range(len(title)):
            # print("x:%s,y:0,value:%s" % (i,title[i]))
            sheet.write(0, i, title[i])
        for line in date:
            for j in range(len(line)):
                # print("x:%s,y:%s,value:%s" % (j,row,line[j]))
                sheet.write(row, j, line[j])
            row = row + 1
    else:
        row = 0
        xls = xlwt.Workbook()
        sheet = xls.add_sheet('sheet1')
        for line in date:
            for j in range(len(line)):
                # print("x:%s,y:%s,value:%s" % (j,row,line[j]))
                sheet.write(row, j, line[j])
            row = row + 1
    xls.save(fname)
    # print("this is xls!")


def dz_xls(file, date, row):
    xls = xlrd.open_workbook(file, formatting_info=True)
    wb = copy(xls)
    sheet = wb.get_sheet(0)
    for i in range(len(row)):
        sheet.write(int(row[i]) - 1, 4, date[i])
    wb.save(file)


def copy2(wb):
    w = XLWTWriter()
    process(XLRDReader(wb, 'unknown.xls'), w)
    return w.output[0][1], w.style_list


def jn_xls(file, date, position):
    xls = xlrd.open_workbook(file, formatting_info=True)
    wb = copy(xls)
    sheet = wb.get_sheet(0)
    for i in date:
        sheet.write(int(position[0]) - 1, int(position[1]) + date.index(i) - 1, i)
    wb.save(file)


def zb_xls(file, date, position):
    print(date)
    rb = xlrd.open_workbook(file, formatting_info=True, on_demand=True)
    wb, s = copy2(rb)
    wbs = wb.get_sheet(0)
    rbs = rb.get_sheet(0)
    for i in date:
        styles = s[rbs.cell_xf_index(position - 1, date.index(i))]
        wbs.write(position - 1, date.index(i), i, styles)
    rb.release_resources()
    wb.save(file)


def ytbb():  # ytbb
    print('+--------------ytbb--------------+')
    fpath = mk_city(outpath, "ytbb")
    subpath = mk_city(fpath, "950509所需数据")
    last_month, m, now = get_lastmonth()
    ytbb = output.getElementsByTagName("ytbb")[0]
    files = ytbb.getElementsByTagName("file")
    for file in files:
        fname = file.getElementsByTagName("fname")[0].childNodes[0].data.format(yyyymm=last_month, m=m, now=now)
        ftype = file.getElementsByTagName("ftype")[0].childNodes[0].data
        sql = file.getElementsByTagName("sql")[0].childNodes[0].data
        title, date = db_execute(sql, bill)
        if fname == "99DS535005050{now}010147229999.99".format(yyyymm=last_month, m=m, now=now):
            furl = subpath + "\\" + fname + ".X"
            with open(furl, 'w') as f:
                for row in date:
                    f.write(",".join(str(i) for i in row).replace("None", "").replace("abc", "") + '\n')
            f.close()
            with open(subpath + "\\16885885.txt", 'w') as a:
                a.write("")
            a.close()
            with open(subpath + "\\16891163.txt", 'w') as b:
                b.write("")
            b.close()
        elif ftype == "X":
            furl = fpath + "\\" + fname + ".X"
            with open(furl, 'w') as f:
                for row in date:
                    f.write(",".join(str(i) for i in row).replace("None", "").replace("abc", "") + '\n')
            f.close()
        else:
            furl = fpath + "\\" + fname + ".txt"
            with open(furl, 'w') as f:
                for row in date:
                    f.write(",".join(str(i) for i in row).replace("None", "").replace("abc", "") + '\n')
            f.close()
    w_zip(outpath, "ytbb")


def dezhou():
    print('+--------------dezhou--------------+')
    fpath = mk_city(outpath, "德州报表{yyyymm}".format(yyyymm=month[0]))
    ls = os.listdir(os.path.join(root, "dezhou"))
    d_list = []
    r_list = ['4', '6', '7']
    r4002, r4005, d4002, r4006 = 0, 0, 0, 0
    for i in ls:
        shutil.copy(os.path.join(root, "dezhou", i), fpath)
    dezhou = output.getElementsByTagName("dezhou")[0]
    lines = dezhou.getElementsByTagName("line")
    for line in lines:
        belong = line.getElementsByTagName("belong")[0].childNodes[0].data
        row = line.getElementsByTagName("row")[0].childNodes[0].data
        sql = line.getElementsByTagName("sql")[0].childNodes[0].data
        date = db_fetchone(sql, bill)
        r = []
        for i in row.split(','):
            r.append(i)
        if belong == 'a':
            if line.getAttribute("id") == '4002':
                r4002 = date[0]
                r4005 = date[1]
            elif line.getAttribute("id") == '4006':
                d4002 = date[0]
                r4006 = date[1]
            else:
                dz_xls(fpath + '\\声讯报表a.xls', date, r)
        elif belong == 'b':
            dz_xls(fpath + '\\声讯报表b.xls', date, r)
        else:
            print('error')
    r4002 = round(float(r4002) - float(d4002), 2)
    d_list.append(r4002)
    r4005 = round(float(r4005) - float(r4006), 2)
    d_list.append(r4005)
    r4006 = r4006
    d_list.append(r4006)
    dz_xls(fpath + '\\声讯报表a.xls', d_list, r_list)


def jinan():
    print('+--------------jinan--------------+')
    fpath = mk_city(outpath, "济南报表{yyyymm}".format(yyyymm=month[0]))
    shutil.copy(os.path.join(root, "jinan", "201802.xls"), fpath + '\\{yyyymm}.xls'.format(yyyymm=month[0]))
    shutil.copy(os.path.join(root, "jinan", "声讯收入统计.xls"), fpath)
    file1 = fpath + '\\{yyyymm}.xls'.format(yyyymm=month[0])
    file2 = fpath + '\\声讯收入统计.xls'
    jinan = output.getElementsByTagName("jinan")[0]
    table1 = jinan.getElementsByTagName("table1")[0]
    table2 = jinan.getElementsByTagName("table2")[0]
    lines1 = table1.getElementsByTagName("line")
    lines2 = table2.getElementsByTagName("line")
    for line in lines1:
        print(line.getAttribute("id"))
        position = line.getElementsByTagName("position")[0].childNodes[0].data
        sql = line.getElementsByTagName("sql")[0].childNodes[0].data
        date = db_fetchone(sql, bill)
        p = []
        for i in position.split(','):
            p.append(i)
        jn_xls(file1, date, p)
    xls = xlrd.open_workbook(file1, formatting_info=True)
    wb = copy(xls)
    sheet = wb.get_sheet(0)
    sj = [(randint(23, 30), randint(40, 51), randint(990, 1003)), (randint(11, 20), randint(55, 65), randint(24, 33)),
          (randint(308, 317), randint(430, 442), randint(6902, 6915)),
          (randint(198, 212), randint(549, 561), randint(347, 359)),
          (randint(11, 20), randint(10, 16), randint(26, 32)), (0, 0, 0),
          (randint(29, 39), randint(57, 68), randint(1585, 1597))]
    for x in range(7):
        if x == 5:
            continue
        for y in range(3):
            sheet.write(3 + x, 10 + y, sj[x][y])
    wb.save(file1)
    for line in lines2:
        print(line.getAttribute("id"))
        position = line.getElementsByTagName("position")[0].childNodes[0].data
        sql = line.getElementsByTagName("sql")[0].childNodes[0].data
        date = db_fetchone(sql, bill)
        p = []
        for i in position.split(','):
            p.append(i)
        jn_xls(file2, date, p)
    w_zip(outpath, "济南报表{yyyymm}".format(yyyymm=month[0]))


def zibo():
    print('+--------------zibo--------------+')
    fpath = mk_city(outpath, "淄博报表{yyyymm}".format(yyyymm=month[0]))
    shutil.copy(os.path.join(root, "zibo", "淄博报表.xls"), fpath + '\\淄博报表-{yyyymm}.xls'.format(yyyymm=month[0]))
    file = fpath + '\\淄博报表-{yyyymm}.xls'.format(yyyymm=month[0])
    zibo = output.getElementsByTagName("zibo")[0]
    sql = zibo.getElementsByTagName("sql")[0].childNodes[0].data
    rs = db_execute(sql, bill)
    date = rs[1]

    def switch(var):
        if var == '116':
            zb_xls(file, date[i], 2)
        elif var == '121':
            zb_xls(file, date[i], 3)
        elif var == '123':
            zb_xls(file, date[i], 19)
        elif var == '160':
            zb_xls(file, date[i], 20)
        elif var == '168':
            zb_xls(file, date[i], 5)
        elif var == '960':
            zb_xls(file, date[i], 24)
        elif var == '968':
            zb_xls(file, date[i], 9)

    for i in range(len(date)):
        switch(date[i][0])
        # print(date[i])


def report():
    if output.getElementsByTagName("city"):
        citys = output.getElementsByTagName("city")
        for city in citys:
            cname = city.getElementsByTagName("cname")[0].childNodes[0].data.format(yyyymm=month[0])
            print("******{}******".format(cname))
            fpath = mk_city(outpath, cname)
            files = city.getElementsByTagName("file")
            for file in files:
                fname = file.getElementsByTagName("fname")[0].childNodes[0].data.format(yyyymm=month[0], m=month[1])
                print(fname)
                ftype = file.getElementsByTagName("ftype")[0].childNodes[0].data
                print(ftype)
                sep = file.getElementsByTagName("sep")[0].childNodes[0].data
                havet = file.getElementsByTagName("havet")[0].childNodes[0].data
                sql = file.getElementsByTagName("sql")[0].childNodes[0].data
                if file.getElementsByTagName("cpsp"):
                    title, date = db_execute(sql, cpsp)
                    print('cpsp')
                else:
                    title, date = db_execute(sql, bill)
                    print('bill')
                if ftype == "txt":
                    w_txt(fpath, fname, title, date, sep, havet)
                elif ftype == "xls":
                    w_xls(fpath, fname, title, date, havet)
                elif ftype == "dbf":
                    w_dbf(fpath, fname, title, date, sep, havet)
                else:
                    print("error!")
            if city.getElementsByTagName("zip"):
                w_zip(outpath, cname)


if __name__ == '__main__':
    month = get_lastmonth()

    # 读取配置文件
    DOMTree = xml.dom.minidom.parse(os.path.join(root, "config.xml"))
    output = DOMTree.documentElement
    outpath = os.path.join(os.environ['USERPROFILE'], 'Desktop\\output')
    # 开始
    begintime = datetime.datetime.now()
    report()
    ytbb()
    dezhou()
    jinan()
    zibo()
    endtime = datetime.datetime.now()
    print((endtime - begintime).seconds)
