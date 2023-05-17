# coding=utf-8
# @Version: 1.0
# @Description:
# @Date: 2023/02/15
# @Author: 徐贵中

import sys
import os
import pandas as pd
import time
import os.path

from sqlalchemy import create_engine


def list_all_files(rootdir, filekey_list):
    if len(filekey_list) > 0:
        filekey_list = filekey_list.replace(",", " ")
        filekey = filekey_list.split(" ")
    else:
        filekey = ''

    _files = []
    list = os.listdir(rootdir)  # 列出文件夹下所有的目录与文件
    for i in range(0, len(list)):
        path = os.path.join(rootdir, list[i])
        if os.path.isdir(path):
            _files.extend(list_all_files(path, filekey_list))
        if os.path.isfile(path):
            if ((path.find("~") < 0) and (path.find(".DS_Store") < 0)):  # 带~符号表示临时文件，不读取
                if len(filekey) > 0:
                    for key in filekey:
                        # print(path)
                        filename = "".join(path.split("\\")[-1:])
                        # print(filename)
                        if filename.find(key) >= 0:  # 只做文件名的过滤
                            _files.append(path)
                else:
                    _files.append(path)

    # print(_files)
    return _files


def read_excel(filename):
    print(filename)
    # try:
    if filename.find("xls") >= 0:
        temp_df = pd.read_excel(filename, skiprows=default_skiprow, dtype=str)
        for column_name in temp_df.columns:
            temp_df.rename(columns={column_name: column_name.replace(" ", "").replace("\n", "").strip()}, inplace=True)
    elif filename.find("csv") >= 0:
        try:
            temp_df = pd.read_csv(filename, skiprows=default_skiprow, dtype=str)
            for column_name in temp_df.columns:
                temp_df.rename(columns={column_name: column_name.replace(" ", "").replace("\n", "").strip()},
                               inplace=True)
        except Exception as e:
            temp_df = pd.read_csv(filename, skiprows=default_skiprow, dtype=str, encoding="gb18030")
            for column_name in temp_df.columns:
                temp_df.rename(columns={column_name: column_name.replace(" ", "").replace("\n", "").strip()},
                               inplace=True)
    else:
        print("不是xls或者csv文件！！")
        dict = {"filename": filename}
        temp_df = pd.DataFrame(dict, index=[0])

    #print(temp_df.head(10).to_markdown())
    return temp_df


def get_all_files(rootdir, filekey):
    filelist = list_all_files(rootdir, filekey)
    # print(filelist)
    mySeries = pd.Series(filelist)
    df = pd.DataFrame(mySeries)
    df.columns = ["filename"]
    count = len(df)
    global default_count
    default_count = count
    return df


def read_all_excel(rootdir, filekey):
    df_files = get_all_files(rootdir, filekey)
    for index, file in df_files.iterrows():
        dd = read_excel(file["filename"])
        upload_douyinvouch(file["filename"], "alibaba")
        print("进度表：{}/{}  文件{}，行数{}".format(index + 1, df_files.shape[0], file["filename"], dd.shape[0]))


def combine_excel():
    print('请选择要合并的文件目录（请确定所有excel的表头都是相同的哦！）:')
    filedir = input()

    # filedir=path.decode('ansi')
    print("你选择的路径是：", filedir)

    print("如果表头不是第一行，请输入需要跳过表头行数")
    skiprow = int(input())

    print("你要跳过的表头行数：", skiprow)

    global default_dir
    default_dir = filedir

    global default_skiprow
    default_skiprow = skiprow

    if len(filedir) == 0:
        print("你没有输入任何目录 :(")
        sys.exit()
        return

    print('请输入要合并的文件名称或者后缀匹配符:（比如合并文件名都包含"邮政" 二个字的文件，那么就输入  "邮政" ， 不输入就表示所有excel都要合并！）')
    filekey = input()

    if len(filedir) == 0:
        print("你没有输入任何关键词 :(")
        filekey = ''
    print("你希望在'{}'目录下找到所有的包含“{}”文件，然后合并。".format(filedir, filekey))

    read_all_excel(filedir, filekey)


def upload_douyinvouch(filename,tablename):
    engine = create_engine('postgresql+psycopg2://odoo:odoo@127.0.0.1:5432/mkldata')

    if filename.find("xls")>=0:
        df = pd.read_excel(filename)
    else:
        df = pd.read_csv(filename)

    #df = df[['订单编号', '订单创建时间', '订单付款时间', '货品标题', '货品总价(元)', '运费(元)', '实付款(元)', '收货人姓名', '收货地址', '联系电话', '联系手机', '单价(元)', '数量', '单位', '物流公司运单号']]
    df = df[['订单编号', '订单创建时间', '收货人姓名', '收货地址', '联系电话', '联系手机']]
    df['订单编号'].astype(str)
    #df['订单创建时间'].astype('datetime64[ns]')
    #df['订单付款时间'].astype('datetime64[ns]')
    #df['货品总价(元)'].astype(float)
    #df['运费(元)'].astype(float)
    #df['实付款(元)'].astype(float)
    #df['联系电话'].astype(str)
    #df['联系手机'].astype(str)
    #df['单价(元)'].astype(float)
    #df['货品标题'].astype(str).replace(':', '')
    #df['货品标题'].astype(str).strip()

    #df['数量'].astype(int)
    #df['单位'].astype(str)
    #df['物流公司运单号'].astype(str)

    print(df.head(10).to_markdown())

    df.to_sql(tablename, con=engine, if_exists='append', index=False)


if __name__ == "__main__":
    print('\n执行开始\n\n开始时间：', time.strftime('%Y-%m-%d %H:%M:%S', time.localtime()))
    start = time.time()
    #try:
    combine_excel()

    # except Exception as ex:
    #     print('\n\n程序错误:')
    #     print(ex)

    print('\n**********OK**********\n')

    print('\n执行结束\n结束时间:', time.strftime('%Y-%m-%d %H:%M:%S', time.localtime()))
    end = time.time()
    print('执行用时:', '%.2f' % (end - start), '秒\n')

    print('-------如需关闭窗口，请回车-------')
    input("按任意键退出......")



