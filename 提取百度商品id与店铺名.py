# __coding=utf8__
# /** 作者：zengyanghui **/
import re
import sys
import os

import future.backports.socketserver
import pandas as pd

# 显示所有列
pd.set_option('display.max_columns', None)
# 显示所有行
pd.set_option('display.max_rows', None)
# 设置value的显示长度为100，默认为50
pd.set_option('max_colwidth', 200)

import numpy as np
# from datetime import datetime
import datetime
import time
import os.path
import xlrd
import xlwt
import pprint
import math
import tabulate


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
                        # print("文件名:",filename)

                        key = key.replace("！", "!")

                        if key.find("!") >= 0:
                            # print("反向选择:",key)
                            if filename.find(key.replace("!", "")) >= 0:  # 此文件不要读取
                                # print("{} 不应该包含 {}，所以剔除:".format(filename,key ))
                                pass
                        elif filename.find(key) > 0:  # 只做文件名的过滤
                            _files.append(path)

                else:
                    _files.append(path)

    # print(_files)
    return _files


def read_bill(filename):
    print(filename)
    df = pd.read_excel(filename,sheet_name=None,dtype=str)
    # print(df.head(5).to_markdown())
    print(list(df))
    sheet_list = list(df)
    df1 = None
    for sheet in sheet_list:
        df = pd.read_excel(filename,sheet_name=sheet,dtype=str)
        for column_name in df.columns:
            df.rename(columns={column_name:column_name.upper()},inplace=True)
        df["sheet"] = sheet
        try:
            df = df.drop_duplicates(["商品ID","sheet"])
        except Exception as e:
            # dict1 = {"商品ID": "", "sheet": ""}
            # df = pd.DataFrame(dict1)
            # return df
            df["商品ID"]=""
        print(len(df))
        print(df.head(5).to_markdown())
        if df1 is None:
            df1 = df
        else:
            print(f"df1:\n{df1.head(1).to_markdown()}")
            print(f"df:\n{df.head(1).to_markdown()}")
            df1 = pd.concat([df1,df])
            print(len(df1))
    print(len(df1))
    print(df1.head(5).to_markdown())
    print("\n\n\n\n")
    return df1


def get_all_files(rootdir, filekey):
    filelist = list_all_files(rootdir, filekey)
    # print(filelist)
    mySeries = pd.Series(filelist)
    df = pd.DataFrame(mySeries)
    df.columns = ["filename"]
    # df["filename"] = df["filename"].apply(lambda x: x.strip().replace(" ", ""))

    # df=df[df["filename"].str.contains("货款")]

    return df


def read_all_bill(rootdir, filekey):
    df_files = get_all_files(rootdir, filekey)
    for index, file in df_files.iterrows():
        if 'df' in locals().keys():  # 如果变量已经存在
            dd = read_bill(file["filename"])
            dd["filename"] = file["filename"]
            print("进度表：{}/{}  文件{}，行数{}".format(index + 1, df_files.shape[0], file["filename"], dd.shape[0]))
            df = df.append(dd)


        else:
            df = read_bill(file["filename"])
            df["filename"] = file["filename"]
            # print(file["filename"],  df.shape[0])
            print("进度表：{}/{}  文件{}，行数{}".format(index + 1, df_files.shape[0], file["filename"], df.shape[0]))

    return df


def combine_bill():
    print('订单数据校对逻辑:')
    print('1.财务订单数据需要放在财务数据文件夹下，例如/校对数据/财务数据/...')
    print('2.导出订单数据需要放在导出数据文件夹下，例如/校对数据/导出数据/...')
    print("请输入财务订单和导出订单所在的文件夹：")
    # filedir=""
    # filedir = input()
    # myTuple = shell.SHBrowseForFolder(0, None, "", 64)
    try:
        # path = shell.SHGetPathFromIDList(myTuple[0])
        filedir = input()
    except:
        print("你没有输入任何目录 :(")
        sys.exit()
        return

    # filedir=path.decode('ansi')
    print("你选择的路径是：", filedir)

    global default_dir
    default_dir = filedir

    if len(filedir) == 0:
        print("你没有输入任何目录 :(")
        sys.exit()
        return

    print('请输入要合并的文件名称或者后缀匹配符:（比如都必须包含 序时账 三个字，那么就请输入  "序时账" ， 不输入就表示所有excel都要合并！）')
    filekey = input()

    if len(filedir) == 0:
        print("你没有输入任何关键词 :(")
        filekey = ''
        # sys.exit()
        # return

    print("你希望在'{}'目录下找到所有的包含'{}'文件，然后合并。".format(filedir, filekey))

    # table = read_all_excel(filedir, filekey)
    table = read_all_bill(filedir, filekey)
    # del table["filename"]

    # if table.shape[0] < 800000:
    #     table.to_excel(default_dir + "/处理后的账单.xlsx", index=False)
    # else:
    #     table.to_csv(default_dir + "/处理后的账单.csv", index=False)
    index = 0
    plat = os.sep.join(default_dir.split(os.sep)[-1:])
    print("第{}个表格,记录数:{}".format(index, table.shape[0]))
    print(table.head(10).to_markdown())
    # df.to_excel(r"work/合并表格_test.xlsx")
    print("账单总行数：")
    print(table.shape[0])
    for i in range(0, int(table.shape[0] / 200000) + 1):
        print("存储分页：{}  from:{} to:{}".format(i, i * 200000, (i + 1) * 200000))
        # df.iloc[i*500000:(i+1)*500000].to_csv(default_dir + "\{}_合并表格_{}.csv".format(index,i))
        table.iloc[i * 200000:(i + 1) * 200000].to_excel(default_dir + "\{}-商品ID与店铺名称关系{}.xlsx".format(plat, i), index=False)

    return table


if __name__ == "__main__":
    print("开始:", time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()))
    # print("abc".find("x"))

    # combine_excel()

    combine_bill()

    # groupby_amt()
    # math_file()
    # get_shopcode("JD","dentylactive旗舰店")

    print("ok")