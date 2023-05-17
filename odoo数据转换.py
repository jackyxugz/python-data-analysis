# __coding=utf8__
# /** 作者：zengyanghui **/

import pandas as pd
import numpy as np
from collections import Counter
import re
import os
import warnings
import time
import datetime
import sys
import math
import datetime as dt
import uuid
import tabulate
import os.path
import xlrd
import xlwt
import xlsxwriter

warnings.filterwarnings("ignore")


# import Tkinter
# import win32api
# import win32ui
# import win32con
# import win32com
# from win32com.shell import shell
# import json


# 检查目录是否存在
def mkdir(default_path):
    path1 = default_path + "/数据转换"

    isExists1 = os.path.exists(path1)
    if not isExists1:
        os.makedirs(path1)
        print(path1 + ' 创建成功')
    else:
        print(path1 + ' 目录已存在')


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


def get_product():
    file = r"D:\沙井\2019年产品.xls"
    df = pd.read_excel(file)
    for column_name in df.columns:
        df.rename(columns={column_name: column_name.replace("\n", "").replace(" ", "").strip()}, inplace=True)
    product = df[["条码", "内部参考", "产品/显示名称", "计量单位/显示名称"]]

    return product


def read_order(filename):
    print(filename)
    df = pd.read_excel(filename, dtype=str)
    if "商家编码" in df.columns:
        df.rename(columns={"商家编码": "条码"}, inplace=True)
    product = get_product()
    df = pd.merge(df, product, how="left", on="条码")

    # 根据是否存在主体进行导入模板分类
    if "主体" in df.columns:
        model = "type2"
    else:
        model = "type1"

    # 导入模板1赋值逻辑
    df1 = pd.DataFrame()
    if model == "type1":
        df["日期1"] = df["日期"].astype("datetime64[D]")
        df1["单据日期"] = pd.to_datetime(df["日期"]).dt.date
        df1["commitment_date"] = ""
        df1["客户"] = df["客户"]
        df1["发票地址"] = df["客户"]
        df1["送货地址"] = df["客户"]
        df1["客户参考"] = ""
        df1["条码"] = df["条码"]
        df1["订单行/产品"] = df["内部参考"]
        df1["订单行/说明"] = df["产品/显示名称"]
        df1["订单行/计量单位"] = df["计量单位/显示名称"]
        df1["订单行/订购数量"] = df["数量"].astype(float)
        df1["订单行/单价"] = df["单价（含税）"].astype(float)
        df1["订单行/税率"] = df["日期1"].apply(
            lambda x: "税收13％（含）" if x >= datetime.datetime.strptime("2019-04-01", "%Y-%m-%d") else "税收16％（含）")
        df1.sort_values(by=["单据日期", "客户"], inplace=True)
        df1 = pd.DataFrame(df1).reset_index()
        df1["id"] = df1.index
        df2 = df1[["id", "单据日期", "commitment_date", "客户", "发票地址", "送货地址"]]
        df2.drop_duplicates(subset=["单据日期", "客户"], inplace=True)
        df3 = df1[["id", "客户参考", "条码", "订单行/产品", "订单行/说明", "订单行/计量单位", "订单行/订购数量", "订单行/单价", "订单行/税率"]]

        dfs = pd.merge(df2, df3, how="right", on="id")
        dfs["model"] = "type1"
        del dfs["id"]
        print(len(dfs))
        print(dfs.head(20).to_markdown())
        return dfs

    elif model == "type2":
        df["日期1"] = df["支付日期"].astype("datetime64[D]")
        df1["主体公司"] = df["主体"]
        df1["进销存标识"] = ""
        df1["date_order"] = pd.to_datetime(df["支付日期"]).dt.date
        df1["commitment_date"] = ""
        df1["平台"] = df["平台"]
        df1["客户"] = df["店铺"]
        df1["发票地址"] = df["店铺"]
        df1["送货地址"] = df["店铺"]
        df1["源单据"] = ""
        df1["客户参考"] = ""
        df1["产品条码"] = df["条码"]
        df1["订单行/产品"] = df["内部参考"]
        df1["订单行/说明"] = df["产品/显示名称"]
        df1["订单行/计量单位"] = df["计量单位/显示名称"]
        df1["订单行/订购数量"] = df["实际卖出数量"].astype(float)
        df1["订单行/单价"] = df["平均价格"].astype(float)
        df1["订单行/税率"] = df["日期1"].apply(
            lambda x: "税收13％（含）" if x >= datetime.datetime.strptime("2019-04-01", "%Y-%m-%d") else "税收16％（含）")
        df1["条款和条件"] = ""
        df1["销售员"] = ""
        df1.sort_values(by=["date_order", "客户"], inplace=True)
        df1 = pd.DataFrame(df1).reset_index()
        df1["id"] = df1.index
        df2 = df1[["id", "主体公司", "进销存标识", "date_order", "commitment_date", "平台", "客户", "发票地址", "送货地址", "源单据"]]
        df2.drop_duplicates(subset=["date_order", "客户"], inplace=True)
        df3 = df1[["id", "产品条码", "订单行/产品", "订单行/说明", "订单行/计量单位", "订单行/订购数量", "订单行/单价", "订单行/税率", "条款和条件", "销售员"]]

        dfs = pd.merge(df2, df3, how="right", on="id")
        dfs["model"] = "type2"
        del dfs["id"]
        print(len(dfs))
        print(dfs.head(20).to_markdown())
        return dfs


def get_order_files(rootdir, filekey):
    filelist = list_all_files(rootdir, filekey)
    # print(filelist)
    mySeries = pd.Series(filelist)
    df = pd.DataFrame(mySeries)
    df.columns = ["filename"]
    # df["filename"] = df["filename"].apply(lambda x: x.strip().replace(" ", ""))

    # df=df[~df["filename"].str.contains("快递")]

    # print(df.to_markdown())
    # print("抽查是否还有快递！")
    # print(df[df.filename.str.contains("快递")].to_markdown())
    return df


def read_order_excel(rootdir, filekey):
    df_files = get_order_files(rootdir, filekey)
    for index, file in df_files.iterrows():
        if 'df' in locals().keys():  # 如果变量已经存在
            dd = read_order(file["filename"])
            dd["filename"] = file["filename"]
            print("进度表：{}/{}  文件{}，行数{}".format(index + 1, df_files.shape[0], file["filename"], dd.shape[0]))
            # print(dd.head(1).to_markdown())
            df = df.append(dd)

        else:
            df = read_order(file["filename"])
            df["filename"] = file["filename"]
            print("进度表：{}/{}  文件{}，行数{}".format(index + 1, df_files.shape[0], file["filename"], df.shape[0]))

    return df


def all_order():
    print('请选择要合并的文件目录（请确定所有excel的表头都是相同的哦！）:')
    # filedir=""
    # filedir = input()
    # myTuple = shell.SHBrowseForFolder(0, None, "", 64)
    # try:
    #     path = shell.SHGetPathFromIDList(myTuple[0])
    # except:
    #     print("你没有输入任何目录 :(")
    #     sys.exit()
    #     return
    filedir = input()
    # filedir = path.decode('ansi')
    print("你选择的路径是：", filedir)

    # mkdir(filedir)

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

    print("你希望在'{}'目录下找到所有的包含“{}”文件，然后合并。".format(filedir, filekey))

    table = read_order_excel(filedir, filekey)
    del table["filename"]

    # 转换后的文件导出
    print(len(table))
    if len(table) > 500000:
        table.to_csv(default_dir + "/数据转换结果.csv", index=False)
    else:
        table.to_excel(default_dir + "/数据转换结果.xlsx", index=False)


if __name__ == "__main__":
    print("开始:", time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()))
    # filename = r"/Users/maclove/Downloads/2019各平台数据原表/京东/订单/Dicora UrbanFit海外旗舰店/Dicora UrbanFit海外旗舰店201904-订单.xlsx"

    all_order()
    # test()

    print("ok")
