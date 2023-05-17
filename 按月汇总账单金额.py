# coding=utf-8

import sys
import os
import pandas as pd
import numpy as np
import time
import os.path
import xlrd
import xlwt
import logging
import tabulate
import math
import win32api
import win32ui
import win32con
import win32com


def read_excel(filename):
    print(filename)
    # try:
    if filename.find("csv")>=0:
        df = pd.read_csv(filename)
    else:
        df = pd.read_excel(filename)
        
    return df

    # print(df.head().to_markdown())
    # df["CREATED"] = df["CREATED"].astype("str")
    # df["CREATED"] = df["CREATED"].astype("datetime64[ns]")
    # df["month"] = df["CREATED"].dt.month
    # df["year"] = df["CREATED"].dt.year
    # df["date"] = df["CREATED"].dt.date
    # # df["date"] = df["date"].astype(str)
    # # df["date"] = df["date"].apply(lambda x: x[:7])
    # df["INCOME_AMOUNT"] = df["INCOME_AMOUNT"].astype(float)
    # df["EXPEND_AMOUNT"] = df["EXPEND_AMOUNT"].astype(float)
    # df["overseas_income"].fillna(0,inplace=True)
    # df["overseas_expend"].fillna(0,inplace=True)
    # df["overseas_income"] = df["overseas_income"].astype(float)
    # df["overseas_expend"] = df["overseas_expend"].astype(float)
    # print(df.head().to_markdown())
    # groupby_df = df.groupby(["year", "month", "PLATFORM", "SHOPNAME", "IS_REFUNDAMOUNT", "IS_AMOUNT"]).agg(
    #     {"INCOME_AMOUNT": "sum", "EXPEND_AMOUNT": "sum", "overseas_income": "sum", "overseas_expend": "sum"})
    # groupby_df = pd.DataFrame(groupby_df).reset_index()
    # groupby_df.columns = ["年","月份","平台","店铺名称","是否退款","是否回款","回款金额","退款金额","海外回款金额","海外退款金额"]
    # print(groupby_df.head().to_markdown())
    # groupby_df = groupby_df.loc[(groupby_df["是否回款"]==1)|(groupby_df["是否退款"]==1)]
    # del groupby_df["是否回款"]
    # del groupby_df["是否退款"]
    # 
    # print(groupby_df.head().to_markdown())
    # return groupby_df


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
            if path.find("~") < 0:  # 带~符号表示临时文件，不读取
                if len(filekey) > 0:
                    for key in filekey:
                        # print(path)
                        filename = "".join(path.split("\\")[-1:])
                        # print(filename)
                        if filename.find(key) > 0:  # 只做文件名的过滤
                            _files.append(path)
                else:
                    _files.append(path)

    # print(_files)
    return _files


def get_all_files(rootdir, filekey):
    filelist = list_all_files(rootdir, filekey)
    # print(filelist)
    mySeries = pd.Series(filelist)
    df = pd.DataFrame(mySeries)
    df.columns = ["filename"]
    # df["filename"] = df["filename"].apply(lambda x: x.strip().replace(" ", ""))

    df=df[~df["filename"].str.contains("汇总")]

    return df


def read_all_excel(rootdir, filekey):
    df_files = get_all_files(rootdir, filekey)
    for index, file in df_files.iterrows():
        if 'df' in locals().keys():  # 如果变量已经存在
            dd = read_excel(file["filename"])
            dd["filename"] = file["filename"]
            print("进度表：{}/{}  文件{}，行数{}".format(index + 1, df_files.shape[0], file["filename"], dd.shape[0]))
            df = df.append(dd)


        else:
            df = read_excel(file["filename"])
            df["filename"] = file["filename"]
            # print(file["filename"],  df.shape[0])
            print("进度表：{}/{}  文件{}，行数{}".format(index + 1, df_files.shape[0], file["filename"], df.shape[0]))

    return df


# def combine_excel():
#     print('请选择要处理的文件！')
#     # filedir=""
#     openfile = open_file()
#
#     outfile_dir = os.sep.join(openfile.split(os.sep)[:-1])
#     outfile = "".join(("".join(openfile.split(os.sep)[-1:])).split(".")[:-1])
#
#     table = read_excel(openfile)
#
#     if len(table) > 500000:
#         table.to_csv(outfile_dir + os.sep + outfile + "-汇总文件.csv", index=False)
#     else:
#         table.to_excel(outfile_dir + os.sep + outfile + "-汇总文件.xlsx", index=False)
#
#     # table.dropna(inplace=True)
#
#     return table


def combine_excel():
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

    table = read_all_excel(filedir, filekey)

    del table["filename"]
    
    # 汇总处理逻辑：
    table["CREATED"] = table["CREATED"].astype("str")
    table["CREATED"] = table["CREATED"].astype("datetime64[ns]")
    table["month"] = table["CREATED"].dt.month
    table["year"] = table["CREATED"].dt.year
    table["date"] = table["CREATED"].dt.date
    # table["date"] = table["date"].astype(str)
    # table["date"] = table["date"].apply(lambda x: x[:7])
    table["INCOME_AMOUNT"] = table["INCOME_AMOUNT"].astype(float)
    table["EXPEND_AMOUNT"] = table["EXPEND_AMOUNT"].astype(float)
    table["overseas_income"].fillna(0, inplace=True)
    table["overseas_expend"].fillna(0, inplace=True)
    table["overseas_income"] = table["overseas_income"].astype(float)
    table["overseas_expend"] = table["overseas_expend"].astype(float)
    table = table.loc[(table["IS_AMOUNT"] == 1) | (table["IS_REFUNDAMOUNT"] == 1)]
    print(table.head().to_markdown())
    groupby_df = table.groupby(["year", "month", "PLATFORM", "SHOPNAME"]).agg(
        {"INCOME_AMOUNT": "sum", "EXPEND_AMOUNT": "sum", "overseas_income": "sum", "overseas_expend": "sum"})
    groupby_df = pd.DataFrame(groupby_df).reset_index()
    print(groupby_df.head().to_markdown())

    plat_df = groupby_df.groupby(["year","month","PLATFORM"]).agg({"INCOME_AMOUNT": "sum", "EXPEND_AMOUNT": "sum", "overseas_income": "sum", "overseas_expend": "sum"})
    plat_df = pd.DataFrame(plat_df).reset_index()
    plat_df["SHOPNAME"] = "平台汇总"
    print(plat_df.head().to_markdown())

    groupby_df = pd.concat([groupby_df,plat_df])
    groupby_df.columns = ["年", "月份", "平台", "店铺名称", "回款金额", "退款金额", "海外回款金额", "海外退款金额"]
    print(groupby_df.head().to_markdown())


    if len(table) > 500000:
        groupby_df.to_csv(default_dir + "\汇总文件.csv", index=False)
    else:
        groupby_df.to_excel(default_dir + "\汇总文件.xlsx", index=False)

    return groupby_df



def open_file():
    dlg = win32ui.CreateFileDialog(1)  # 1表示打开文件对话框
    dlg.SetOFNInitialDir('D:/')  # 设置打开文件对话框中的初始显示目录
    dlg.DoModal()
    filename = dlg.GetPathName()  # 获取选择的文件名称
    print("filename=",filename)
    print("read ok")
    return filename


def error(self,filename):
    logging.basicConfig(filename=outfile_dir + "\错误日志.log",
                        format=f'[%(asctime)s-%(filename)s-%(levelname)s:%(message)s:{filename}]', level=logging.DEBUG,
                        filemode='a', datefmt='%Y-%m-%d%I:%M:%S %p')

    logging.error("这是一条error信息的打印")
    # logging.info("这是一条info信息的打印")
    # logging.warning("这是一条warn信息的打印")
    # logging.debug("这是一条debug信息的打印")


if __name__ == "__main__":
    print("开始:", time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()))

    combine_excel()

    print("ok")

