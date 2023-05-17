# coding=utf-8

import sys
import os
import pandas as pd
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
                        if filename.find(key) > 0:  # 只做文件名的过滤
                            _files.append(path)
                else:
                    _files.append(path)

    # print(_files)
    return _files


def read_excel(filename):
    print(filename)
    # try:
    if filename.find("xls")>=0:
        try:
            temp_df = pd.read_excel(filename, sheet_name=default_sheet, skiprows=default_skiprow, dtype=str)
            for column_name in temp_df.columns:
                temp_df.rename(columns={column_name:column_name.replace(" ","").replace("\n","").strip()},inplace=True)
        except Exception as e:
            # ms = win32api.MessageBox(0, "{} 文件缺少 {} 分页！！请查看原文件".format(filename,default_sheet), "提醒", win32con.MB_OK)
            dict = {"filename": filename}
            temp_df = pd.DataFrame(dict, index=[0])
    elif filename.find("csv")>=0:
        try:
            temp_df = pd.read_csv(filename, sheet_name=default_sheet, skiprows=default_skiprow, dtype=str)
            for column_name in temp_df.columns:
                temp_df.rename(columns={column_name: column_name.replace(" ", "").replace("\n", "").strip()},
                               inplace=True)
        except Exception as e:
            try:
                temp_df = pd.read_csv(filename, sheet_name=default_sheet, skiprows=default_skiprow, dtype=str, encoding="gb18030")
                for column_name in temp_df.columns:
                    temp_df.rename(columns={column_name: column_name.replace(" ", "").replace("\n", "").strip()},
                                   inplace=True)
            except Exception as e:
                # ms = win32api.MessageBox(0, "{} 文件缺少 {} 分页！！请查看原文件".format(filename, default_sheet), "提醒",win32con.MB_OK)
                dict = {"filename": filename}
                temp_df = pd.DataFrame(dict, index=[0])
    else:
        print("不是xls或者csv文件！！")
        dict = {"filename":filename}
        temp_df = pd.DataFrame(dict,index=[0])

    print(temp_df.head(1).to_markdown())
    return temp_df


def get_all_files(rootdir, filekey):
    filelist = list_all_files(rootdir, filekey)
    # print(filelist)
    mySeries = pd.Series(filelist)
    df = pd.DataFrame(mySeries)
    df.columns = ["filename"]
    # df["filename"] = df["filename"].apply(lambda x: x.strip().replace(" ", ""))

    # print(df.to_markdown())
    print(df)
    count = len(df)
    global default_count
    default_count = count
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
            print("进度表：{}/{}  文件{}，行数{}".format(index + 1, df_files.shape[0], file["filename"], df.shape[0]))

    return df


def combine_excel():
    print('请选择要合并的文件目录（请确定所有excel的表头都是相同的哦！）:')
    # filedir=""
    filedir = input()
    # myTuple = shell.SHBrowseForFolder(0, None, "", 64)
    # try:
    #     path = shell.SHGetPathFromIDList(myTuple[0])
    # except:
    #     print("你没有输入任何目录 :(")
    #     sys.exit()
    #     return

    # filedir=path.decode('ansi')
    print("你选择的路径是：", filedir)

    global default_dir
    default_dir = filedir

    print("如果表头不是第一行，请输入需要跳过表头行数")
    skiprow = int(input())

    global default_skiprow
    default_skiprow = skiprow

    print("你要跳过的表头行数：", skiprow)

    print('请输入要合并的分表！')
    sheet_name = input()

    global default_sheet
    default_sheet = sheet_name


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

    table = read_all_excel(filedir, filekey)

    # table.dropna(inplace=True)

    return table


def caiwu_xushizhang():
    df = combine_excel()
    if 'df' in locals().keys():  # 如果变量已经存在
        # print(df.head(10).to_markdown())
        print(df.head(3))
        # df.to_clipboard(index=False)
        print(f"合并的总文件数量：{default_count}")
        print("ok1")
        if len(df)>500000:
            df.to_csv(default_dir + r"\合并表格.csv")
        else:
            df.to_excel(default_dir + r"\合并表格.xlsx")

        print("生成完毕!")
        byebye = input()
    else:
        print("不好意思，什么也没有做哦 :(")
        # pyinstaller -p D:\Anaconda3\envs\duizhang -F .\xushizhang.py


def error(self,filename):
    logging.basicConfig(filename=default_dir + "\错误日志.log",
                        format=f'[%(asctime)s-%(filename)s-%(levelname)s:%(message)s:{filename}]', level=logging.DEBUG,
                        filemode='a', datefmt='%Y-%m-%d%I:%M:%S %p')

    logging.error("这是一条error信息的打印")
    # logging.info("这是一条info信息的打印")
    # logging.warning("这是一条warn信息的打印")
    # logging.debug("这是一条debug信息的打印")


if __name__ == "__main__":
    print("开始:", time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()))
    # print("abc".find("x"))

    caiwu_xushizhang()

    print("ok")

