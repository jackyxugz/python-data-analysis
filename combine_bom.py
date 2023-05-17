# coding=utf-8

import sys
import os
import pandas as pd
import numpy as np
import time
import os.path
import xlrd
import xlwt

import tabulate

import win32api
import win32ui
import win32con


def get_files_list(rootdir, filekey_list):
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
            _files.extend(get_files_list(path, filekey_list))
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


def get_files_df(rootdir, filekey):
    filelist = get_files_list(rootdir, filekey)
    # print(filelist)
    mySeries = pd.Series(filelist)
    df = pd.DataFrame(mySeries)
    df.columns = ["filename"]
    # df["filename"] = df["filename"].apply(lambda x: x.strip().replace(" ", ""))

    # print(df.to_markdown())
    print(df)
    return df


def comb_bom(rootdir, filekey):
    df_files = get_files_df(rootdir, filekey)
    for index, file in df_files.iterrows():
        if 'df' in locals().keys():  # 如果变量已经存在
            print(file["filename"])
            dd = read_bom(file["filename"])
            dd["filename"] = file["filename"]
            print("进度表：{}/{}  文件{}，行数{}".format(index + 1, df_files.shape[0], file["filename"], dd.shape[0]))
            df = df.append(dd)
        else:
            print(file["filename"])
            df = read_bom(file["filename"])
            df["filename"] = file["filename"]
            print("进度表：{}/{}  文件{}，行数{}".format(index + 1, df_files.shape[0], file["filename"], df.shape[0]))

    return df


def read_bom(filename):
    df = pd.read_excel(filename, dtype='object')
    df["零件編號"].fillna(method="ffill", inplace=True)
    df["品牌名称"].fillna(method="ffill", inplace=True)
    df["產品名称"].fillna(method="ffill", inplace=True)
    df["条形码编号"].fillna(method="ffill", inplace=True)
    df["供应商"].fillna(method="ffill", inplace=True)
    df.fillna('', inplace=True)
    df["filename"] = filename

    print(df.to_markdown())
    return df


if __name__ == "__main__":
    print("开始:", time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()))
    filename = input("请输入文件所在的路径：")
    df = comb_bom(filename, '')
    print(df.head(10).to_markdown())
    df.to_excel(filename +'\BOM汇总结果'+time.strftime("%Y-%m-%d") + '.xlsx')
    print("finished")
