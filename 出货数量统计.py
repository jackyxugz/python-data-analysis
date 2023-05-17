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
import win32api
import win32ui
import win32con
import win32com
warnings.filterwarnings("ignore")

def read_excel():
    filename = open_file()
    print(filename)
    out_path = os.sep.join(filename.split(os.sep)[:-1])
    df1 = pd.read_excel(filename,sheet_name="出货",dtype=str)
    for column_name in df1.columns:
        df1.rename(columns={column_name: column_name.replace("\n", "").replace(" ", "").strip()}, inplace=True)
    df2 = pd.read_excel(filename,sheet_name="已销",dtype=str)
    for column_name in df2.columns:
        df2.rename(columns={column_name: column_name.replace("\n", "").replace(" ", "").strip()}, inplace=True)
    del df2["序号"]
    df1["月份-出货号"] = df1["出货月份"] + "-" + df1["商品码"]
    df2["月份-出货号"] = df2["代销月份"] + "-" + df2["商品号"]
    df = pd.merge(df1,df2,how="outer",on="月份-出货号")
    for column_name in df.columns:
        df.rename(columns={column_name: column_name.replace("\n", "").replace(" ", "").strip()}, inplace=True)
    print(df.head().to_markdown())
    # df.to_excel(out_path + os.sep + "代销与出货-结果.xlsx")
    df["求和项:入库数量"] = df["求和项:入库数量"].astype(float)
    df["9月已销"] = df.apply(lambda x:x["销售数量"] if x["代销月份"]=="202009" else 0,axis=1)
    df["9月已销"] = df["9月已销"].astype(float)
    df["9月库存"] = df["求和项:入库数量"] - df["9月已销"]
    df["10月已销"] = df.apply(lambda x:get_oct(x["出货月份"],x["商品码"],x["代销月份"],x["商品号"],x["销售数量"],10),axis=1)

    # df["11月已销"] = df.apply(lambda x:x["销售数量"] if x["代销月份"]=="202011" else np.nan,axis=1)
    # df["12月已销"] = df.apply(lambda x:x["销售数量"] if x["代销月份"]=="202012" else np.nan,axis=1)
    # df["9月已销"] = df["9月已销"].astype(float)
    # df["10月已销"] = df["10月已销"].astype(float)
    # df["11月已销"] = df["11月已销"].astype(float)
    # df["12月已销"] = df["12月已销"].astype(float)

    # df = df[["出货月份","商品码","求和项:入库数量","求和项:供货总额（不含税）","求和项:含税总额","代销月份","商品号","9月已销","10月已销","11月已销","12月已销"]]

    # del df["月份-出货号"]
    # del df["销售数量"]
    print(df.head().to_markdown())
    df.to_excel(out_path + os.sep + "代销与出货-结果.xlsx",index=False)

def get_oct(lmonth,lsku,rmonth,rsku,qty,month):
    if month == 10:
        if lmonth == "202009":
            pass



def open_file():
    dlg = win32ui.CreateFileDialog(1)  # 1表示打开文件对话框
    dlg.SetOFNInitialDir('D:/')  # 设置打开文件对话框中的初始显示目录
    dlg.DoModal()
    filename = dlg.GetPathName()  # 获取选择的文件名称
    print("filename=",filename)
    print("read ok")
    return filename


if __name__ == "__main__":
    print("开始:", time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()))

    read_excel()

    print("OK!")
    print("结束:", time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()))