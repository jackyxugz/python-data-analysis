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
    if filename.find("csv")>=0:
        df = pd.read_csv(filename,skiprows=2)
    else:
        df = pd.read_excel(filename,skiprows=2)
    df1 = pd.DataFrame()
    print(df.head(1).to_markdown())

    # df1=df1[["平台","店铺名","项目","收入RMB","收入RMB","收入RMB.1","收入RMB.2","收入RMB.3","收入RMB.4","收入RMB.5","收入RMB.6","收入RMB.7","收入RMB.8","收入RMB.9","收入RMB.10","收入RMB.11"]]
    list = ["2020-01","2020-02","2020-03","2020-04","2020-05","2020-06","2020-07","2020-08","2020-09","2020-10","2020-11","2020-12"]
    index = 0
    df1=df[["平台","店铺名","项目"]]
    for col in df.columns:
        if col.find("收入RMB")>=0:
            print(col)
            col1 = "收入RMB "+list[index]
            df.rename(columns={col:col1},inplace=True)
            print(df.head(1).to_markdown())
            df1[col1]=df[col1]
            index += 1

    df1 = df1.loc[df1["项目"]=="Amount"]
    df2 = df1.T
    print(df1.head(1).to_markdown())
    print(df2.to_markdown())
    return df2


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

        # 输出按月和店铺统计金额
        # df["SHOPNAME"] = df.apply(lambda x:x["SHOPNAME"] if pd.notnull(x["SHOPNAME"]) else x["filename"],axis=1)
        # df["CREATED"] = df["CREATED"].astype("datetime64[ns]")
        # df["month"] = df["CREATED"].dt.month
        # df["year"] = df["CREATED"].dt.year
        # df["date"] = df["CREATED"].dt.date
        # # df["date"] = df["date"].astype(str)
        # # df["date"] = df["date"].apply(lambda x: x[:7])
        # df["INCOME_AMOUNT"] = df["INCOME_AMOUNT"].astype(float)
        # df["EXPEND_AMOUNT"] = df["EXPEND_AMOUNT"].astype(float)
        # groupby_df = df.groupby(["PLATFORM", "SHOPNAME", "year", "month", "IS_REFUNDAMOUNT", "IS_AMOUNT"]).agg(
        #     {"INCOME_AMOUNT": "sum", "EXPEND_AMOUNT": "sum", })
        # groupby_df = pd.DataFrame(groupby_df).reset_index()
        # groupby_df.to_excel(default_dir + "\分组统计后的账单.xlsx", index=False)

        # pagecount = math.ceil(df.shape[0] / 800000.00)
        # pagecount = "{:d}".format(pagecount)
        # print("总共需要拆分{}个文件".format(pagecount))
        # writer = pd.ExcelWriter(default_dir + os.sep +"合并表格.xlsx")
        # for x in range(int(pagecount)):
        #     from_line = x * 800000
        #     to_line = (x + 1) * 800000
        #     df1 = df[from_line:to_line]
        #     print(df1.head(5).to_markdown())
        #     print("输出文件总行数：{}".format(df1.shape[0]))
        #     sheetname = "Sheet{}".format(x + 1)
        #     # df7.to_excel(result_path + "/{}_快递单号校对结果.xlsx".format(company), sheet_name=sheetname,engine='xlsxwriter')
        #     df1.to_excel(writer, sheetname)
        # writer.save()
        # df.to_pickle("data\订单主体_支付日期合并表格.pkl")
        print("生成完毕!")
        byebye = input()
    else:
        print("不好意思，什么也没有做哦 :(")
        # pyinstaller -p D:\Anaconda3\envs\duizhang -F .\xushizhang.py

    return df

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

    # df = caiwu_xushizhang()

    # file = r"D:\沙井\财务账单\2020\抖音\汇总\合并表格.csv"
    # df = pd.read_csv(file)
    # # 输出按月和店铺统计金额
    # # df["SHOPNAME"] = df.apply(lambda x:x["SHOPNAME"] if pd.notnull(x["SHOPNAME"]) else x["filename"],axis=1)
    # df["CREATED"] = df["CREATED"].astype("datetime64[ns]")
    # df["month"] = df["CREATED"].dt.month
    # df["year"] = df["CREATED"].dt.year
    # df["date"] = df["CREATED"].dt.date
    # # df["date"] = df["date"].astype(str)
    # # df["date"] = df["date"].apply(lambda x: x[:7])
    # df["INCOME_AMOUNT"] = df["INCOME_AMOUNT"].astype(float)
    # df["EXPEND_AMOUNT"] = df["EXPEND_AMOUNT"].astype(float)
    # groupby_df = df.groupby(["PLATFORM", "SHOPNAME", "year", "month", "IS_REFUNDAMOUNT", "IS_AMOUNT"]).agg(
    #     {"INCOME_AMOUNT": "sum", "EXPEND_AMOUNT": "sum", })
    # groupby_df = pd.DataFrame(groupby_df).reset_index()
    # groupby_df.to_excel(r"D:\沙井\财务账单\2020\抖音\汇总\分组统计后的账单.xlsx", index=False)

    print("ok")

