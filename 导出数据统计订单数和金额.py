# __coding=utf8__
# /** 作者：zengyanghui **/
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


# import win32api
# import win32ui
# import win32con
#
# import win32com
# from win32com.shell import shell


# # 1表示打开文件对话框
# dlg = win32ui.CreateFileDialog(1)
# # 设置打开文件对话框中的初始显示目录
# dlg.SetOFNInitialDir('E:/Python')
# # 弹出文件选择对话框
# dlg.DoModal()
# # 获取选择的文件名称
# filename = dlg.GetPathName()
# print(filename)


# xx=shell.SHGetPathFromIDList()
# print(xx)


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
            if (path.find("~") < 0) and (path.find(".DS_Store") < 0):  # 带~符号表示临时文件，不读取
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


def read_excel(filename):
    # 原代码，备份：
    # print(filename)
    if filename.find("xls") > 0:
        df = pd.read_excel(filename, dtype=str)
    else:
        df = pd.read_csv(filename, dtype=str, encoding="gb18030")
    if filename.find("TAOBAO") > 0:
        plat = "淘宝"
    elif filename.find("TMALL") > 0:
        plat = "天猫"
    elif filename.find("JD") > 0:
        plat = "京东"
    elif filename.find("PDD") > 0:
        plat = "拼多多"
    elif filename.find("DY") > 0:
        plat = "抖音"
    # elif filename.find("快手") > 0:
    #     plat = "快手"
    elif filename.find("XHS") > 0:
        plat = "小红书"
    elif filename.find("KAOLA") > 0:
        plat = "网易考拉"
    elif filename.find("YZ") > 0:
        plat = "有赞"
    else:
        plat = ""
    df["平台"] = plat

    if "开始时间" in df.columns:
        df["订单数量"] = df["订单数量"].astype(float)
        df["订单数量"] = df["订单数量"].astype(int)
        df["订单金额"] = df["订单金额"].astype(float)
        df["开始时间"] = df["开始时间"].astype(str).apply(lambda x: x.replace(".", "-"))
        df["开始时间"] = df["开始时间"].astype("datetime64[ns]")
        df["年度"] = df["开始时间"].apply(lambda x: x.year)
        df["月份"] = df["开始时间"].apply(lambda x: x.month)
        temp_df = df.groupby(["平台", "店铺名称", "年度", "月份"]).agg({"订单数量": "sum", "订单金额": "sum"})
        temp_df = pd.DataFrame(temp_df).reset_index()
        # temp_df.dropna(subset=["店铺名称"],axis=0,inplace=True)
        return temp_df
    else:
        dict = {"平台": "", "店铺名称": "", "年度": "", "月份": "", "订单数量": "", "订单金额": ""}
        temp_df = pd.DataFrame(dict, index=[0])
        # print(temp_df.dtypes)
        # print(temp_df.to_markdown())
        return temp_df

    # temp_df = pd.read_excel(filename,dtype=str)
    # print(temp_df.dtypes)
    # for column_name in temp_df.columns:
    #     temp_df.rename(columns={column_name:column_name.replace(" ","").replace("\n","").strip()},inplace=True)
    #     if column_name == "子订单编号":
    #         temp_df.rename(columns={"子订单编号":"订单号"},inplace=True)
    #     elif column_name == "订单编号":
    #         temp_df.rename(columns={"订单编号": "订单号"},inplace=True)
    #     elif column_name == "主订单编号":
    #         temp_df.rename(columns={"主订单编号": "订单号"},inplace=True)
    #     elif column_name == "主订单编号":
    #         temp_df.rename(columns={"主订单编号": "订单号"},inplace=True)
    #     elif column_name == "物流单号":
    #         temp_df.rename(columns={"物流单号": "运单号码(HWBNo.)"},inplace=True)
    #     elif column_name == "单号":
    #         temp_df.rename(columns={"单号": "运单号码(HWBNo.)"},inplace=True)
    #     elif column_name == "快递单号":
    #         temp_df.rename(columns={"快递单号": "运单号码(HWBNo.)"},inplace=True)
    #     elif column_name == "总价":
    #         temp_df.rename(columns={"总价": "订单金额"},inplace=True)
    #     elif column_name == "实收款（到付按此收费）":
    #         temp_df.rename(columns={"实收款（到付按此收费）": "订单金额"},inplace=True)
    # temp_df = temp_df.loc[:,~temp_df.columns.duplicates()]
    # print(temp_df.columns)
    # temp_df = temp_df[["运单号码(HWBNo.)","订单号","订单金额"]].copy()
    # temp_df = temp_df.loc[:, ~temp_df.columns.duplicated()]
    # temp_df = temp_df[columns].T.drop_duplicates().T
    # print(temp_df.to_markdown())
    # print(temp_df.columns)
    # print(len(temp_df))
    # return temp_df


def get_all_files(rootdir, filekey):
    filelist = list_all_files(rootdir, filekey)
    # print(filelist)
    mySeries = pd.Series(filelist)
    df = pd.DataFrame(mySeries)
    df.columns = ["filename"]
    # df["filename"] = df["filename"].apply(lambda x: x.strip().replace(" ", ""))

    df = df[~df["filename"].str.contains("账单")]
    df = df[~df["filename"].str.contains("订单")]
    df = df[~df["filename"].str.contains("回款")]
    # df=df[~df["filename"].str.contains("账单")]
    df = df[df["filename"].str.contains("总表")]

    # print(df.to_markdown())
    # print("抽查是否还有快递！")
    # print(df[df.filename.str.contains("快递")].to_markdown())
    return df


# def read_all_excel(rootdir, filekey):
#     df_files = get_all_files(rootdir, filekey)
#     for index, file in df_files.iterrows():
#         if 'df' in locals().keys():  # 如果变量已经存在
#             dd = read_excel(file["filename"])
#             dd["filename"] = file["filename"]
#             print("进度表：{}/{}  文件{}，行数{}".format(index + 1, df_files.shape[0], file["filename"], dd.shape[0]))
#             df = df.append(dd)
#         else:
#             df = read_excel(file["filename"])
#             df["filename"] = file["filename"]
#             print("进度表：{}/{}  文件{}，行数{}".format(index + 1, df_files.shape[0], file["filename"], df.shape[0]))
#
#     return df

def get_amount(filename):
    df = pd.read_excel(r"/Users/maclove/PycharmProjects/pythonConda/data/文件分类1.xlsx", sheet_name="Sheet2")
    # df = pd.read_excel(filename)
    # print(df.to_markdown())
    for index, row in df.iterrows():
        # print(filename)
        # print(index)
        # print(row)
        if filename.find(row["平台"]) >= 0:
            # print(row["平台"])
            try:
                amount_column = row["金额字段"]
                # print("文件名:", filename, " 金额字段为：", amount_column)
                if filename.find("小红书") > 0:
                    tempdb = pd.read_excel(filename, sheet_name="商品销售")
                else:
                    tempdb = pd.read_excel(filename)
                    if filename.find("快手") > 0:
                        tempdb = tempdb.apply(lambda x: x.astype(str).str.replace("¥", ""))
                        if "实付款" in tempdb.columns:
                            tempdb["实付款"] = tempdb["实付款"].astype(float)
                        elif "实付款(元)" in tempdb.columns:
                            tempdb["实付款(元)"] = tempdb["实付款(元)"].astype(float)
                    else:
                        pass
                # print(tempdb.head(1).to_markdown())
                if amount_column.find(",") > 0:
                    # print(amount_column," is in  ",tempdb.columns)
                    amount_columns = amount_column.split(",")
                    for acl in amount_columns:
                        if acl in tempdb.columns:
                            return tempdb[[acl]].sum()
                else:
                    return tempdb[[amount_column]].sum()
            except Exception as e:
                print("没有找到金额字段！", filename)
                return 0
    print("没有找到平台！", filename)
    return 0


def read_all_excel(rootdir, filekey):
    df_files = get_all_files(rootdir, filekey)
    for index, file in df_files.iterrows():
        if 'df' in locals().keys():  # 如果变量已经存在
            dd = read_excel(file["filename"])
            dd["filename"] = file["filename"]
            print("进度表：{}/{}  文件{}，行数{}".format(index + 1, df_files.shape[0], file["filename"], dd.shape[0]))

            df = df.append(dd)
            # hangshu= dd.shape[0]
            # df = df.append( file["filename"],hangshu]  )
            # print(file["filename"],hangshu)

            # print(file["filename"] )
            # amount=get_amount(file["filename"])
            # print(file["filename"], dd.shape[0], amount)

        else:
            df = read_excel(file["filename"])
            df["filename"] = file["filename"]
            # print(file["filename"],  df.shape[0])
            print("进度表：{}/{}  文件{}，行数{}".format(index + 1, df_files.shape[0], file["filename"], df.shape[0]))

    return df


def combine_excel():
    print('请选择要合并的文件目录（请确定所有excel的表头都是相同的哦！）:')
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

    print("你希望在'{}'目录下找到所有的包含“{}”文件，然后合并。".format(filedir, filekey))

    table = read_all_excel(filedir, filekey)
    table.drop_duplicates(inplace=True)

    if len(table) > 500000:
        table.to_csv("data/导出订单的数量和金额.csv", index=False)
    else:
        table.to_excel("data/导出订单的数量和金额.xlsx", index=False)

    return table


def caiwu_xushizhang():
    df = combine_excel()
    if 'df' in locals().keys():  # 如果变量已经存在
        # print(df.head(10).to_markdown())
        print(df.head(3))
        # df.to_clipboard(index=False)
        print("ok1")
        # if len(df)>500000:
        # df.to_csv(default_dir + r"\合并表格.csv")
        # else:
        # df.to_excel(default_dir + r"\合并表格.xlsx")
        # print("生成完毕，现在关闭吗？yes/no")
        # byebye = input()
        # print('bybye:', byebye)
    else:
        print("不好意思，什么也没有做哦 :(")
        # pyinstaller -p D:\Anaconda3\envs\duizhang -F .\xushizhang.py


if __name__ == "__main__":
    print("开始:", time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()))
    # print("abc".find("x"))

    # caiwu_xushizhang()

    combine_excel()

    print("ok")
