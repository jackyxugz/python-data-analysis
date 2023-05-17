# __coding=utf8__
# /** 作者：zengyanghui **/
import shutil
import zipfile

import pandas as pd
import numpy as np
from collections import Counter
import os
# import Tkinter
import tabulate
import math
import time
import sys


def merge_excel():
    df1 = pd.read_excel(
        r"/Users/maclove/Downloads/2019各平台数据原表/网易考拉/网易考拉DentylActive官方旗舰店/网易考拉DentylActive官方旗舰店201907-订单.xlsx",
        dtype=str)
    df2 = pd.read_csv(r"/Users/maclove/Downloads/2019各平台数据原表/数据转换/全平台_数据转换结果.csv", dtype=str)
    df2 = pd.DataFrame(df2)
    print(df2.head(1).to_markdown())
    df1 = df1[["订单号", "下单时间", "数量", "订单实付金额", "商品条形码"]]
    df1["订单实付金额"] = df1["订单实付金额"].astype(float)
    df1.rename(columns={"订单号": "订单编号"}, inplace=True)
    df2 = df2[((df2["订单时间"].str.contains("2019-07")) & (df2["平台"].str.contains("网易考拉")))]
    print(len(df2))
    print(df2.head(1).to_markdown())
    # df2 = df2.drop_duplicates(subset=["订单编号","出现序号","总序号","商家编码"],keep="first",inplace=True)
    df2 = df2.drop_duplicates()
    print(len(df2))
    print(df2.head(1).to_markdown())
    df2 = df2[["订单编号", "订单时间", "购买数量", "销售金额", "商家编码"]]
    df2["销售金额"] = df2["销售金额"].astype(float)
    # df2 = df2[((df2["订单时间"].str.contains("2019-09")) & (df2["平台"].str.contains("网易考拉")))]

    df1["数据"] = "原数据"
    df2["数据"] = "转换后"

    dfs = pd.merge(df1, df2, how="left", on="订单编号")
    dfs["销售金额"] = dfs["销售金额"].astype(float)
    print(df1["订单实付金额"].sum())
    print(df2["销售金额"].sum())
    print(dfs["销售金额"].sum())

    print(f"原数据行数:{len(df1)} \n转换后行数:{len(df2)} \nmerge后行数:{len(dfs)}")
    print(dfs.head(5).to_markdown())
    dfs.to_excel("data/对比文件.xlsx", index=False)


def read_all_sheet():
    file = r"/Users/maclove/Downloads/2019各平台数据原表/天猫、淘宝/肖琼/天猫/mades海外旗舰店/201906.xlsx"
    df = pd.read_excel(file, sheet_name=None, dtype=str)
    # print(df.head(5).to_markdown())
    print(list(df))
    sheet_list = list(df)
    df1 = None
    for sheet in sheet_list:
        df = pd.read_excel(file, sheet_name=sheet, dtype=str)
        print(len(df))
        print(df.head(5).to_markdown())
        if df1 is None:
            df1 = df
        else:
            print(f"df1:\n{df1.head(1).to_markdown()}")
            print(f"df:\n{df.head(1).to_markdown()}")
            dfs = pd.merge(df1, df, how="left", on="订单编号")
            print(len(dfs))
    print(len(dfs))
    print(dfs.head(5).to_markdown())
    dfs.to_excel("data/读取所有sheet.xlsx")


def cover(filename):
    # df = pd.read_pickle("data/全平台数据.pkl")
    # # df = pd.DataFrame(df)
    # # print(df.head(5).to_markdown())
    # print(len(df))
    # df = df[~df["平台"].str.contains("小红书")]
    # print(len(df))
    # if "小红书" in default_dir:
    print(filename)
    # df = pd.read_excel(filename, dtype=str)
    # # 给文件增加平台字段
    # df["平台"] = ""
    if "xls" in filename:
        df = pd.read_excel(filename, dtype=str)
    else:
        df = pd.read_csv(filename, skiprows=4, keep_default_na=False, dtype=str, encoding="gb18030")
    # print(df.head(5).to_markdown())
    plat = os.path.basename(filename)
    if ((filename.find("淘宝") > 0) and (filename.count("淘宝") < filename.count("天猫"))):
        plat = "淘宝"
        if filename.find("账单") > 0:
            df["平台"] = plat
            if filename.find("淘宝") > 0:
                if df.shape[0] > 0:
                    df["订单编号"] = df["Partner_transaction_id"]
                    df["退款金额"] = 0
                    df["退款金额（外币）"] = df["Refund"]
                    df["是否回款"] = "是"
                    df["结算方式"] = ""
                    df["回款金额"] = 0
                    df["回款日期"] = df["Payment_time"]
                    df["回款差额"] = 0
                    df["回款名称"] = "".join(filename.split("/")[-1:])
                    df["回款币种"] = df["Currency"]
                    df["回款汇率"] = df["Rate"]
                    df["回款金额（外币）"] = df["Amount"]
                    df["税金"] = 0.019
                    df["财务费用"] = 2.2
                    df = df[["订单编号", "退款金额", "退款金额（外币）", "是否回款", "结算方式", "回款金额", "回款日期", "回款差额", "回款名称", "回款币种", "回款汇率",
                             "回款金额（外币）", "税金", "财务费用"]]
                    df = df[~df["退款金额（外币）"].str.contains(0) & ~df["回款金额（外币）"].str.contains(0)]
                    print("淘宝-海外-有回款")
                    print(df.head(5).to_markdown())
                else:
                    dict = {"订单编号": "", "退款金额": "", "退款金额（外币）": "", "是否回款": "", "结算方式": "", "回款金额": "", "回款日期": "",
                            "回款差额": "", "回款名称": "", "回款币种": "", "回款汇率": "", "回款金额（外币）": "", "税金": "", "财务费用": ""}
                    df = pd.DataFrame(dict, index=[0])
                    # df = df[["订单编号", "退款金额", "退款金额（外币）", "是否回款", "结算方式", "回款金额", "回款日期", "回款差额", "回款名称", "回款币种", "回款汇率",
                    #          "回款金额（外币）", "税金", "财务费用"]]
                    print("淘宝-海外-无回款")
                    print(df.head(5).to_markdown())
            else:
                if "业务基础订单号" not in df.columns and "商户订单号" in df.columns:
                    print("无业务基础订单号，但有商户订单号")
                    if ((df["商户订单号"].shape[0] > 0) and (df[df["商户订单号"].str.contains("T200P")].shape[0] > 0)):
                        print("商户订单号行数>0，且包含T200P的数据")
                        df = df[df["商户订单号"].str.contains("T200P")]
                        df["业务基础订单号"] = df["商户订单号"].apply(lambda x: x.replace("T200P", ""))
                    else:
                        print("商户订单号行数=0,则业务基础订单号取空")
                        df["业务基础订单号"] = ""
                if df["业务基础订单号"].shape[0] > 0:
                    print("业务基础订单号行数>0，订单编号取业务基础订单号")
                    df["订单编号"] = df["业务基础订单号"]
                    if "业务描述" not in df.columns:
                        # print("没有业务描述的列")
                        if "业务类型" in df.columns:
                            # if ((df["业务类型"].shape[0] > 0) and (df[df["业务类型"].str.contains("退款")].shape[0] > 0)):
                            print("有业务类型的列")
                            df.loc[df["业务类型"].str.contains("退款"), "退款金额"] = df["支出金额（-元）"]
                            df.loc[~df["业务类型"].str.contains("退款"), "退款金额"] = 0
                        else:
                            print("没有业务类型的列")
                            df["退款金额"] = 0
                    else:
                        print("有业务描述的列")
                        df.loc[df["业务描述"].str.contains("退款"), "退款金额"] = df["支出金额（-元）"]
                        df.loc[~df["业务描述"].str.contains("退款"), "退款金额"] = 0
                    df["退款金额（外币）"] = 0
                    df["是否回款"] = "是"
                    df["结算方式"] = ""
                    df["回款金额"] = 0
                    if "业务描述" not in df.columns:
                        # if ((df["业务类型"].shape[0] > 0) and (df[df["业务类型"].str.contains("交易付款")].shape[0] > 0)):
                        if "业务类型" in df.columns:
                            df.loc[df["业务类型"].str.contains("交易付款"), "回款金额"] = df["收入金额（+元）"]
                            df.loc[~df["业务类型"].str.contains("交易付款"), "回款金额"] = 0
                        else:
                            df["回款金额"] = 0
                    else:
                        df.loc[df["业务描述"].str.contains("交易收款"), "回款金额"] = df["收入金额（+元）"]
                        df.loc[~df["业务描述"].str.contains("交易付款"), "回款金额"] = 0
                    df["回款日期"] = df["发生时间"]
                    df["回款差额"] = 0
                    df["回款名称"] = "".join(filename.split("/")[-1:])
                    df["回款币种"] = "RMB"
                    df["回款汇率"] = 0
                    df["回款金额（外币）"] = 0
                    df["税金"] = 0
                    df["财务费用"] = 0
                    df = df[["订单编号", "退款金额", "退款金额（外币）", "是否回款", "结算方式", "回款金额", "回款日期", "回款差额", "回款名称", "回款币种", "回款汇率",
                             "回款金额（外币）", "税金", "财务费用"]]
                    # df = df[((df["回款金额"].astype(float)>0) & (df["退款金额"].astype(float)>0))]
                    print("淘宝-国内-有回款")
                    print(df.head(5).to_markdown())
                else:
                    dict = {"订单编号": "", "退款金额": "", "退款金额（外币）": "", "是否回款": "", "结算方式": "", "回款金额": "", "回款日期": "",
                            "回款差额": "", "回款名称": "", "回款币种": "", "回款汇率": "", "回款金额（外币）": "", "税金": "", "财务费用": ""}
                    df = pd.DataFrame(dict, index=[0])
                    # df = df[["订单编号", "退款金额", "退款金额（外币）", "是否回款", "结算方式", "回款金额", "回款日期", "回款差额", "回款名称", "回款币种", "回款汇率",
                    #          "回款金额（外币）", "税金", "财务费用"]]
                    print("淘宝-国内-无回款")
                    print(df.head(5).to_markdown())
        else:
            print("没有账单文件")

    elif ((filename.find("淘宝") > 0) and (filename.count("淘宝") < filename.count("天猫"))):
        plat = "天猫"
        if filename.find("账单") > 0:
            df["平台"] = plat
            if filename.find("海外") > 0:
                if df.shape[0] > 0:
                    df["订单编号"] = df["Partner_transaction_id"]
                    df["退款金额"] = 0
                    df["退款金额（外币）"] = df["Refund"]
                    df["是否回款"] = "是"
                    df["结算方式"] = ""
                    df["回款金额"] = 0
                    df["回款日期"] = df["Payment_time"]
                    df["回款差额"] = 0
                    df["回款名称"] = "".join(filename.split("/")[-1:])
                    df["回款币种"] = df["Currency"]
                    df["回款汇率"] = df["Rate"]
                    df["回款金额（外币）"] = df["Amount"]
                    df["税金"] = 0.019
                    df["财务费用"] = 2.2
                    df = df[["订单编号", "退款金额", "退款金额（外币）", "是否回款", "结算方式", "回款金额", "回款日期", "回款差额", "回款名称", "回款币种", "回款汇率",
                             "回款金额（外币）", "税金", "财务费用"]]
                    df = df[~df["退款金额（外币）"].str.contains(0) & ~df["回款金额（外币）"].str.contains(0)]
                    print("天猫-海外-有回款")
                    print(df.head(5).to_markdown())
                else:
                    dict = {"订单编号": "", "退款金额": "", "退款金额（外币）": "", "是否回款": "", "结算方式": "", "回款金额": "", "回款日期": "",
                            "回款差额": "", "回款名称": "", "回款币种": "", "回款汇率": "", "回款金额（外币）": "", "税金": "", "财务费用": ""}
                    df = pd.DataFrame(dict, index=[0])
                    # df = df[["订单编号","退款金额","退款金额（外币）","是否回款","结算方式","回款金额","回款日期","回款差额","回款名称","回款币种","回款汇率","回款金额（外币）","税金","财务费用"]]
                    print("天猫-海外-无回款")
                    print(df.head(5).to_markdown())
            else:
                if "业务基础订单号" not in df.columns and "商户订单号" in df.columns:
                    print("无业务基础订单号，但有商户订单号")
                    if ((df["商户订单号"].shape[0] > 0) and (df[df["商户订单号"].str.contains("T200P")].shape[0] > 0)):
                        print("商户订单号行数>0，且包含T200P的数据")
                        df = df[df["商户订单号"].str.contains("T200P")]
                        df["业务基础订单号"] = df["商户订单号"].apply(lambda x: x.replace("T200P", ""))
                    else:
                        print("商户订单号行数=0,则业务基础订单号取空")
                        df["业务基础订单号"] = ""
                if df["业务基础订单号"].shape[0] > 0:
                    print("业务基础订单号行数>0，订单编号取业务基础订单号")
                    df["订单编号"] = df["业务基础订单号"]
                    if "业务描述" not in df.columns:
                        # print("没有业务描述的列")
                        if "业务类型" in df.columns:
                            # if ((df["业务类型"].shape[0] > 0) and (df[df["业务类型"].str.contains("退款")].shape[0] > 0)):
                            print("有业务类型的列")
                            df.loc[df["业务类型"].str.contains("退款"), "退款金额"] = df["支出金额（-元）"]
                            df.loc[~df["业务类型"].str.contains("退款"), "退款金额"] = 0
                        else:
                            print("没有业务类型的列")
                            df["退款金额"] = 0
                    else:
                        print("有业务描述的列")
                        df.loc[df["业务描述"].str.contains("退款"), "退款金额"] = df["支出金额（-元）"]
                        df.loc[~df["业务描述"].str.contains("退款"), "退款金额"] = 0
                    df["退款金额（外币）"] = 0
                    df["是否回款"] = "是"
                    df["结算方式"] = ""
                    df["回款金额"] = 0
                    if "业务描述" not in df.columns:
                        # if ((df["业务类型"].shape[0] > 0) and (df[df["业务类型"].str.contains("交易付款")].shape[0] > 0)):
                        if "业务类型" in df.columns:
                            df.loc[df["业务类型"].str.contains("交易付款"), "回款金额"] = df["收入金额（+元）"]
                            df.loc[~df["业务类型"].str.contains("交易付款"), "回款金额"] = 0
                        else:
                            df["回款金额"] = 0
                    else:
                        df.loc[df["业务描述"].str.contains("交易收款"), "回款金额"] = df["收入金额（+元）"]
                        df.loc[~df["业务描述"].str.contains("交易付款"), "回款金额"] = 0
                    df["回款日期"] = df["发生时间"]
                    df["回款差额"] = 0
                    df["回款名称"] = "".join(filename.split("/")[-1:])
                    df["回款币种"] = "RMB"
                    df["回款汇率"] = 0
                    df["回款金额（外币）"] = 0
                    df["税金"] = 0
                    df["财务费用"] = 0
                    df = df[["订单编号", "退款金额", "退款金额（外币）", "是否回款", "结算方式", "回款金额", "回款日期", "回款差额", "回款名称", "回款币种", "回款汇率",
                             "回款金额（外币）", "税金", "财务费用"]]
                    # df = df[((df["回款金额"].astype(float)>0) & (df["退款金额"].astype(float)>0))]
                    print("天猫-国内-有回款")
                    print(df.head(5).to_markdown())
                else:
                    dict = {"订单编号": "", "退款金额": "", "退款金额（外币）": "", "是否回款": "", "结算方式": "", "回款金额": "", "回款日期": "",
                            "回款差额": "", "回款名称": "", "回款币种": "", "回款汇率": "", "回款金额（外币）": "", "税金": "", "财务费用": ""}
                    df = pd.DataFrame(dict, index=[0])
                    # df = df[["订单编号","退款金额","退款金额（外币）","是否回款","结算方式","回款金额","回款日期","回款差额","回款名称","回款币种","回款汇率","回款金额（外币）","税金","财务费用"]]
                    print("天猫-国内-无回款")
                    print(df.head(5).to_markdown())
        else:
            print("没有账单文件")

    # elif filename.find("京东") > 0:
    #     plat = "京东"
    # elif filename.find("拼多多") > 0:
    #     plat = "拼多多"
    # elif filename.find("抖音") > 0:
    #     plat = "抖音"
    # elif filename.find("快手") > 0:
    #     plat = "快手"
    # elif filename.find("小红书") > 0:
    #     plat = "小红书"
    # elif filename.find("考拉") > 0:
    #     plat = "网易考拉"
    # elif filename.find("有赞") > 0:
    #     plat = "有赞"

    return df


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


def get_all_files(rootdir, filekey):
    filelist = list_all_files(rootdir, filekey)
    # print(filelist)
    mySeries = pd.Series(filelist)
    df = pd.DataFrame(mySeries)
    df.columns = ["filename"]
    # df["filename"] = df["filename"].apply(lambda x: x.strip().replace(" ", ""))
    # if "zip" in df["filename"]:
    #     with zipfile.ZipFile(zip_file, "r") as zfile:
    #         for file in df["filename"]:
    #             zfile.extract(file, default_dir)

    df = df[~df["filename"].str.contains("快递")]
    # df=df[~df["filename"].str.contains("账单|小红书")]
    df = df[~df["filename"].str.contains("订单")]
    df = df[~df["filename"].str.contains("无")]
    df = df[~df["filename"].str.contains("推广费")]
    df = df[~df["filename"].str.contains("商品")]
    df = df[~df["filename"].str.contains("数据转换")]
    df = df[~df["filename"].str.contains("汇总")]
    df = df[~df["filename"].str.contains("zip")]
    # df=df[~df["filename"].str.contains("支付宝")]

    return df


def unzip_file():
    print("输入要解压的文件夹路径")
    path = input()
    filenames = os.listdir(path)  # 获取目录下所有文件名
    for filename in filenames:
        filepath = os.path.join(path, filename)
        zip_file = zipfile.ZipFile(filepath)  # 获取压缩文件
        # print(filename)
        newfilepath = filename.split(".", 1)[0]  # 获取压缩文件的文件名
        newfilepath = os.path.join(path, newfilepath)
        # print(newfilepath)
        if os.path.isdir(newfilepath):  # 根据获取的压缩文件的文件名建立相应的文件夹
            pass
        else:
            os.mkdir(newfilepath)
        for name in zip_file.namelist():  # 解压文件
            zip_file.extract(name, newfilepath)
        zip_file.close()
        Conf = os.path.join(newfilepath, 'conf')
        if os.path.exists(Conf):  # 如存在配置文件，则删除（需要删则删，不要的话不删）
            shutil.rmtree(Conf)
        if os.path.exists(filepath):  # 删除原先压缩包
            os.remove(filepath)
        print("解压{0}成功".format(filename))


def read_all_excel(rootdir, filekey):
    df_files = get_all_files(rootdir, filekey)
    for index, file in df_files.iterrows():
        if 'df' in locals().keys():  # 如果变量已经存在
            dd = cover(file["filename"])
            dd["filename"] = file["filename"]
            print("进度表：{}/{}  文件{}，行数{}".format(index + 1, df_files.shape[0], file["filename"], dd.shape[0]))
            # print(dd.head(1).to_markdown())
            df = df.append(dd)

        else:
            df = cover(file["filename"])
            df["filename"] = file["filename"]
            print("进度表：{}/{}  文件{}，行数{}".format(index + 1, df_files.shape[0], file["filename"], df.shape[0]))

    return df


def combine_excel():
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

    table = read_all_excel(filedir, filekey)
    del table["filename"]
    table["订单编号"].replace(" ", np.nan, inplace=True)
    table["订单编号"].replace("", np.nan, inplace=True)
    table.dropna(axis=0, subset=["订单编号"], inplace=True)
    table = table.drop_duplicates()

    # table.to_excel(default_dir + "/数据转换/合并表格.xlsx",index=False)
    table.to_csv("data/合并表格.csv", index=False)
    # 转换后的文件合并导出
    print(len(table))
    # print(table.head(5).to_markdown())
    # if len(table) > 500000:
    #     table.to_csv(default_dir + "/数据转换/全平台_数据转换结果.csv",index=False)
    # else:
    #     table.to_excel(default_dir + "/数据转换/全平台_数据转换后订单.xlsx",index=False)
    #
    # # 转换后的文件按平台拆分导出
    # plat_list = ["淘宝", "天猫", "京东", "拼多多","抖音","快手","小红书","网易考拉","有赞","阿里巴巴"]
    # for plat in plat_list:
    #     df_plat = table[table["平台"].str.contains(plat)]
    #     if df_plat.shape[0] > 0:
    #         pagecount = math.ceil(df_plat.shape[0] / 300000.00)
    #         pagecount = "{:d}".format(pagecount)
    #         print("总共需要拆分{}个文件".format(pagecount))
    #         writer = pd.ExcelWriter(default_dir + "/数据转换/{}_数据转换结果.xlsx".format(plat))
    #         # writer = pd.ExcelWriter("data/{}_数据转换结果.xlsx".format(plat))
    #         for x in range(int(pagecount)):
    #             from_line = x * 300000
    #             to_line = (x + 1) * 300000
    #             plat_table = df_plat[from_line:to_line]
    #             print(plat_table.head(5).to_markdown())
    #             print("输出文件总行数：{}".format(plat_table.shape[0]))
    #             sheetname = "Sheet{}".format(x + 1)
    #             plat_table.to_excel(writer,sheetname,engine='xlsxwriter',index=False)
    #             format1 = writer.book.add_format({'num_format': '0.00'})
    #             writer.book.sheetnames[sheetname].set_column('O:O', cell_format=format1)
    #         writer.save()
    #     else:
    #         pass

    return table


if __name__ == "__main__":
    print("开始:", time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()))
    # merge_excel()
    # read_all_sheet()
    # cover()
    # unzip_file()
    combine_excel()

    print("结束:", time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()))
