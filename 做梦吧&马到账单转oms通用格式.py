# __coding=utf8__
# /** 作者：zengyanghui **/
import re
import sys
import os
import future.backports.socketserver
import pandas as pd
import numpy as np
import datetime
import time
import os.path
import xlrd
import xlwt
import pprint
import math
import tabulate

# 显示所有列
pd.set_option('display.max_columns', None)
# 显示所有行
pd.set_option('display.max_rows', None)
# 设置value的显示长度为100，默认为50
pd.set_option('max_colwidth', 200)


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
    # 做梦吧账单逻辑
    if filename.find("做梦吧") >= 0:
        # if ((filename.find("支付宝")>=0)&(filename.find("博滴官方旗舰店")>=0)&(filename.find("汇总")<0)):
        #     df = pd.read_excel(filename, skiprows=4, dtype=str)
        #     for column_name in df.columns:
        #         df.rename(columns={column_name: column_name.replace("\n", "").replace("\t", "").replace(" ", "")},
        #                   inplace=True)
        #     df = df[~df["账务流水号"].str.contains("#")]
        #     print(df.head(5).to_markdown())
        #     if "业务基础订单号" in df.columns:
        #         df["业务基础订单号"] = df["业务基础订单号"].str.replace("\s+", "")
        #     df["账务流水号"] = df["账务流水号"].str.replace("\s+", "")
        #     df["业务流水号"] = df["业务流水号"].str.replace("\s+", "")
        #     df["商户订单号"] = df["商户订单号"].str.replace("\s+", "")
        #     df = df.replace("", np.nan)
        #     print(df.head(5).to_markdown())
        #     plat = "ZMB"
        #     # 订单实付
        #     df1 = pd.DataFrame()
        #     if "业务基础订单号" in df.columns:
        #         df["TID"] = df.apply(lambda x:x["业务基础订单号"] if pd.notnull(x["业务基础订单号"]) else x["商户订单号"],axis=1)
        #         df1["TID"] = df.apply(lambda x:x["TID"] if pd.notnull(x["TID"]) else x["业务流水号"],axis=1)
        #     else:
        #         df1["TID"] = df.apply(lambda x:x["商户订单号"] if pd.notnull(x["商户订单号"]) else x["业务流水号"],axis=1)
        #     df1["SHOPNAME"] = "博滴官方旗舰店"
        #     df1["PLATFORM"] = plat
        #     df1["SHOPCODE"] = "bdghqjd"
        #     df1["BILLPLATFORM"] = plat
        #     df1["CREATED"] = df["发生时间"]
        #     df1["TITLE"] = ""
        #     df1["TRADE_TYPE"] = df["业务类型"]
        #     df1["BUSINESS_NO"] = df["账务流水号"]
        #     df1["INCOME_AMOUNT"] = df["收入金额（+元）"]
        #     df1["EXPEND_AMOUNT"] = df["支出金额（-元）"]
        #     df1["TRADING_CHANNELS"] = df["交易渠道"]
        #     df1["BUSINESS_DESCRIPTION"] = df["业务类型"]
        #     df1["remark"] = df["备注"]
        #     if "业务描述" in df.columns:
        #         df1["IS_REFUNDAMOUNT"] = df.apply(
        #             lambda x: taobao_is_refund(x["业务描述"], x["商品名称"], x["业务类型"], x["备注"], x["支出金额（-元）"], x["商户订单号"],
        #                                        filename), axis=1)
        #     else:
        #         df1["IS_REFUNDAMOUNT"] = df.apply(
        #             lambda x: taobao_is_refund("nan", x["商品名称"], x["业务类型"], x["备注"], x["支出金额（-元）"], x["商户订单号"],
        #                                        filename), axis=1)
        #     if "业务描述" in df.columns:
        #         df1["IS_AMOUNT"] = df.apply(
        #             lambda x: taobao_is_amount(x["业务描述"], x["商品名称"], x["业务类型"], x["商户订单号"], x["收入金额（+元）"], filename),
        #             axis=1)
        #     else:
        #         df1["IS_AMOUNT"] = df.apply(
        #             lambda x: taobao_is_amount("nan", x["商品名称"], x["业务类型"], x["商户订单号"], x["收入金额（+元）"], filename),
        #             axis=1)
        #     df1["OID"] = ""
        #     df1["SOURCEDATA"] = "EXCEL"
        #     df1["RECIPROCAL_ACCOUNT"] = ""
        #     df1["BATCHNO"] = ""
        #     df1["currency"] = ""
        #     df1["overseas_income"] = ""
        #     df1["overseas_expend"] = ""
        #     df1["currency_cny_rate"] = ""
        #     df1["INCOME_AMOUNT"] = df1["INCOME_AMOUNT"].astype(float).abs()
        #     df1["EXPEND_AMOUNT"] = -df1["EXPEND_AMOUNT"].astype(float).abs()
        #     print(df1.head(5).to_markdown())

        if ((filename.find("交易账单") >= 0) & (filename.find("博滴官方旗舰店") >= 0)):
            df = pd.read_excel(filename, dtype=str)
            df["交易状态"] = df["交易状态"].astype(str)
            df = df[~df["交易状态"].str.contains("nan")]
            df = df.replace("[`]","",regex=True)
            df["商品名称"] = df["商品名称"].astype(str)
            df["手续费"] = df["手续费"].astype(float)
            # df = df[df["商品名称"].str.contains("做梦吧 Enjoy Dream")]
            df = df[~df["商品名称"].str.contains("来客电商|订单支付|test|测试小程序下单|自研商城测试订单|测试商品")]
            # if df.shape[0]>0:
            print(df.shape[0])
            print(df.head(5).to_markdown())
            plat = "ZMB"
            # 应结订单金额
            df1 = pd.DataFrame()
            df1["TID"] = df["商户订单号"]
            df1["SHOPNAME"] = ""
            df1["PLATFORM"] = plat
            df1["SHOPCODE"] = ""
            df1["BILLPLATFORM"] = plat
            df1["CREATED"] = df["交易时间"]
            df1["TITLE"] = ""
            df1["TRADE_TYPE"] = "应结订单金额"
            df1["BUSINESS_NO"] = df["商户订单号"]+df["交易状态"]+df["付款银行"]
            df1["INCOME_AMOUNT"] = df["应结订单金额"]
            df1["EXPEND_AMOUNT"] = 0
            df1["TRADING_CHANNELS"] = "微信"
            df1["BUSINESS_DESCRIPTION"] = "应结订单金额"
            df1["remark"] = ""
            df1["IS_REFUNDAMOUNT"] = 0
            df1["IS_AMOUNT"] = df["交易状态"].apply(lambda x:1 if x == "SUCCESS" else 0)
            df1["OID"] = ""
            df1["SOURCEDATA"] = "EXCEL"
            df1["RECIPROCAL_ACCOUNT"] = ""
            df1["BATCHNO"] = ""
            df1["currency"] = ""
            df1["overseas_income"] = ""
            df1["overseas_expend"] = ""
            df1["currency_cny_rate"] = ""
            df1["商户订单号"] = df["商户订单号"]

            # 退款金额
            df2 = df1.copy()
            df2["TRADE_TYPE"] = "退款金额"
            df2["INCOME_AMOUNT"] = 0
            df2["EXPEND_AMOUNT"] = df["退款金额"]
            df2["BUSINESS_DESCRIPTION"] = "退款金额"
            df2["IS_REFUNDAMOUNT"] = df["交易状态"].apply(lambda x: 1 if x == "REFUND" else 0)
            df2["IS_AMOUNT"] = 0

            # 手续费
            df3 = df1.copy()
            df3["TRADE_TYPE"] = "手续费"
            df3["INCOME_AMOUNT"] = df["手续费"].apply(lambda x:x if x<0 else 0)
            df3["EXPEND_AMOUNT"] = df["手续费"].apply(lambda x:x if x>0 else 0)
            df3["BUSINESS_DESCRIPTION"] = "手续费"
            df3["IS_REFUNDAMOUNT"] = 0
            df3["IS_AMOUNT"] = 0

            dfs = [df1,df2,df3]
            df1 = pd.concat(dfs)
            df1 = df1.loc[~((df1.INCOME_AMOUNT == 0) & (df1.EXPEND_AMOUNT == 0))]
            df1["INCOME_AMOUNT"] = df1["INCOME_AMOUNT"].astype(float).abs()
            df1["EXPEND_AMOUNT"] = -df1["EXPEND_AMOUNT"].astype(float).abs()
            if filename.find("博滴官方旗舰店") >= 0:
                df1["SHOPCODE"] = "bbgfqjd"
                df1["SHOPNAME"] = "博滴官方旗舰店"
            elif filename.find("bodyaid") >= 0:
                df1["SHOPCODE"] = "bodyaid"
                df1["SHOPNAME"] = "博滴bodyaid旗舰店"
            elif filename.find("麦凯莱好物精选") >= 0:
                df1["SHOPCODE"] = "mklhwjx"
                df1["SHOPNAME"] = "麦凯莱好物精选"
            elif filename.find("多多提旗舰店") >= 0:
                df1["SHOPCODE"] = "ddt"
                df1["SHOPNAME"] = "多多提旗舰店"
            elif filename.find("Enjoy Dream") >= 0:
                df1["SHOPCODE"] = "ed"
                df1["SHOPNAME"] = "做梦吧 Enjoy Dream"
            print("做梦吧-交易账单")
            print(df1.head(5).to_markdown())
            # else:
            #     dict = {"TID": "nan", "SHOPNAME": "", "PLATFORM": "", "SHOPCODE": "", "BILLPLATFORM": "",
            #             "CREATED": "",
            #             "TITLE": "", "TRADE_TYPE": "", "BUSINESS_NO": "", "INCOME_AMOUNT": "", "EXPEND_AMOUNT": "",
            #             "TRADING_CHANNELS": "", "BUSINESS_DESCRIPTION": "", "remark": "",
            #             "BUSINESS_BILL_SOURCE": "",
            #             "IS_REFUNDAMOUNT": "", "IS_AMOUNT": "", "OID": "", "SOURCEDATA": "",
            #             "RECIPROCAL_ACCOUNT": "",
            #             "BATCHNO": "", "currency": "", "overseas_income": "", "overseas_expend": "",
            #             "currency_cny_rate": ""}
            #     df = pd.DataFrame(dict, index=[0])
            #     return df

        elif ((filename.find("基本账户") >= 0) & (filename.find("博滴官方旗舰店") < 0)):
            df = pd.read_excel(filename, dtype=str)
            df["收支类型"] = df["收支类型"].astype(str)
            df = df[~df["收支类型"].str.contains("nan")]
            df = df.replace("[`]", "", regex=True)
            print(df.head(5).to_markdown())
            plat = "ZMB"
            # 订单实付
            df1 = pd.DataFrame()
            df1["TID"] = df["业务凭证号"]
            df1["SHOPNAME"] = ""
            df1["PLATFORM"] = plat
            df1["SHOPCODE"] = ""
            df1["BILLPLATFORM"] = plat
            df1["CREATED"] = df["记账时间"]
            df1["TITLE"] = ""
            df1["TRADE_TYPE"] = df["业务类型"]
            df1["BUSINESS_NO"] = df["微信支付业务单号"]
            df1["INCOME_AMOUNT"] = df.apply(lambda x: x["收支金额(元)"] if x["收支类型"] == "收入" else 0, axis=1)
            df1["EXPEND_AMOUNT"] = df.apply(lambda x: get_expend(x["备注"], x["收支金额(元)"]) if x["收支类型"] == "支出" else 0, axis=1)
            df1["TRADING_CHANNELS"] = "微信"
            df1["BUSINESS_DESCRIPTION"] = df["业务类型"]
            df1["remark"] = df["备注"]
            df1["IS_REFUNDAMOUNT"] = df["业务类型"].apply(lambda x:1 if x=="退款" else 0)
            df1["IS_AMOUNT"] = df["业务类型"].apply(lambda x:1 if x=="交易" else 0)
            df1["OID"] = ""
            df1["SOURCEDATA"] = "EXCEL"
            df1["RECIPROCAL_ACCOUNT"] = ""
            df1["BATCHNO"] = ""
            df1["资金流水单号"] = df["资金流水单号"]
            df1 = df1.loc[~((df1.INCOME_AMOUNT == 0) & (df1.EXPEND_AMOUNT == 0))]
            df1["INCOME_AMOUNT"] = df1["INCOME_AMOUNT"].astype(float).abs()
            df1["EXPEND_AMOUNT"] = -df1["EXPEND_AMOUNT"].astype(float).abs()
            if filename.find("博滴官方旗舰店") >= 0:
                df1["SHOPCODE"] = "bbgfqjd"
                df1["SHOPNAME"] = "博滴官方旗舰店"
            elif filename.find("bodyaid") >= 0:
                df1["SHOPCODE"] = "bodyaid"
                df1["SHOPNAME"] = "博滴bodyaid旗舰店"
            elif filename.find("麦凯莱好物精选") >= 0:
                df1["SHOPCODE"] = "mklhwjx"
                df1["SHOPNAME"] = "麦凯莱好物精选"
            elif filename.find("多多提旗舰店") >= 0:
                df1["SHOPCODE"] = "ddt"
                df1["SHOPNAME"] = "多多提旗舰店"
            elif filename.find("Enjoy Dream") >= 0:
                df1["SHOPCODE"] = "ed"
                df1["SHOPNAME"] = "做梦吧 Enjoy Dream"
            print("做梦吧-基本账户")
            print(df1.head(5).to_markdown())

        else:
            dict = {"TID": "", "SHOPNAME": "", "PLATFORM": "", "SHOPCODE": "", "BILLPLATFORM": "",
                    "CREATED": "",
                    "TITLE": "", "TRADE_TYPE": "", "BUSINESS_NO": "", "INCOME_AMOUNT": "", "EXPEND_AMOUNT": "",
                    "TRADING_CHANNELS": "", "BUSINESS_DESCRIPTION": "", "remark": "",
                    "BUSINESS_BILL_SOURCE": "",
                    "IS_REFUNDAMOUNT": "", "IS_AMOUNT": "", "OID": "", "SOURCEDATA": "",
                    "RECIPROCAL_ACCOUNT": "",
                    "BATCHNO": "", "currency": "", "overseas_income": "", "overseas_expend": "",
                    "currency_cny_rate": ""}
            df = pd.DataFrame(dict, index=[0])
            return df

    elif filename.find("马到")>=0:
        if filename.find("基本账户") >= 0:
            if filename.find("xls")>=0:
                df = pd.read_excel(filename, dtype=str)
            else:
                try:
                    df = pd.read_csv(filename, dtype=str)
                except Exception as e:
                    df = pd.read_csv(filename, dtype=str, encoding="gb18030")
            df["收支类型"] = df["收支类型"].astype(str)
            df = df[~df["收支类型"].str.contains("nan")]
            df = df.replace("[`]", "", regex=True)
            print(df.head(5).to_markdown())
            plat = "MD"
            df1 = pd.DataFrame()
            df1["TID"] = df["业务凭证号"]
            df1["SHOPNAME"] = "麦凯莱臻选"
            df1["PLATFORM"] = plat
            df1["SHOPCODE"] = "mklzx"
            df1["BILLPLATFORM"] = plat
            df1["CREATED"] = df["记账时间"]
            df1["TITLE"] = ""
            df1["TRADE_TYPE"] = df["业务类型"]
            df1["BUSINESS_NO"] = df["微信支付业务单号"]
            df1["INCOME_AMOUNT"] = df.apply(lambda x:x["收支金额(元)"] if x["收支类型"] == "收入" else 0, axis=1)
            df1["EXPEND_AMOUNT"] = df.apply(lambda x: get_expend(x["备注"], x["收支金额(元)"]) if x["收支类型"] == "支出" else 0, axis=1)
            df1["TRADING_CHANNELS"] = "微信"
            df1["BUSINESS_DESCRIPTION"] = df["业务类型"]
            df1["remark"] = df["备注"]
            df1["IS_REFUNDAMOUNT"] = df["业务类型"].apply(lambda x:1 if x=="退款" else 0)
            df1["IS_AMOUNT"] = df["业务类型"].apply(lambda x:1 if x=="交易" else 0)
            df1["OID"] = ""
            df1["SOURCEDATA"] = "EXCEL"
            df1["RECIPROCAL_ACCOUNT"] = ""
            df1["BATCHNO"] = ""
            df1["资金流水单号"] = df["资金流水单号"]
            df1 = df1.loc[~((df1.INCOME_AMOUNT == 0) & (df1.EXPEND_AMOUNT == 0))]
            df1["INCOME_AMOUNT"] = df1["INCOME_AMOUNT"].astype(float).abs()
            df1["EXPEND_AMOUNT"] = -df1["EXPEND_AMOUNT"].astype(float).abs()
            print("做梦吧-基本账户")
            print(df1.head(5).to_markdown())

        else:
            dict = {"TID": "", "SHOPNAME": "", "PLATFORM": "", "SHOPCODE": "", "BILLPLATFORM": "",
                    "CREATED": "",
                    "TITLE": "", "TRADE_TYPE": "", "BUSINESS_NO": "", "INCOME_AMOUNT": "", "EXPEND_AMOUNT": "",
                    "TRADING_CHANNELS": "", "BUSINESS_DESCRIPTION": "", "remark": "",
                    "BUSINESS_BILL_SOURCE": "",
                    "IS_REFUNDAMOUNT": "", "IS_AMOUNT": "", "OID": "", "SOURCEDATA": "",
                    "RECIPROCAL_ACCOUNT": "",
                    "BATCHNO": "", "currency": "", "overseas_income": "", "overseas_expend": "",
                    "currency_cny_rate": ""}
            df = pd.DataFrame(dict, index=[0])
            return df

    return df1


def get_expend(remark,amount):
    if remark.find("退款总金额") >= 0:
        expend = "".join(("".join(remark.split("退款总金额")[1:])).split("元")[:1])
        return expend
    else:
        return amount



def read_bill2(filename):
    print(filename)
    # 做梦吧订单逻辑
    if filename.find("做梦吧") >= 0:
        if filename.find("订单") >= 0:
            df = pd.read_excel(filename, dtype=str)
            df = df[["订单号","支付单号"]]
            df["支付单号"] = df["支付单号"].astype(str)
            df = df[~df["支付单号"].str.contains("nan")]
            # df = df.replace("[`]", "", regex=True)
            print(df.head(5).to_markdown())
            print("做梦吧-订单")

        else:
            print("非订单文件！")
            dict = {"订单号": "", "支付单号": ""}
            df = pd.DataFrame(dict, index=[0])
            return df

    elif filename.find("马到") >= 0:
        if filename.find("订单") >= 0:
            df = pd.read_excel(filename, dtype=str)
            if "订单编号" in df.columns:
                df.rename(columns={"订单编号":"订单号"},inplace=True)
            df = df[["订单号","支付单号"]]
            df["支付单号"] = df["支付单号"].astype(str)
            df = df[~df["支付单号"].str.contains("nan")]
            # df = df.replace("[`]", "", regex=True)
            print(df.head(5).to_markdown())
            print("做梦吧-订单")

        else:
            print("非订单文件！")
            dict = {"订单号": "", "支付单号": ""}
            df = pd.DataFrame(dict, index=[0])
            return df

    return df


def taobao_is_refund(desc,title,type,rmark,amount,btid,filename):
    if desc.find("nan")<0:
        if ((desc.find("0020001|交易退款-余额退款")>=0) & (amount < 0)):
            return 1
        elif ((desc.find("0020002|交易退款-保证金退款")>=0) & (amount < 0)):
            return 1
        elif ((desc.find("0020005|交易退款-售中退款（极速回款）")>=0) & (amount < 0)):
            return 1
        elif ((desc.find("0020011|交易退款-交易退款")>=0) & (amount < 0)):
            return 1
        elif ((desc.find("064000200001|交易还款-提前收款-花呗交易")>=0) & (amount < 0)):
            return 1
        elif ((desc.find("064000200002|交易还款-提前收款-售中退款")>=0) & (amount < 0)):
            return 1
        else:
            return 0
    else:
        if ((type == "交易退款") & (amount < 0)):
            return 1
        elif ((rmark.find("售后支付") >= 0) & (amount < 0)):
            return 1
        elif ((rmark.find("保证金退款") >= 0) & (amount < 0)):
            return 1
        else:
            return 0


def taobao_is_amount(desc,title,type,btid,amount,filename):
    if desc.find("nan")<0:
        if ((desc.find("0010001|交易收款-交易收款")>=0) & (amount > 0)):
            return 1
        elif ((desc.find("0010002|交易收款-预售定金（买家责任不退还）")>=0) & (amount > 0)):
            return 1
        elif ((desc.find("0010022|交易收款-提前收款")>=0) & (amount > 0)):
            return 1
        elif ((desc.find("001002200001|交易收款-提前收款-花呗交易")>=0) & (amount > 0)):
            return 1
        else:
            return 0
    else:
        if ((type == "交易付款") & (amount > 0)):
            return 1
        elif ((type == "在线支付") & (amount > 0)):
            return 1
        else:
            return 0


def get_all_files(rootdir, filekey):
    filelist = list_all_files(rootdir, filekey)
    # print(filelist)
    mySeries = pd.Series(filelist)
    df = pd.DataFrame(mySeries)
    df.columns = ["filename"]
    # df["filename"] = df["filename"].apply(lambda x: x.strip().replace(" ", ""))

    df = df[~df["filename"].str.contains("汇总")]

    return df


def read_all_bill(rootdir, filekey):
    print("账单文件处理中......")
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


def read_all_bill2(rootdir, filekey):
    print("订单文件处理中......")
    df_files = get_all_files(rootdir, filekey)
    for index, file in df_files.iterrows():
        if 'df' in locals().keys():  # 如果变量已经存在
            dd = read_bill2(file["filename"])
            dd["filename"] = file["filename"]
            print("进度表：{}/{}  文件{}，行数{}".format(index + 1, df_files.shape[0], file["filename"], dd.shape[0]))
            df = df.append(dd)


        else:
            df = read_bill2(file["filename"])
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

    print("你希望在'{}'目录下找到所有的包含“{}”文件，然后合并。".format(filedir, filekey))

    # table = read_all_excel(filedir, filekey)
    table1 = read_all_bill(filedir, filekey)
    table2 = read_all_bill2(filedir, filekey)
    # del table1["filename"]
    del table2["filename"]
    table2["支付单号"] = table2["支付单号"].astype(str)
    table2 = table2[~table2["支付单号"].str.contains("nan")]
    print(table2.head().to_markdown())

    if (("商户订单号" in table1.columns) & ("资金流水单号" not in table1.columns)):
        table = pd.merge(table1, table2, how="left", left_on="商户订单号", right_on="支付单号")
        # table = pd.merge(table,table2,how="left",left_on="资金流水单号",right_on="支付单号")
        print(table.head().to_markdown())
        table["TID"] = table.apply(lambda x: x["订单号"] if pd.isnull(x["TID"]) else x["TID"], axis=1)
        table["TID"] = table.apply(
            lambda x: x["商户订单号"] if ((pd.isnull(x["TID"])) & (x["SHOPNAME"] == "博滴官方旗舰店")) else x["TID"], axis=1)
        # table["TID"] = table.apply(lambda x:x["订单号_y"] if pd.isnull(x["TID"]) else x["TID"],axis=1)
        # table["TID"] = table.apply(lambda x:x["订单号_y"] if ((x["TRADE_TYPE"]=="退款")&(pd.notnull(x["订单号_y"]))) else x["TID"],axis=1)
        del table["商户订单号"]
        # del table["资金流水单号"]
        del table["支付单号"]
        # del table["支付单号_y"]
        del table["订单号"]
        # del table["订单号_y"]

    elif "商户订单号" in table1.columns:
        table = pd.merge(table1,table2,how="left",left_on="商户订单号",right_on="支付单号")
        table = pd.merge(table,table2,how="left",left_on="资金流水单号",right_on="支付单号")
        print(table.head().to_markdown())
        table["TID"] = table.apply(lambda x:x["订单号_x"] if pd.isnull(x["TID"]) else x["TID"],axis=1)
        table["TID"] = table.apply(lambda x:x["商户订单号"] if ((pd.isnull(x["TID"]))&(x["SHOPNAME"]=="博滴官方旗舰店")) else x["TID"],axis=1)
        table["TID"] = table.apply(lambda x:x["订单号_y"] if pd.isnull(x["TID"]) else x["TID"],axis=1)
        table["TID"] = table.apply(lambda x:x["订单号_y"] if ((x["TRADE_TYPE"]=="退款")&(pd.notnull(x["订单号_y"]))) else x["TID"],axis=1)
        del table["商户订单号"]
        del table["资金流水单号"]
        del table["支付单号_x"]
        del table["支付单号_y"]
        del table["订单号_x"]
        del table["订单号_y"]

    else:
        table = pd.merge(table1, table2, how="left", left_on="资金流水单号", right_on="支付单号")
        print(table.head().to_markdown())
        table["TID"] = table.apply(lambda x: x["订单号"] if pd.isnull(x["TID"]) else x["TID"], axis=1)
        table["TID"] = table.apply(lambda x: x["订单号"] if ((x["TRADE_TYPE"]=="退款")&(pd.notnull(x["订单号"]))) else x["TID"], axis=1)
        del table["资金流水单号"]
        del table["支付单号"]
        del table["订单号"]
    table = table.sort_values(by=["TID", "CREATED"])
    print(table.head().to_markdown())

    # del table["filename"]

    # if table.shape[0] < 800000:
    #     table.to_excel(default_dir + "/处理后的账单.xlsx", index=False)
    # else:
    #     table.to_csv(default_dir + "/处理后的账单.csv", index=False)
    index = 0
    # if "TID" in table.columns:

    table["TID"] = table["TID"].astype(str)
    # table["TID"] = table["TID"].apply(lambda x:x.replace(" ",np.nan).replace("\n",np.nan))
    # table["TID"].fillna("nan",inplace=True)
    table["TID"] = table["TID"].apply(lambda x:"nan" if len(x)<1 else x)
    table.replace("nan", np.nan, inplace=True)
    table.dropna(subset=["TID"], inplace=True)
    # table.drop_duplicates(inplace=True)
    table = table.loc[~((table.INCOME_AMOUNT == 0) & (table.EXPEND_AMOUNT == 0))]
    plat = os.sep.join(default_dir.split(os.sep)[-1:])
    print("第{}个表格,记录数:{}".format(index, table.shape[0]))
    print(table.head(10).to_markdown())
    # df.to_excel(r"work/合并表格_test.xlsx")
    print("账单总行数：")
    print(table.shape[0])
    for i in range(0, int(table.shape[0] / 200000) + 1):
        print("存储分页：{}  from:{} to:{}".format(i, i * 200000, (i + 1) * 200000))
        # df.iloc[i*500000:(i+1)*500000].to_csv(default_dir + "\{}_合并表格_{}.csv".format(index,i))
        table.iloc[i * 200000:(i + 1) * 200000].to_excel(default_dir + "\{}-处理后的账单{}.xlsx".format(plat, i), index=False)

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