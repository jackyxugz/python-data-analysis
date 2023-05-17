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
    # 马到逻辑
    if filename.find("马到") >= 0:
        df = pd.read_excel(filename, dtype=str)
        print(df.head(5).to_markdown())
        # df = df[~df["状态"].str.contains("进行中")]
        plat = "MD"
        df1 = pd.DataFrame()
        df1["TID"] = df["订单号"]
        df1["SHOPNAME"] = "麦凯莱臻选"
        df1["PLATFORM"] = plat
        df1["SHOPCODE"] = "mklzx"
        df1["BILLPLATFORM"] = plat
        df1["CREATED"] = df["支付时间"]
        df1["TITLE"] = df["商品名称"]
        df1["TRADE_TYPE"] = df.apply(
            lambda x: "收入" if ((x["状态"].find("已结算") >= 0) or (x["状态"].find("进行中") >= 0)) else "支出", axis=1)
        df1["BUSINESS_NO"] = df["支付单号"]
        df1["INCOME_AMOUNT"] = df.apply(
            lambda x: x["金额（元）"] if ((x["状态"].find("已结算") >= 0) or (x["状态"].find("进行中") >= 0)) else 0, axis=1)
        df1["EXPEND_AMOUNT"] = df.apply(lambda x: x["金额（元）"] if x["状态"] == "退款" else 0, axis=1)
        df1["TRADING_CHANNELS"] = df["支付方式"]
        df1["BUSINESS_DESCRIPTION"] = df["状态"]
        df1["remark"] = ""
        df1["BUSINESS_BILL_SOURCE"] = ""
        df1["IS_REFUNDAMOUNT"] = df.apply(lambda x: 1 if x["状态"] == "退款" else 0, axis=1)
        df1["IS_AMOUNT"] = df.apply(lambda x: 1 if x["状态"] == "已结算" else 0, axis=1)
        df1["OID"] = ""
        df1["SOURCEDATA"] = "EXCEL"
        df1["RECIPROCAL_ACCOUNT"] = ""
        df1["BATCHNO"] = ""
        df1["INCOME_AMOUNT"] = df1["INCOME_AMOUNT"].astype(float).abs()
        df1["EXPEND_AMOUNT"] = -df1["EXPEND_AMOUNT"].astype(float).abs()

        df2 = df1.copy()
        df2["TRADE_TYPE"] = df.apply(
            lambda x: "支出" if ((x["状态"].find("已结算") >= 0) or (x["状态"].find("进行中") >= 0)) else "收入", axis=1)
        df2["INCOME_AMOUNT"] = df.apply(lambda x: x["手续费"] if x["状态"] == "退款" else 0, axis=1)
        df2["EXPEND_AMOUNT"] = df.apply(
            lambda x: x["手续费"] if ((x["状态"].find("已结算") >= 0) or (x["状态"].find("进行中") >= 0)) else 0, axis=1)
        df2["BUSINESS_DESCRIPTION"] = "手续费"
        df2["IS_REFUNDAMOUNT"] = 0
        df2["IS_AMOUNT"] = 0
        df2["INCOME_AMOUNT"] = df2["INCOME_AMOUNT"].astype(float).abs()
        df2["EXPEND_AMOUNT"] = -df2["EXPEND_AMOUNT"].astype(float).abs()

        dfs = [df1, df2]
        df1 = pd.concat(dfs)
        df1 = df1.sort_values(by=["TID", "CREATED"])
        df1.drop_duplicates(inplace=True)
        print(df1.head(5).to_markdown())
        print("马到")

    # 枫叶小店逻辑
    elif filename.find("枫叶") >= 0:
        df = pd.read_excel(filename, dtype=str)
        print(df.head(5).to_markdown())
        if "关联单号(订单号/退款单号)" in df.columns:
            df.rename(columns={"关联单号(订单号/退款单号)": "订单号"}, inplace=True)
        if "订单编号" in df.columns:
            df.rename(columns={"订单编号": "订单号"}, inplace=True)
        else:
            pass
        plat = "FY"
        df1 = pd.DataFrame()
        df1["TID"] = df["订单号"]
        df1["SHOPNAME"] = ""
        df1["PLATFORM"] = plat
        df1["SHOPCODE"] = ""
        df1["BILLPLATFORM"] = plat
        df1["CREATED"] = df["交易时间"]
        df1["TITLE"] = ""
        df1["TRADE_TYPE"] = df["业务类型"]
        df1["BUSINESS_NO"] = df["交易单号"]
        df1["INCOME_AMOUNT"] = df.apply(lambda x: x["收支金额(元)"] if x["收支类型"] == "收入" else 0, axis=1)
        df1["EXPEND_AMOUNT"] = df.apply(lambda x: x["收支金额(元)"] if x["收支类型"] == "支出" else 0, axis=1)
        df1["TRADING_CHANNELS"] = ""
        df1["BUSINESS_DESCRIPTION"] = df["收支类型"]
        df1["IS_REFUNDAMOUNT"] = df.apply(lambda x: 1 if ((x["收支类型"] == "支出") & (x["业务类型"] == "订单退款")) else 0, axis=1)
        df1["IS_AMOUNT"] = df.apply(lambda x: 1 if ((x["收支类型"] == "收入") & (x["业务类型"] == "订单交易")) else 0, axis=1)
        df1["OID"] = ""
        df1["SOURCEDATA"] = "EXCEL"
        df1["RECIPROCAL_ACCOUNT"] = ""
        df1["BATCHNO"] = ""
        df1["currency"] = ""
        df1["overseas_income"] = ""
        df1["overseas_expend"] = ""
        df1["currency_cny_rate"] = ""
        # df.loc[(df.业务类型 == "订单退款" & df.收支类型 == "支出"), "是否是退款"] = "1"
        # df.loc[df.业务类型 == "订单退款", "是否是回款"] = "0"
        # df.loc[df.业务类型 == "订单交易", "是否是退款"] = "0"
        # df.loc[df.业务类型 == "订单交易", "是否是回款"] = "1"
        # df1["是否是退款"] = df["是否是退款"]
        # df1["是否是回款"] = df["是否是回款"]

        if filename.find("好物精选") >= 0:
            df1["SHOPNAME"] = "麦凯莱好物精选"
            df1["SHOPCODE"] = "mklhwjx"
        elif filename.find("精选好物") >= 0:
            df1["SHOPNAME"] = "麦凯莱精选好物"
            df1["SHOPCODE"] = "mkljxhw"
        elif filename.find("麦凯莱严选") >= 0:
            df1["SHOPNAME"] = "麦凯莱严选"
            df1["SHOPCODE"] = "mklyx"
        elif filename.find("Mega洗护优选") >= 0:
            df1["SHOPNAME"] = "Mega洗护优选"
            df1["SHOPCODE"] = "mxhyx"
        elif filename.find("Mega严选") >= 0:
            df1["SHOPNAME"] = "Mega严选"
            df1["SHOPCODE"] = "megayx"

        df1["INCOME_AMOUNT"] = df1["INCOME_AMOUNT"].astype(float).abs()
        df1["EXPEND_AMOUNT"] = -df1["EXPEND_AMOUNT"].astype(float)
        print(df1.head(5).to_markdown())
        print("枫叶")

        return df1

    # 有赞逻辑
    elif filename.find("有赞") >= 0:
        if filename.find("汇总") >= 0:
            dict = {"TID": "nan", "SHOPNAME": "", "PLATFORM": "", "SHOPCODE": "", "BILLPLATFORM": "", "CREATED": "",
                    "TITLE": "", "TRADE_TYPE": "", "BUSINESS_NO": "", "INCOME_AMOUNT": "", "EXPEND_AMOUNT": "",
                    "TRADING_CHANNELS": "", "BUSINESS_DESCRIPTION": "", "remark": "", "BUSINESS_BILL_SOURCE": "",
                    "IS_REFUNDAMOUNT": "", "IS_AMOUNT": "", "OID": "", "SOURCEDATA": "", "RECIPROCAL_ACCOUNT": "",
                    "BATCHNO": "", "currency": "", "overseas_income": "", "overseas_expend": "",
                    "currency_cny_rate": ""}
            df = pd.DataFrame(dict, index=[0])
            return df
        elif filename.find("储值卡交易记录") >= 0:
            if filename.find("xls") >= 0:
                df = pd.read_excel(filename, dtype=str)
            else:
                df = pd.read_csv(filename, dtype=str)
                print(df.head().to_markdown())

            # df = df[df["类型"].str.contains("消费|退款")]
            # df = df[~df["单号"].str.contains("RN")]
            if filename.find("BodyAid品牌商城") >= 0:
                plat = "YZ"
                # 类型=导入/调账，类型2包含消费/退款
                if "类型2" in df.columns:
                    dfx = df.loc[(df["类型"].str.contains("导入|调账") & (df["类型2"].str.contains("消费|退款")))]
                    dfx["消费对应订单"] = dfx["消费对应订单"].astype(str)
                    df0 = pd.DataFrame()
                    df0["TID"] = dfx["消费对应订单"]
                    df0["SHOPNAME"] = ""
                    df0["PLATFORM"] = plat
                    df0["SHOPCODE"] = ""
                    df0["BILLPLATFORM"] = plat
                    df0["CREATED"] = dfx["时间"]
                    df0["TITLE"] = dfx["备注"]
                    df0["TRADE_TYPE"] = dfx.apply(lambda x: "消费" if x["类型2"].find("消费") >= 0 else "退款", axis=1)
                    df0["BUSINESS_NO"] = dfx["储值卡号"]
                    df0["INCOME_AMOUNT"] = dfx.apply(lambda x: x["消费金额"] if x["类型2"].find("消费") >= 0 else 0, axis=1)
                    df0["EXPEND_AMOUNT"] = dfx.apply(
                        lambda x: x["消费金额"] if ((x["类型2"].find("退款") >= 0) & (x["消费对应订单"].find("RN") < 0)) else 0,
                        axis=1)
                    df0["TRADING_CHANNELS"] = ""
                    df0["BUSINESS_DESCRIPTION"] = "储值卡交易记录：类型=" + dfx["类型"] + "，类型2=" + dfx["类型2"]
                    df0["remark"] = "卡名称：" + dfx["卡名称"] + "。储值卡号：" + dfx["储值卡号"]
                    df0["BUSINESS_BILL_SOURCE"] = ""
                    df0["IS_REFUNDAMOUNT"] = dfx.apply(
                        lambda x: 1 if ((x["类型2"].find("退款") >= 0) & (x["消费对应订单"].find("RN") < 0)) else 0, axis=1)
                    df0["IS_AMOUNT"] = dfx.apply(lambda x: 1 if x["类型2"].find("消费") >= 0 else 0, axis=1)
                    df0["OID"] = ""
                    df0["SOURCEDATA"] = "EXCEL"
                    df0["RECIPROCAL_ACCOUNT"] = ""
                    df0["BATCHNO"] = ""
                    df0["currency"] = ""
                    df0["overseas_income"] = ""
                    df0["overseas_expend"] = ""
                    df0["currency_cny_rate"] = ""
                else:
                    dict = {"TID": "nan", "SHOPNAME": "", "PLATFORM": "", "SHOPCODE": "", "BILLPLATFORM": "",
                            "CREATED": "",
                            "TITLE": "", "TRADE_TYPE": "", "BUSINESS_NO": "", "INCOME_AMOUNT": "", "EXPEND_AMOUNT": "",
                            "TRADING_CHANNELS": "", "BUSINESS_DESCRIPTION": "", "remark": "",
                            "BUSINESS_BILL_SOURCE": "",
                            "IS_REFUNDAMOUNT": "", "IS_AMOUNT": "", "OID": "", "SOURCEDATA": "",
                            "RECIPROCAL_ACCOUNT": "",
                            "BATCHNO": "", "currency": "", "overseas_income": "", "overseas_expend": "",
                            "currency_cny_rate": ""}
                    df0 = pd.DataFrame(dict, index=[0])

                # 类型=消费/退款
                df = df.loc[df["类型"].str.contains("消费|退款")]
                df["余额变动"] = df["余额变动"].apply(lambda x: x.replace("+", "").replace("-", ""))
                df1 = pd.DataFrame()
                df1["TID"] = df["单号"]
                df1["SHOPNAME"] = ""
                df1["PLATFORM"] = plat
                df1["SHOPCODE"] = ""
                df1["BILLPLATFORM"] = plat
                df1["CREATED"] = df["时间"]
                df1["TITLE"] = df["备注"]
                df1["TRADE_TYPE"] = df["类型"]
                df1["BUSINESS_NO"] = df["储值卡号"]
                df1["INCOME_AMOUNT"] = df.apply(lambda x: x["余额变动"] if x["类型"] == "消费" else 0, axis=1)
                df1["EXPEND_AMOUNT"] = df.apply(
                    lambda x: x["余额变动"] if ((x["类型"] == "退款") & (x["单号"].find("RN") < 0)) else 0, axis=1)
                df1["TRADING_CHANNELS"] = ""
                df1["BUSINESS_DESCRIPTION"] = "储值卡交易记录"
                df1["remark"] = "卡名称：" + df["卡名称"] + "。储值卡号：" + df["储值卡号"]
                df1["BUSINESS_BILL_SOURCE"] = ""
                df1["IS_REFUNDAMOUNT"] = df.apply(lambda x: 1 if ((x["类型"] == "退款") & (x["单号"].find("RN") < 0)) else 0,
                                                  axis=1)
                df1["IS_AMOUNT"] = df.apply(lambda x: 1 if x["类型"] == "消费" else 0, axis=1)
                df1["OID"] = ""
                df1["SOURCEDATA"] = "EXCEL"
                df1["RECIPROCAL_ACCOUNT"] = ""
                df1["BATCHNO"] = ""
                df1["currency"] = ""
                df1["overseas_income"] = ""
                df1["overseas_expend"] = ""
                df1["currency_cny_rate"] = ""

                df1 = pd.concat([df0, df1])

            else:
                plat = "YZ"
                df1 = pd.DataFrame()
                df1["TID"] = df["单号"]
                df1["SHOPNAME"] = ""
                df1["PLATFORM"] = plat
                df1["SHOPCODE"] = ""
                df1["BILLPLATFORM"] = plat
                df1["CREATED"] = df["时间"]
                df1["TITLE"] = df["备注"]
                df1["TRADE_TYPE"] = df["类型"]
                df1["BUSINESS_NO"] = df["储值卡号"]
                df1["INCOME_AMOUNT"] = df.apply(lambda x: x["本金变动"] if x["类型"] == "消费" else 0, axis=1)
                df1["EXPEND_AMOUNT"] = df.apply(
                    lambda x: x["本金变动"] if ((x["类型"] == "退款") & (x["单号"].find("RN") < 0)) else 0, axis=1)
                df1["TRADING_CHANNELS"] = ""
                df1["BUSINESS_DESCRIPTION"] = "储值卡交易记录"
                df1["remark"] = "卡名称：" + df["卡名称"] + "。储值卡号：" + df["储值卡号"]
                df1["BUSINESS_BILL_SOURCE"] = ""
                df1["IS_REFUNDAMOUNT"] = df.apply(lambda x: 1 if ((x["类型"] == "退款") & (x["单号"].find("RN") < 0)) else 0,
                                                  axis=1)
                df1["IS_AMOUNT"] = df.apply(lambda x: 1 if x["类型"] == "消费" else 0, axis=1)
                df1["OID"] = ""
                df1["SOURCEDATA"] = "EXCEL"
                df1["RECIPROCAL_ACCOUNT"] = ""
                df1["BATCHNO"] = ""
                df1["currency"] = ""
                df1["overseas_income"] = ""
                df1["overseas_expend"] = ""
                df1["currency_cny_rate"] = ""

            if filename.find("BodyAid品牌商城") >= 0:
                df1["SHOPNAME"] = "bodyaid品牌商城"
                df1["SHOPCODE"] = "bodyaidppsc"
            elif filename.find("卖家联合全球购") >= 0:
                df1["SHOPNAME"] = "卖家联合全球购"
                df1["SHOPCODE"] = "mjlhqqg"
            elif filename.find("爱家白皮书") >= 0:
                df1["SHOPNAME"] = "爱家白皮书"
                df1["SHOPCODE"] = "bpscm"
            elif ((filename.find("麦凯莱国际好物") >= 0) | (filename.find("麦凯莱严选") >= 0)):
                df1["SHOPNAME"] = "麦凯莱国际好物"
                df1["SHOPCODE"] = "mklyx"
            # df1.to_excel("data/test.xlsx")
            print(df1.tail().to_markdown())
            df1 = df1[~df1["TID"].str.contains("nan")]
            df1["INCOME_AMOUNT"] = df1["INCOME_AMOUNT"].astype(float).abs()
            df1["EXPEND_AMOUNT"] = -df1["EXPEND_AMOUNT"].astype(float)
            print(df1.head(5).to_markdown())
            print("有赞")
        elif filename.find("账单") >= 0:
            if filename.find("xls") >= 0:
                df = pd.read_excel(filename, dtype=str)
            else:
                # df = pd.read_csv(filename, dtype=str,index_col=False)
                try:
                    df = pd.read_csv(filename, dtype=str, index_col=False)
                except Exception as e:
                    df = pd.read_csv(filename, dtype=str, index_col=False, encoding="gb18030")
                print(df.head().to_markdown())

            # if filename.find("对账单")>=0:
            #     df = pd.DataFrame(df).reset_index()
            print(len(df))
            df["下单时间"] = df["下单时间"].astype("datetime64[ns]")
            if filename.find("BodyAid品牌商城") >= 0:
                df = df.loc[df["下单时间"] >= "2021-05-12 00:00:00"]
            print(len(df))
            if len(df) == 0:
                dict = {"TID": "nan", "SHOPNAME": "", "PLATFORM": "", "SHOPCODE": "", "BILLPLATFORM": "", "CREATED": "",
                        "TITLE": "", "TRADE_TYPE": "", "BUSINESS_NO": "", "INCOME_AMOUNT": "", "EXPEND_AMOUNT": "",
                        "TRADING_CHANNELS": "", "BUSINESS_DESCRIPTION": "", "remark": "", "BUSINESS_BILL_SOURCE": "",
                        "IS_REFUNDAMOUNT": "", "IS_AMOUNT": "", "OID": "", "SOURCEDATA": "", "RECIPROCAL_ACCOUNT": "",
                        "BATCHNO": "", "currency": "", "overseas_income": "", "overseas_expend": "",
                        "currency_cny_rate": ""}
                df = pd.DataFrame(dict, index=[0])
                return df
            df["业务单号"] = df["业务单号"].astype(str)
            df["类型"] = df["类型"].str.replace("\s+", "")
            df["收入(元)"] = df["收入(元)"].astype(float)
            print(df.head(5).to_markdown())

            if filename.find("白皮书") >= 0:
                if len(df.loc[df["类型"] == "采购"]) > 0:
                    df0 = df[df["类型"].str.contains("采购")]
                    df0 = df0[["关联单号", "支出(元)"]]
                    df0.columns = ["关联单号1", "支出(元)1"]
                    df0["类型"] = "订单入账"
                    df = pd.merge(df, df0, how="left", left_on=["业务单号", "类型"], right_on=["关联单号1", "类型"])
                    df["支出(元)1"].fillna(0, inplace=True)
                    df["支出(元)1"] = df["支出(元)1"].astype(float)
                else:
                    pass

            plat = "YZ"
            df1 = pd.DataFrame()
            df1["TID"] = df["业务单号"]
            df1["SHOPNAME"] = ""
            df1["PLATFORM"] = plat
            df1["SHOPCODE"] = ""
            df1["BILLPLATFORM"] = plat
            df1["CREATED"] = df["入账时间"]
            df1["TITLE"] = df["名称"]
            df1["TRADE_TYPE"] = df["类型"]
            df1["BUSINESS_NO"] = df["关联单号"]
            df1["INCOME_AMOUNT"] = df["收入(元)"]
            df1["EXPEND_AMOUNT"] = df["支出(元)"]
            df1["TRADING_CHANNELS"] = df["渠道"]
            df1["BUSINESS_DESCRIPTION"] = df["类型"]
            df1["remark"] = "附加信息：" + df["附加信息"] + "。备注：" + df["备注"]
            df1["BUSINESS_BILL_SOURCE"] = ""
            df1["IS_REFUNDAMOUNT"] = df.apply(lambda x: 1 if ((x["类型"] == "退款") & (x["业务单号"].find("RN") < 0)) else 0,
                                              axis=1)
            if "账户类型" in df.columns:
                df["账户类型"] = df["账户类型"].str.replace("\s+", "")
                if "支出(元)1" in df.columns:
                    df1["IS_AMOUNT"] = df.apply(
                        lambda x: 1 if ((x["类型"] == "订单入账") & (x["收入(元)"] - x["支出(元)1"] != 0)) else 0, axis=1)
                else:
                    df1["IS_AMOUNT"] = df.apply(
                        lambda x: 1 if x["类型"] == "订单入账" else 0, axis=1)
            else:
                df["账户"] = df["账户"].str.replace("\s+", "")
                print(df.to_markdown())
                df1["IS_AMOUNT"] = df.apply(
                    lambda x: 1 if x["类型"] == "订单入账" else 0, axis=1)
            df1["OID"] = ""
            df1["SOURCEDATA"] = "EXCEL"
            df1["RECIPROCAL_ACCOUNT"] = ""
            df1["BATCHNO"] = ""
            df1["currency"] = ""
            df1["overseas_income"] = ""
            df1["overseas_expend"] = ""
            df1["currency_cny_rate"] = ""

            if filename.find("BodyAid品牌商城") >= 0:
                df1["SHOPNAME"] = "bodyaid品牌商城"
                df1["SHOPCODE"] = "bodyaidppsc"
            elif filename.find("卖家联合全球购") >= 0:
                df1["SHOPNAME"] = "卖家联合全球购"
                df1["SHOPCODE"] = "mjlhqqg"
            elif filename.find("爱家白皮书") >= 0:
                df1["SHOPNAME"] = "爱家白皮书"
                df1["SHOPCODE"] = "bpscm"
            elif ((filename.find("麦凯莱国际好物") >= 0) | (filename.find("麦凯莱严选") >= 0)):
                df1["SHOPNAME"] = "麦凯莱严选/国际好物"
                df1["SHOPCODE"] = "mklyx"
            # df1.to_excel("data/test.xlsx")

            # BodyAid品牌商城,如果有储值卡交易记录，合并后排除相关的账单数据
            if filename.find("BodyAid品牌商城") >= 0:
                print("合并储值卡交易记录")
                filepath = default_dir

                filepath_export = filepath + os.sep + "shunfen" + ".xlsx"
                if os.path.isfile(filepath_export):
                    os.remove(filepath_export)

                filepathlist = os.listdir(filepath)
                print(filepathlist)
                filelist = []
                for i in filepathlist:
                    if ((i.find("储值卡交易记录") >= 0) & (i.find("~") < 0)):
                        # del filepathlist[i]
                        # filepathlist.remove(i)
                        filelist.append(i)
                print(filelist)
                for b in filelist:
                    filename = filepath + os.sep + str(b)
                    print(filename)
                    df = pd.read_excel(filename, dtype=str)
                    print(df.head(1).to_markdown())
                    if "类型2" not in df.columns:
                        continue
                    print(f"df:{len(df)}")
                    df2 = df.loc[df["类型2"].str.contains('充值', na=False), :]
                    print(f"df2:{len(df2)}")
                    df3 = None
                    if df3 is None:
                        df3 = df2
                        print(len(df3))
                    else:
                        df3 = pd.concat([df3, df2])
                        print(len(df3))

                print("账单文件删除储值卡交易记录的充值订单前行数{}".format(len(df1)))
                print("储值卡交易记录的充值订单行数{}".format(len(df3)))

                try:
                    df1["TID"] = df1["TID"].astype(str)
                    df1["TID"] = df1["TID"].apply(lambda x: x.replace(" ", "").replace("\n", "").replace("\t", ""))
                    df3["充值对应订单"] = df3["充值对应订单"].astype(str)
                    df3["TID"] = df3["充值对应订单"]
                    df3["TID2"] = df3["TID"].apply(lambda x: x if x.find("/") >= 0 else "nan")
                    print(df3.head().to_markdown())
                    df4 = df3[~df3["TID2"].str.contains("nan")]
                    if len(df4) > 0:
                        print(df4.head().to_markdown())
                        df3_split = df4["TID2"].str.split("/", expand=True)
                        print(df3_split.head().to_markdown())
                        df3_split = df3_split.stack()
                        df3_split = df3_split.reset_index()
                        print(df3_split.head().to_markdown())
                        df3_split = df3_split.set_index("level_0")
                        df3_split.columns = ["Num", "TID"]
                        print(df3_split.head().to_markdown())
                        print(len(df3))
                        df3 = pd.concat([df3, df3_split])
                        print(len(df3))
                        print(df3.head().to_markdown())
                    df1 = df1[~df1["TID"].isin(df3["TID"])]
                except Exception as e:
                    print("报错或者没有df3")
                    pass
                print("账单文件删除储值卡交易记录的充值订单后行数{}".format(len(df1)))

            print(df1.tail().to_markdown())
            df1["INCOME_AMOUNT"] = df1["INCOME_AMOUNT"].astype(float).abs()
            df1["EXPEND_AMOUNT"] = -df1["EXPEND_AMOUNT"].astype(float)
            print(df1.head(5).to_markdown())
            print("有赞")

    # 小红书逻辑
    elif filename.find("小红书") >= 0:
        try:
            df = pd.read_excel(filename, sheet_name="商品销售", dtype=str)
            for column_name in df.columns:
                df.rename(columns={column_name: column_name.replace(" ", "").replace("\n", "").strip()}, inplace=True)
            # df["类型"] = df["类型"].apply(lambda x:x.replace(" ","").strip())
            print(df.head(5).to_markdown())

            # 收入总额
            plat = "XHS"
            df1 = pd.DataFrame()
            if "包裹号" in df.columns:
                df1["TID"] = df["包裹号"]
            else:
                df1["TID"] = df["订单号"]
            df1["SHOPNAME"] = ""
            df1["PLATFORM"] = plat
            df1["SHOPCODE"] = ""
            df1["BILLPLATFORM"] = plat
            if "订单完成时间" in df.columns:
                df1["CREATED"] = df["订单完成时间"]
            elif "发货时间" in df.columns:
                df1["CREATED"] = df["发货时间"]
            else:
                df1["CREATED"] = df["用户下单时间"]
            df1["TITLE"] = df["商品名称"]
            df1["TRADE_TYPE"] = "收入总额"
            df1["BUSINESS_NO"] = df["订单号"]
            df1["INCOME_AMOUNT"] = df["收入总额"]
            df1["INCOME_AMOUNT"].fillna(0, inplace=True)
            df1["EXPEND_AMOUNT"] = 0
            df1["TRADING_CHANNELS"] = ""
            df1["BUSINESS_DESCRIPTION"] = "收入总额"
            df1["remark"] = ""
            df1["BUSINESS_BILL_SOURCE"] = ""
            df1["IS_REFUNDAMOUNT"] = 0
            df1["IS_AMOUNT"] = 1
            df1["OID"] = ""
            df1["SOURCEDATA"] = "EXCEL"
            df1["RECIPROCAL_ACCOUNT"] = ""
            df1["BATCHNO"] = ""
            df1["currency"] = ""
            df1["overseas_income"] = ""
            df1["overseas_expend"] = ""
            df1["currency_cny_rate"] = ""
            df1["INCOME_AMOUNT"] = df1["INCOME_AMOUNT"].astype(float).abs()
            # df1 = df1.loc[df1["INCOME_AMOUNT"] != 0]

            # 佣金总额
            df2 = df1.copy()
            df2["TRADE_TYPE"] = "佣金总额"
            df2["INCOME_AMOUNT"] = df["佣金总额"]
            df2["EXPEND_AMOUNT"] = 0
            df2["BUSINESS_DESCRIPTION"] = "佣金总额"
            df2["IS_REFUNDAMOUNT"] = 0
            df2["IS_AMOUNT"] = 1
            df2["INCOME_AMOUNT"] = df2["INCOME_AMOUNT"].astype(float)
            # df2 = df1.loc[df2["INCOME_AMOUNT"] != 0]

            # 支付渠道费(商品)
            df3 = df1.copy()
            df3["TRADE_TYPE"] = "支付渠道费(商品)"
            df3["INCOME_AMOUNT"] = 0
            df3["EXPEND_AMOUNT"] = df["支付渠道费(商品)"]
            df3["BUSINESS_DESCRIPTION"] = "支付渠道费(商品)"
            df3["IS_REFUNDAMOUNT"] = 0
            df3["IS_AMOUNT"] = 0
            df3["EXPEND_AMOUNT"] = -df3["EXPEND_AMOUNT"].astype(float)
            # df3 = df3.loc[df1["EXPEND_AMOUNT"] != 0]
            print("小红书")

            # 运费
            df4 = df1.copy()
            df4["TRADE_TYPE"] = "运费"
            df4["INCOME_AMOUNT"] = 0
            if "运费" in df.columns:
                df4["EXPEND_AMOUNT"] = df["运费"]
            else:
                df4["EXPEND_AMOUNT"] = 0
            df4["EXPEND_AMOUNT"].fillna(0, inplace=True)
            df4["BUSINESS_DESCRIPTION"] = "订单运费"
            df4["IS_REFUNDAMOUNT"] = 0
            df4["IS_AMOUNT"] = 0
            df4["EXPEND_AMOUNT"] = -df4["EXPEND_AMOUNT"].astype(float).abs()
            # df4 = df1.loc[df4["INCOME_AMOUNT"] != 0]

            # 支付渠道费(运费)
            df5 = df1.copy()
            df5["TRADE_TYPE"] = "支付渠道费(运费)"
            df5["INCOME_AMOUNT"] = 0
            if "支付渠道费(运费)" in df.columns:
                df5["EXPEND_AMOUNT"] = df["支付渠道费(运费)"]
            else:
                df5["EXPEND_AMOUNT"] = 0
            df5["BUSINESS_DESCRIPTION"] = "支付渠道费(运费)"
            df5["IS_REFUNDAMOUNT"] = 0
            df5["IS_AMOUNT"] = 0
            df5["EXPEND_AMOUNT"] = -df5["EXPEND_AMOUNT"].astype(float).abs()
            # df3 = df3.loc[df1["EXPEND_AMOUNT"] != 0]
            print("小红书")

        except Exception as e:
            print("不符合账单格式：没有商品销售分页！")
            dict = {"TID": "nan", "SHOPNAME": "", "PLATFORM": "", "SHOPCODE": "", "BILLPLATFORM": "", "CREATED": "",
                    "TITLE": "", "TRADE_TYPE": "", "BUSINESS_NO": "", "INCOME_AMOUNT": "", "EXPEND_AMOUNT": "",
                    "TRADING_CHANNELS": "", "BUSINESS_DESCRIPTION": "", "remark": "", "BUSINESS_BILL_SOURCE": "",
                    "IS_REFUNDAMOUNT": "", "IS_AMOUNT": "", "OID": "", "SOURCEDATA": "", "RECIPROCAL_ACCOUNT": "",
                    "BATCHNO": ""}
            df1 = pd.DataFrame(dict, index=[0])
            df2 = pd.DataFrame(dict, index=[0])
            df3 = pd.DataFrame(dict, index=[0])
            df4 = pd.DataFrame(dict, index=[0])
            df5 = pd.DataFrame(dict, index=[0])

        try:
            # 退货
            df = pd.read_excel(filename, sheet_name="退货", dtype=str)
            plat = "XHS"
            df6 = pd.DataFrame()
            if "包裹号" in df.columns:
                df6["TID"] = df["包裹号"]
            else:
                df6["TID"] = df["订单号"]
            df6["SHOPNAME"] = ""
            df6["PLATFORM"] = plat
            df6["SHOPCODE"] = ""
            df6["BILLPLATFORM"] = plat
            df6["CREATED"] = df["退款时间"]
            df6["TITLE"] = df["商品名称"]
            df6["TRADE_TYPE"] = "支出总额"
            df6["BUSINESS_NO"] = df["退货单号"]
            df6["INCOME_AMOUNT"] = 0
            df6["EXPEND_AMOUNT"] = df["支出总额"]
            df6["TRADING_CHANNELS"] = ""
            df6["BUSINESS_DESCRIPTION"] = "支出总额"
            df6["remark"] = ""
            df6["BUSINESS_BILL_SOURCE"] = ""
            df6["IS_REFUNDAMOUNT"] = 1
            df6["IS_AMOUNT"] = 0
            df6["OID"] = ""
            df6["SOURCEDATA"] = "EXCEL"
            df6["RECIPROCAL_ACCOUNT"] = ""
            df6["BATCHNO"] = ""
            df6["currency"] = ""
            df6["overseas_income"] = ""
            df6["overseas_expend"] = ""
            df6["currency_cny_rate"] = ""
            df6["EXPEND_AMOUNT"] = -df6["EXPEND_AMOUNT"].astype(float)
            # df4 = df4.loc[df1["EXPEND_AMOUNT"] != 0]

            # 退货佣金
            df7 = df6.copy()
            if "佣金总额" in df.columns:
                commission_fee = "佣金总额"
            elif "退货佣金总额" in df.columns:
                commission_fee = "退货佣金总额"
            elif "销售佣金总额" in df.columns:
                commission_fee = "销售佣金总额"
            df7["TRADE_TYPE"] = "佣金总额"
            df7["INCOME_AMOUNT"] = 0
            if "佣金总额" in df.columns:
                df7["EXPEND_AMOUNT"] = df["佣金总额"]
            elif "退货佣金总额" in df.columns:
                df7["EXPEND_AMOUNT"] = df["退货佣金总额"]
            elif "销售佣金总额" in df.columns:
                df7["EXPEND_AMOUNT"] = df["销售佣金总额"]
            df7["BUSINESS_DESCRIPTION"] = "佣金总额"
            df7["IS_REFUNDAMOUNT"] = 1
            df7["IS_AMOUNT"] = 0
            df7["EXPEND_AMOUNT"] = -df7["EXPEND_AMOUNT"].astype(float)
            # df7 = df7.loc[df1["EXPEND_AMOUNT"] != 0]

            # 支付渠道费
            df8 = df6.copy()
            df8["TRADE_TYPE"] = "支付渠道费"
            df8["INCOME_AMOUNT"] = df["支付渠道费"]
            df8["EXPEND_AMOUNT"] = 0
            df8["BUSINESS_DESCRIPTION"] = "支付渠道费"
            df8["IS_REFUNDAMOUNT"] = 0
            df8["IS_AMOUNT"] = 0
            df8["INCOME_AMOUNT"] = df8["INCOME_AMOUNT"].astype(float).abs()
            # df8 = df8.loc[df1["INCOME_AMOUNT"] != 0]
            print("小红书")

        except Exception as e:
            print("没有退货账单！")
            dict = {"TID": "nan", "SHOPNAME": "", "PLATFORM": "", "SHOPCODE": "", "BILLPLATFORM": "", "CREATED": "",
                    "TITLE": "", "TRADE_TYPE": "", "BUSINESS_NO": "", "INCOME_AMOUNT": "", "EXPEND_AMOUNT": "",
                    "TRADING_CHANNELS": "", "BUSINESS_DESCRIPTION": "", "remark": "", "BUSINESS_BILL_SOURCE": "",
                    "IS_REFUNDAMOUNT": "", "IS_AMOUNT": "", "OID": "", "SOURCEDATA": "", "RECIPROCAL_ACCOUNT": "",
                    "BATCHNO": ""}
            df6 = pd.DataFrame(dict, index=[0])
            df7 = pd.DataFrame(dict, index=[0])
            df8 = pd.DataFrame(dict, index=[0])

        # 人工退款
        try:
            df = pd.read_excel(filename, sheet_name="人工退款", dtype=str)
            plat = "XHS"
            df9 = pd.DataFrame()
            if "包裹号" in df.columns:
                df9["TID"] = df["包裹号"]
            else:
                df9["TID"] = df["订单号"]
            df9["SHOPNAME"] = ""
            df9["PLATFORM"] = plat
            df9["SHOPCODE"] = ""
            df9["BILLPLATFORM"] = plat
            df9["CREATED"] = df["退款时间"]
            df9["TITLE"] = ""
            df9["TRADE_TYPE"] = "支出总额"
            df9["BUSINESS_NO"] = ""
            df9["INCOME_AMOUNT"] = 0
            df9["EXPEND_AMOUNT"] = df["支出金额"]
            df9["TRADING_CHANNELS"] = ""
            df9["BUSINESS_DESCRIPTION"] = "人工退款-支出总额"
            if "退款原因" in df.columns:
                df9["remark"] = df["退款原因"]
            else:
                df9["remark"] = ""
            df9["BUSINESS_BILL_SOURCE"] = ""
            df9["IS_REFUNDAMOUNT"] = 0
            df9["IS_AMOUNT"] = 0
            df9["OID"] = ""
            df9["SOURCEDATA"] = "EXCEL"
            df9["RECIPROCAL_ACCOUNT"] = ""
            df9["BATCHNO"] = ""
            df9["currency"] = ""
            df9["overseas_income"] = ""
            df9["overseas_expend"] = ""
            df9["currency_cny_rate"] = ""
            df9["EXPEND_AMOUNT"] = -df9["EXPEND_AMOUNT"].astype(float)
            # df9 = df9.loc[df1["EXPEND_AMOUNT"] != 0]

            # 人工退款佣金
            df10 = df9.copy()
            df10["TRADE_TYPE"] = "佣金总额"
            df10["INCOME_AMOUNT"] = 0
            df10["EXPEND_AMOUNT"] = df["佣金金额"]
            df10["BUSINESS_DESCRIPTION"] = "人工退款-佣金总额"
            df10["IS_REFUNDAMOUNT"] = 0
            df10["IS_AMOUNT"] = 0
            df10["EXPEND_AMOUNT"] = -df10["EXPEND_AMOUNT"].astype(float)
            # df8 = df8.loc[df1["EXPEND_AMOUNT"] != 0]

        except Exception as e:
            print("没有人工退款账单！")
            dict = {"TID": "nan", "SHOPNAME": "", "PLATFORM": "", "SHOPCODE": "", "BILLPLATFORM": "", "CREATED": "",
                    "TITLE": "", "TRADE_TYPE": "", "BUSINESS_NO": "", "INCOME_AMOUNT": "", "EXPEND_AMOUNT": "",
                    "TRADING_CHANNELS": "", "BUSINESS_DESCRIPTION": "", "remark": "", "BUSINESS_BILL_SOURCE": "",
                    "IS_REFUNDAMOUNT": "", "IS_AMOUNT": "", "OID": "", "SOURCEDATA": "", "RECIPROCAL_ACCOUNT": "",
                    "BATCHNO": ""}
            df9 = pd.DataFrame(dict, index=[0])
            df10 = pd.DataFrame(dict, index=[0])

        # 赔偿用户薯券
        try:
            df = pd.read_excel(filename, sheet_name="赔偿用户薯券", dtype=str)
            plat = "XHS"
            df11 = pd.DataFrame()
            if "包裹号" in df.columns:
                df11["TID"] = df["包裹号"]
            else:
                df11["TID"] = df["订单号"]
            df11["SHOPNAME"] = ""
            df11["PLATFORM"] = plat
            df11["SHOPCODE"] = ""
            df11["BILLPLATFORM"] = plat
            df11["CREATED"] = df["时间"]
            df11["TITLE"] = ""
            df11["TRADE_TYPE"] = "赔偿用户薯券"
            df11["BUSINESS_NO"] = ""
            df11["INCOME_AMOUNT"] = 0
            df11["EXPEND_AMOUNT"] = df["支出金额"]
            df11["TRADING_CHANNELS"] = ""
            df11["BUSINESS_DESCRIPTION"] = "赔偿用户薯券"
            if "原因" in df.columns:
                df11["remark"] = df["原因"]
            else:
                df11["remark"] = ""
            df11["BUSINESS_BILL_SOURCE"] = ""
            df11["IS_REFUNDAMOUNT"] = 0
            df11["IS_AMOUNT"] = 0
            df11["OID"] = ""
            df11["SOURCEDATA"] = "EXCEL"
            df11["RECIPROCAL_ACCOUNT"] = ""
            df11["BATCHNO"] = ""
            df11["currency"] = ""
            df11["overseas_income"] = ""
            df11["overseas_expend"] = ""
            df11["currency_cny_rate"] = ""
            df11["EXPEND_AMOUNT"] = -df11["EXPEND_AMOUNT"].astype(float).abs()
        except Exception as e:
            print("没有赔偿用户薯券账单！")
            dict = {"TID": "nan", "SHOPNAME": "", "PLATFORM": "", "SHOPCODE": "", "BILLPLATFORM": "", "CREATED": "",
                    "TITLE": "", "TRADE_TYPE": "", "BUSINESS_NO": "", "INCOME_AMOUNT": "", "EXPEND_AMOUNT": "",
                    "TRADING_CHANNELS": "", "BUSINESS_DESCRIPTION": "", "remark": "", "BUSINESS_BILL_SOURCE": "",
                    "IS_REFUNDAMOUNT": "", "IS_AMOUNT": "", "OID": "", "SOURCEDATA": "", "RECIPROCAL_ACCOUNT": "",
                    "BATCHNO": ""}
            df11 = pd.DataFrame(dict, index=[0])

        # 人工调账
        try:
            # 结算
            df = pd.read_excel(filename, sheet_name="人工调账", dtype=str)
            file = os.sep.join(filename.split(os.sep)[-1:])
            if file.find("-") >= 0:
                print("定位1")
                if ((file.find("1901") >= 0) | (file.find("1903") >= 0) | (file.find("1905") >= 0) | (
                        file.find("1907") >= 0) | (file.find("1908") >= 0) | (file.find("1910") >= 0) | (
                        file.find("1912") >= 0)):
                    time = ("".join(file.split("-")[:1]))[-6:] + "31 23:59:59"
                elif ((file.find("1904") >= 0) | (file.find("1906") >= 0) | (file.find("1909") >= 0) | (
                        file.find("1911") >= 0)):
                    time = ("".join(file.split("-")[:1]))[-6:] + "30 23:59:59"
                else:
                    time = ("".join(file.split("-")[:1]))[-6:] + "28 23:59:59"
            else:
                if file.find("2019") >= 0:
                    print("定位2")
                    if ((file.find("1901") >= 0) | (file.find("1903") >= 0) | (file.find("1905") >= 0) | (
                            file.find("1907") >= 0) | (file.find("1908") >= 0) | (file.find("1910") >= 0) | (
                            file.find("1912") >= 0)):
                        time = ("".join(file.split(".")[:1]))[-6:] + "31 23:59:59"
                    elif ((file.find("1904") >= 0) | (file.find("1906") >= 0) | (file.find("1909") >= 0) | (
                            file.find("1911") >= 0)):
                        time = ("".join(file.split(".")[:1]))[-6:] + "30 23:59:59"
                    else:
                        time = ("".join(file.split(".")[:1]))[-6:] + "28 23:59:59"
                else:
                    print("定位3")
                    if ((file.find("1901") >= 0) | (file.find("1903") >= 0) | (file.find("1905") >= 0) | (
                            file.find("1907") >= 0) | (file.find("1908") >= 0) | (file.find("1910") >= 0) | (
                            file.find("1912") >= 0)):
                        time = "20" + ("".join(file.split(".")[:1]))[-4:] + "31 23:59:59"
                    elif ((file.find("1904") >= 0) | (file.find("1906") >= 0) | (file.find("1909") >= 0) | (
                            file.find("1911") >= 0)):
                        time = "20" + ("".join(file.split(".")[:1]))[-4:] + "30 23:59:59"
                    else:
                        time = "20" + ("".join(file.split(".")[:1]))[-4:] + "28 23:59:59"
                    # time = "20" + (("".join(file.split(".")[:1]))[-4:]) + "01 00:00:00"

                # if ((yearmonth[-2:].find("02")>=0)&(yearmonth.find("2020")>=0)):
                #     time = yearmonth + "29 23:59:59"
                # if ((yearmonth[-2:].find("02")>=0)&(yearmonth.find("2020")<0)):
                #     time = yearmonth + "28 23:59:59"
                # elif ((yearmonth[-2:].find("04")>=0)|(yearmonth[-2:].find("06")>=0)|(yearmonth[-2:].find("09")>=0)|(yearmonth[-2:].find("11")>=0)):
                #     time = yearmonth + "30 23:59:59"
                # else:
                #     time = yearmonth + "31 23:59:59"
            plat = "XHS"
            df12 = pd.DataFrame()
            df12["TID"] = df["结算单号"]
            df12["SHOPNAME"] = ""
            df12["PLATFORM"] = plat
            df12["SHOPCODE"] = ""
            df12["BILLPLATFORM"] = plat
            df12["CREATED"] = time
            df12["TITLE"] = ""
            df12["TRADE_TYPE"] = "人工调账"
            df12["BUSINESS_NO"] = ""
            df12["INCOME_AMOUNT"] = df.apply(lambda x: x["结算总额"] if x["类型"] == "收入" else 0, axis=1)
            df12["EXPEND_AMOUNT"] = df.apply(lambda x: x["结算总额"] if x["类型"] == "支出" else 0, axis=1)
            df12["TRADING_CHANNELS"] = ""
            df12["BUSINESS_DESCRIPTION"] = "人工调账-结算"
            df12["remark"] = df["说明"]
            df12["BUSINESS_BILL_SOURCE"] = ""
            df12["IS_REFUNDAMOUNT"] = 0
            df12["IS_AMOUNT"] = 0
            df12["OID"] = ""
            df12["SOURCEDATA"] = "EXCEL"
            df12["RECIPROCAL_ACCOUNT"] = ""
            df12["BATCHNO"] = ""
            df12["currency"] = ""
            df12["overseas_income"] = ""
            df12["overseas_expend"] = ""
            df12["currency_cny_rate"] = ""
            df12["INCOME_AMOUNT"] = df12["INCOME_AMOUNT"].astype(float).abs()
            df12["EXPEND_AMOUNT"] = -df12["EXPEND_AMOUNT"].astype(float).abs()

            # 佣金
            df13 = df12.copy()
            df13["INCOME_AMOUNT"] = df.apply(lambda x: x["佣金"] if x["类型"] == "支出" else 0, axis=1)
            df13["EXPEND_AMOUNT"] = df.apply(lambda x: x["佣金"] if x["类型"] == "收入" else 0, axis=1)
            df13["BUSINESS_DESCRIPTION"] = "人工调账-佣金"
            df13["INCOME_AMOUNT"] = df13["INCOME_AMOUNT"].astype(float).abs()
            df13["EXPEND_AMOUNT"] = -df13["EXPEND_AMOUNT"].astype(float).abs()

        except Exception as e:
            print("没有人工调账！")
            dict = {"TID": "nan", "SHOPNAME": "", "PLATFORM": "", "SHOPCODE": "", "BILLPLATFORM": "", "CREATED": "",
                    "TITLE": "", "TRADE_TYPE": "", "BUSINESS_NO": "", "INCOME_AMOUNT": "", "EXPEND_AMOUNT": "",
                    "TRADING_CHANNELS": "", "BUSINESS_DESCRIPTION": "", "remark": "", "BUSINESS_BILL_SOURCE": "",
                    "IS_REFUNDAMOUNT": "", "IS_AMOUNT": "", "OID": "", "SOURCEDATA": "", "RECIPROCAL_ACCOUNT": "",
                    "BATCHNO": ""}
            df12 = pd.DataFrame(dict, index=[0])
            df13 = pd.DataFrame(dict, index=[0])

        dfs = [df1, df2, df3, df4, df5, df6, df7, df8, df9, df10, df11, df12, df13]
        df1 = pd.concat(dfs)
        print(len(df1))
        df1 = df1.loc[~((df1.INCOME_AMOUNT == 0) & (df1.EXPEND_AMOUNT == 0))]
        df1["CREATED"] = df1["CREATED"].astype("datetime64[ns]")
        # df1 = df1[~df1["TID"].str.contains("nan")]
        df1.dropna(subset=["TID"], inplace=True)
        if filename.find("BodyAid") >= 0:
            df1["SHOPCODE"] = "bodyaidqjd"
            df1["SHOPNAME"] = "BodyAid旗舰店"
        elif filename.find("UNIX海外") >= 0:
            df1["SHOPCODE"] = "unixylsovppd"
            df1["SHOPNAME"] = "unix优丽氏海外品牌店"
        elif ((filename.find("UNIX优丽氏") >= 0) or (filename.find("UNIX专卖") >= 0)):
            df1["SHOPCODE"] = "unixzmd"
            df1["SHOPNAME"] = "unix优丽氏品牌店"
        elif filename.find("ultra") >= 0:
            df1["SHOPCODE"] = "ultradexppd"
            df1["SHOPNAME"] = "ultradex品牌店"
        elif filename.find("LOSHI") >= 0:
            df1["SHOPCODE"] = "loshippd"
            df1["SHOPNAME"] = "loshi品牌店"
        elif filename.find("Denty") >= 0:
            df1["SHOPCODE"] = "dentylactive"
            df1["SHOPNAME"] = "dentylactive旗舰店"
        elif filename.find("Goat") >= 0:
            df1["SHOPCODE"] = "goatsoapmklzmd"
            df1["SHOPNAME"] = "goatsoap麦凯莱专卖店"
        elif filename.find("morei") >= 0:
            df1["SHOPCODE"] = "moreiov"
            df1["SHOPNAME"] = "morei海外旗舰店"
        elif ((filename.find("Samou") >= 0) or (filename.find("samou") >= 0)):
            df1["SHOPCODE"] = "samouraiwomanovppd"
            df1["SHOPNAME"] = "samouraiwoman海外品牌店"
        elif filename.find("Dicora") >= 0:
            df1["SHOPCODE"] = "dicoraurbanfitovppd"
            df1["SHOPNAME"] = "dicoraurbanfit海外品牌店"
        elif filename.find("iwhite旗舰店") >= 0:
            df1["SHOPCODE"] = "iwhite"
            df1["SHOPNAME"] = "iwhite旗舰店"
        elif filename.find("LCN品牌") >= 0:
            df1["SHOPCODE"] = "lcnppd"
            df1["SHOPNAME"] = "lcn品牌店"
        elif filename.find("LCN海外") >= 0:
            df1["SHOPCODE"] = "lcnvoppd"
            df1["SHOPNAME"] = "lcn海外品牌店"
        elif filename.find("Lilac") >= 0:
            df1["SHOPCODE"] = "lilacovppd"
            df1["SHOPNAME"] = "lilac海外品牌店"
        elif ((filename.find("mades") >= 0) or (filename.find("MADES") >= 0)):
            df1["SHOPCODE"] = "madesppdov"
            df1["SHOPNAME"] = "mades海外品牌店"
        elif ((filename.find("Smile") >= 0) or (filename.find("smile") >= 0)):
            df1["SHOPCODE"] = "smilelab"
            df1["SHOPNAME"] = "smilelab旗舰店"
        elif filename.find("SWISSIMAGE") >= 0:
            df1["SHOPCODE"] = "swissimagevoppd"
            df1["SHOPNAME"] = "swissimage海外品牌店"
        elif ((filename.find("Chiara") >= 0) or (filename.find("Ambra") >= 0)):
            df1["SHOPCODE"] = "chiarabcaambravoppd"
            df1["SHOPNAME"] = "chiarabcaambra海外品牌店"
        elif filename.find("魔法符号") >= 0:
            df1["SHOPCODE"] = "mffhppd"
            df1["SHOPNAME"] = "魔法符号品牌店"
        elif filename.find("AllNaturalAdvice肌先知旗舰店") >= 0:
            df1["SHOPCODE"] = "anajxz"
            df1["SHOPNAME"] = "All Natural Advice肌先知旗舰店"
        elif filename.find("樱加美旗舰店") >= 0:
            df1["SHOPCODE"] = "yjmqjd"
            df1["SHOPNAME"] = "樱加美旗舰店"
        elif filename.find("樱语旗舰店") >= 0:
            df1["SHOPCODE"] = "yyqjd"
            df1["SHOPNAME"] = "樱语旗舰店"
        elif filename.find("惠优购的店") >= 0:
            df1["SHOPCODE"] = "hygdd"
            df1["SHOPNAME"] = "惠优购的店"
        elif filename.find("你莫愁旗舰店") >= 0:
            df1["SHOPCODE"] = "nmcqjd"
            df1["SHOPNAME"] = "你莫愁旗舰店"

        print(len(df1))
        print(df1.head(5).to_markdown())
        print("小红书")

    # 拼多多逻辑
    elif filename.find("拼多多") >= 0:
        df = pd.read_csv(filename, skiprows=4, dtype=str, encoding="gb18030")
        for column_name in df.columns:
            df.rename(columns={column_name: column_name.replace(" ", "").replace("\n", "").strip()}, inplace=True)
        df["账务类型"] = df["账务类型"].replace(" ", "")
        df["收入金额（+元）"] = df["收入金额（+元）"].astype(float)
        df["支出金额（-元）"] = df["支出金额（-元）"].astype(float)
        if filename.find("海外") >= 0:
            df["收入金额（+$）"] = df["收入金额（+$）"].astype(float)
            df["支出金额（-$）"] = df["支出金额（-$）"].astype(float)
        df["商户订单号"] = df["商户订单号"].astype(str)
        df = df[~df["商户订单号"].str.contains("#")]
        df = df[~df["商户订单号"].str.contains("结算汇总|提现汇总")]
        df.dropna(subset=["商户订单号", "账务类型"], inplace=True)
        df["备注"] = df["备注"].str.replace("：", ":")
        df["商户订单号"] = df.apply(lambda x: x["商户订单号"] if ((len(x["商户订单号"]) > 3) | (x["备注"].find("金额") > 0)) else "".join(
            x["备注"].split(":")[1:]), axis=1)
        print(df.head(5).to_markdown())
        plat = "PDD"
        df1 = pd.DataFrame()
        df1["TID"] = df.apply(lambda x: ("".join(x["备注"].split("关联批次号为")[-1:])) if (
                    (x["账务类型"] == "其他") & (x["备注"].find("批次订单") >= 0) & (x["备注"].find("共计汇入金额") >= 0)) else x["商户订单号"],
                              axis=1)
        df1["SHOPNAME"] = ""
        df1["PLATFORM"] = plat
        df1["SHOPCODE"] = ""
        df1["BILLPLATFORM"] = plat
        df1["CREATED"] = df["发生时间"]
        df1["TITLE"] = ""
        df1["TRADE_TYPE"] = df["账务类型"]
        df1["BUSINESS_NO"] = ""
        df1["INCOME_AMOUNT"] = df["收入金额（+元）"]
        # df1["收入"].fillna(0, inplace=True)
        df1["EXPEND_AMOUNT"] = df["支出金额（-元）"]
        # df1["支出"].fillna(0, inplace=True)
        df1["TRADING_CHANNELS"] = ""
        if filename.find("海外") >= 0:
            df1["BUSINESS_DESCRIPTION"] = df.apply(lambda x: x["账务类型"] + "-" + x["备注"] if x["备注"] != "-" else x["账务类型"],
                                                   axis=1)
        else:
            df1["BUSINESS_DESCRIPTION"] = df["账务类型"]
        df1["remark"] = df["备注"]
        df1["BUSINESS_BILL_SOURCE"] = ""
        df1["IS_REFUNDAMOUNT"] = df.apply(
            lambda x: 1 if ((x["账务类型"] == "退款") | ((x["账务类型"] == "优惠券结算") & (x["支出金额（-元）"] < 0))) else 0, axis=1)
        df1["IS_AMOUNT"] = df.apply(lambda x: 1 if (
                    (x["账务类型"] == "交易收入") | ((x["账务类型"] == "优惠券结算") & (x["收入金额（+元）"] > 0)) | (
                        (x["账务类型"] == "其他") & (x["备注"].find("批次订单") >= 0) & (x["备注"].find("共计汇入金额") >= 0))) else 0,
                                    axis=1)
        df1["OID"] = ""
        df1["SOURCEDATA"] = "EXCEL"
        df1["RECIPROCAL_ACCOUNT"] = ""
        df1["BATCHNO"] = ""
        df1["INCOME_AMOUNT"] = df1["INCOME_AMOUNT"].astype(float).abs()
        df1["EXPEND_AMOUNT"] = -df1["EXPEND_AMOUNT"].astype(float).abs()

        df1["TID"] = df1["TID"].astype(str)
        df1 = df1[~df1["TID"].str.contains("nan")]

        if filename.find("若蘅旗舰店") >= 0:
            df1["SHOPCODE"] = "yhjjd"
            df1["SHOPNAME"] = "若蘅旗舰店"
        elif filename.find("若蘅美容个护") >= 0:
            df1["SHOPCODE"] = "rhmr"
            df1["SHOPNAME"] = "拼多多若蘅美容个护专营店"
        elif filename.find("若蘅美妆") >= 0:
            df1["SHOPCODE"] = "rhmz"
            df1["SHOPNAME"] = "拼多多若蘅美妆专营店"
        elif filename.find("smilelab") >= 0:
            df1["SHOPCODE"] = "smilelab"
            df1["SHOPNAME"] = "smilelab旗舰店"
        elif filename.find("vitabloom") >= 0:
            df1["SHOPCODE"] = "vitabloom"
            df1["SHOPNAME"] = "vitabloom旗舰店"
        elif filename.find("乐丝") >= 0:
            df1["SHOPCODE"] = "loshi"
            df1["SHOPNAME"] = "乐丝旗舰店"
        elif filename.find("麦凯莱美妆") >= 0:
            df1["SHOPCODE"] = "mklmzzyd"
            df1["SHOPNAME"] = "麦凯莱美妆专营店"
        elif filename.find("LILAC海外") >= 0:
            df1["SHOPCODE"] = "lilacvoqjd"
            df1["SHOPNAME"] = "lilac海外旗舰店"
        elif filename.find("联合海外") >= 0:
            df1["SHOPCODE"] = "mjlhovzyd"
            df1["SHOPNAME"] = "卖家联合海外专营店"
        elif filename.find("卖家联合家清") >= 0:
            df1["SHOPCODE"] = "mjlhjq"
            df1["SHOPNAME"] = "卖家联合家清专营店"
        elif filename.find("樱语旗舰店") >= 0:
            df1["SHOPCODE"] = "yyqjd"
            df1["SHOPNAME"] = "樱语旗舰店"
        elif filename.find("博滴美容个护") >= 0:
            df1["SHOPCODE"] = "bdmrgh"
            df1["SHOPNAME"] = "博滴美容个护官方旗舰店"
        elif filename.find("montooth旗舰店") >= 0:
            df1["SHOPCODE"] = "montooth"
            df1["SHOPNAME"] = "montooth旗舰店"
        elif filename.find("植之璨美容个护") >= 0:
            df1["SHOPCODE"] = "zzc"
            df1["SHOPNAME"] = "植之璨美容个护专营店"
        elif filename.find("博滴旗舰店") >= 0:
            df1["SHOPCODE"] = "bdqjd"
            df1["SHOPNAME"] = "博滴旗舰店"
        elif filename.find("Dentyl") >= 0:
            df1["SHOPCODE"] = "dentylactive"
            df1["SHOPNAME"] = "dentylactive旗舰店"
        elif filename.find("卖家联合居家") >= 0:
            df1["SHOPCODE"] = "mjlh"
            df1["SHOPNAME"] = "卖家联合居家日用品专营店"
        elif filename.find("卖家联合日化") >= 0:
            df1["SHOPCODE"] = "mjlhrh"
            df1["SHOPNAME"] = "卖家联合日化专营店"
        elif filename.find("BODYAID旗舰店") >= 0:
            df1["SHOPCODE"] = "bqjd"
            df1["SHOPNAME"] = "bodyaid旗舰店"
        elif filename.find("MOREI旗舰店") >= 0:
            df1["SHOPCODE"] = "mqjd"
            df1["SHOPNAME"] = "morei旗舰店"
        elif filename.find("MOREI植之璨") >= 0:
            df1["SHOPCODE"] = "morei"
            df1["SHOPNAME"] = "MOREI植之璨专卖店"
        elif filename.find("植之璨美妆") >= 0:
            df1["SHOPCODE"] = "zzcmz"
            df1["SHOPNAME"] = "拼多多植之璨美妆专营店"
        elif filename.find("拼多多BODYAID家居生活旗舰店") >= 0:
            df1["SHOPCODE"] = "bjjsh"
            df1["SHOPNAME"] = "拼多多BODYAID家居生活旗舰店"
        elif filename.find("植之璨洗护专营店") >= 0:
            df1["SHOPCODE"] = "zzcxh"
            df1["SHOPNAME"] = "拼多多植之璨洗护专营店"
        elif filename.find("芭葆兔化妆品专营店") >= 0:
            df1["SHOPCODE"] = "bbthzp"
            df1["SHOPNAME"] = "芭葆兔化妆品专营店"
        elif filename.find("芭葆兔美容个护专营店") >= 0:
            df1["SHOPCODE"] = "bbtmrgh"
            df1["SHOPNAME"] = "芭葆兔美容个护专营店"
        elif filename.find("芭葆兔美容护肤专营") >= 0:
            df1["SHOPCODE"] = "bbtmrhfzy"
            df1["SHOPNAME"] = "芭葆兔美容护肤专营"
        elif filename.find("芭葆兔美妆专营店") >= 0:
            df1["SHOPCODE"] = "bbtmzzy"
            df1["SHOPNAME"] = "芭葆兔美妆专营店"
        elif filename.find("宝贝港湾家居生活专营店") >= 0:
            df1["SHOPCODE"] = "bbgwjjsh"
            df1["SHOPNAME"] = "宝贝港湾家居生活专营店"
        elif filename.find("宝贝港湾家居专营店") >= 0:
            df1["SHOPCODE"] = "bbgwjj"
            df1["SHOPNAME"] = "宝贝港湾家居专营店"
        elif filename.find("宝贝魔术师护肤专营店") >= 0:
            df1["SHOPCODE"] = "bbmsshf"
            df1["SHOPNAME"] = "宝贝魔术师护肤专营店"
        elif filename.find("宝贝魔术师化妆品专营店") >= 0:
            df1["SHOPCODE"] = "bbmsshzp"
            df1["SHOPNAME"] = "宝贝魔术师化妆品专营店"
        elif filename.find("宝贝魔术师美妆专营店") >= 0:
            df1["SHOPCODE"] = "bbmssmz"
            df1["SHOPNAME"] = "宝贝魔术师美妆专营店"
        elif filename.find("贝贝港湾美容个护专营店") >= 0:
            df1["SHOPCODE"] = "bbgwmrgh"
            df1["SHOPNAME"] = "贝贝港湾美容个护专营店"
        elif filename.find("贝贝港湾美妆专营店") >= 0:
            df1["SHOPCODE"] = "bbgwmz"
            df1["SHOPNAME"] = "贝贝港湾美妆专营店"
        elif filename.find("补舍美容个护专营店") >= 0:
            df1["SHOPCODE"] = "bsmrgh"
            df1["SHOPNAME"] = "补舍美容个护专营店"
        elif filename.find("补舍美妆专营店") >= 0:
            df1["SHOPCODE"] = "bsmz"
            df1["SHOPNAME"] = "补舍美妆专营店"
        elif filename.find("航星彩妆专营店") >= 0:
            df1["SHOPCODE"] = "hxcz"
            df1["SHOPNAME"] = "航星彩妆专营店"
        elif filename.find("航星个护专营店") >= 0:
            df1["SHOPCODE"] = "hxgh"
            df1["SHOPNAME"] = "航星个护专营店"
        elif filename.find("航星护肤品专营店") >= 0:
            df1["SHOPCODE"] = "hxhfpzy"
            df1["SHOPNAME"] = "航星护肤品专营店"
        elif filename.find("航星护肤专营") >= 0:
            df1["SHOPCODE"] = "hxhfzy"
            df1["SHOPNAME"] = "航星护肤专营"
        elif filename.find("航星化妆品专营") >= 0:
            df1["SHOPCODE"] = "hxhzpzy"
            df1["SHOPNAME"] = "航星化妆品专营"
        elif filename.find("航星美容专营店") >= 0:
            df1["SHOPCODE"] = "hxmrzy"
            df1["SHOPNAME"] = "航星美容专营店"
        elif filename.find("控师护肤品专营店") >= 0:
            df1["SHOPCODE"] = "kshf"
            df1["SHOPNAME"] = "控师护肤品专营店"
        elif filename.find("无极爽护肤品专营店") >= 0:
            df1["SHOPCODE"] = "wjshfp"
            df1["SHOPNAME"] = "无极爽护肤品专营店"
        elif filename.find("无极爽护肤专营店") >= 0:
            df1["SHOPCODE"] = "wjshf"
            df1["SHOPNAME"] = "无极爽护肤专营店"
        elif filename.find("无极爽美容个护专营店") >= 0:
            df1["SHOPCODE"] = "wjsmrgh"
            df1["SHOPNAME"] = "无极爽美容个护专营店"
        elif filename.find("无极爽美妆专营店") >= 0:
            df1["SHOPCODE"] = "wjsmz"
            df1["SHOPNAME"] = "无极爽美妆专营店"
        elif filename.find("戏酱美容个护专营店") >= 0:
            df1["SHOPCODE"] = "xjmrgh"
            df1["SHOPNAME"] = "戏酱美容个护专营店"
        elif filename.find("戏酱美妆专营店") >= 0:
            df1["SHOPCODE"] = "xjmz"
            df1["SHOPNAME"] = "戏酱美妆专营店"
        elif filename.find("博滴家居旗舰店") >= 0:
            df1["SHOPCODE"] = "bdjjqjd"
            df1["SHOPNAME"] = "博滴家居旗舰店"
        elif filename.find("麦凯莱美妆专营店") >= 0:
            df1["SHOPNAME"] = "麦凯莱美妆专营店"
            df1["SHOPCODE"] = "mklmzzyd"
        elif filename.find("若蘅旗舰店") >= 0:
            df1["SHOPNAME"] = "若蘅旗舰店"
            df1["SHOPCODE"] = "yhjjd"
        elif ((filename.find("dentylactive旗舰店") >= 0) | (filename.find("Dentyl Active旗舰店") >= 0)):
            df1["SHOPNAME"] = "dentylactive旗舰店"
            df1["SHOPCODE"] = "dentylactive"
        elif filename.find("ontooth") >= 0:
            df1["SHOPNAME"] = "montooth旗舰店"
            df1["SHOPCODE"] = "montooth"
        elif filename.find("vitabloom旗舰店") >= 0:
            df1["SHOPNAME"] = "vitabloom旗舰店"
            df1["SHOPCODE"] = "vitabloom"
        elif filename.find("樱语旗舰店") >= 0:
            df1["SHOPNAME"] = "樱语旗舰店"
            df1["SHOPCODE"] = "yyqjd"
        elif filename.find("博滴旗舰店") >= 0:
            df1["SHOPNAME"] = "博滴旗舰店"
            df1["SHOPCODE"] = "bdqjd"
        elif filename.find("博滴家居旗舰店") >= 0:
            df1["SHOPNAME"] = "博滴家居旗舰店"
            df1["SHOPCODE"] = "bdjjqjd"
        elif filename.find("smilelab旗舰店") >= 0:
            df1["SHOPNAME"] = "smilelab旗舰店"
            df1["SHOPCODE"] = "smilelab"
        elif ((filename.find("bodyaid") >= 0) | (filename.find("BODYAID家居生活旗舰店") >= 0) | (
                filename.find("BODYAID旗舰店") >= 0)):
            df1["SHOPNAME"] = "bodyaid旗舰店"
            df1["SHOPCODE"] = "bqjd"
        elif ((filename.find("morei旗舰店") >= 0) | (filename.find("MOREI旗舰店") >= 0)):
            df1["SHOPNAME"] = "morei旗舰店"
            df1["SHOPCODE"] = "mqjd"
        elif filename.find("MOREI植之璨") >= 0:
            df1["SHOPNAME"] = "MOREI植之璨专卖店"
            df1["SHOPCODE"] = "morei"
        elif filename.find("博滴美容个护官方旗舰店") >= 0:
            df1["SHOPNAME"] = "博滴美容个护官方旗舰店"
            df1["SHOPCODE"] = "bdmrgh"
        elif filename.find("卖家联合日化专营店") >= 0:
            df1["SHOPNAME"] = "卖家联合日化专营店"
            df1["SHOPCODE"] = "mjlhrh"
        elif filename.find("卖家联合家清") >= 0:
            df1["SHOPNAME"] = "卖家联合家清专营店"
            df1["SHOPCODE"] = "mjlhjq"
        elif filename.find("卖家联合居家日用") >= 0:
            df1["SHOPNAME"] = "卖家联合居家日用品专营店"
            df1["SHOPCODE"] = "mjlh"
        elif filename.find("卖家联合海外") >= 0:
            df1["SHOPNAME"] = "卖家联合海外专营店"
            df1["SHOPCODE"] = "mjlhovzyd"
        elif filename.find("若蘅美妆专营店") >= 0:
            df1["SHOPNAME"] = "拼多多若蘅美妆专营店"
            df1["SHOPCODE"] = "rhmz"
        elif filename.find("若蘅美容个护") >= 0:
            df1["SHOPNAME"] = "拼多多若蘅美容个护专营店"
            df1["SHOPCODE"] = "rhmr"
        elif filename.find("植之璨美容个护专营店") >= 0:
            df1["SHOPNAME"] = "植之璨美容个护专营店"
            df1["SHOPCODE"] = "zzc"
        elif filename.find("植之璨美妆") >= 0:
            df1["SHOPNAME"] = "植之璨美妆专营店"
            df1["SHOPCODE"] = "zzcmz"
        elif filename.find("植之璨洗护") >= 0:
            df1["SHOPNAME"] = "植之璨洗护专营店"
            df1["SHOPCODE"] = "zzcxh"
        elif filename.find("芭葆兔化妆品") >= 0:
            df1["SHOPNAME"] = "芭葆兔化妆品专营店"
            df1["SHOPCODE"] = "bbthzp"
        elif filename.find("芭葆兔美妆专营店") >= 0:
            df1["SHOPNAME"] = "芭葆兔美妆专营店"
            df1["SHOPCODE"] = "bbtmzzy"
        elif filename.find("宝贝港湾家居生活") >= 0:
            df1["SHOPNAME"] = "宝贝港湾家居生活专营店"
            df1["SHOPCODE"] = "bbgwjjsh"
        elif filename.find("宝贝港湾家居专营店") >= 0:
            df1["SHOPNAME"] = "宝贝港湾家居专营店"
            df1["SHOPCODE"] = "bbgwjj"
        elif filename.find("宝贝魔术师护肤专营店") >= 0:
            df1["SHOPNAME"] = "宝贝魔术师护肤专营店"
            df1["SHOPCODE"] = "bbmsshf"
        elif filename.find("宝贝魔术师化妆品专营店") >= 0:
            df1["SHOPNAME"] = "宝贝魔术师化妆品专营店"
            df1["SHOPCODE"] = "bbmsshzp"
        elif filename.find("宝贝魔术师美妆专营店") >= 0:
            df1["SHOPNAME"] = "宝贝魔术师美妆专营店"
            df1["SHOPCODE"] = "bbmssmz"
        elif filename.find("贝贝港湾美容个护专营店") >= 0:
            df1["SHOPNAME"] = "贝贝港湾美容个护专营店"
            df1["SHOPCODE"] = "bbgwmrgh"
        elif filename.find("贝贝港湾美妆专营店") >= 0:
            df1["SHOPNAME"] = "贝贝港湾美妆专营店"
            df1["SHOPCODE"] = "bbgwmz"
        elif filename.find("补舍美容个护专营店") >= 0:
            df1["SHOPNAME"] = "补舍美容个护专营店"
            df1["SHOPCODE"] = "bsmrgh"
        elif filename.find("补舍美妆专营店") >= 0:
            df1["SHOPNAME"] = "补舍美妆专营店"
            df1["SHOPCODE"] = "bsmz"
        elif filename.find("航星彩妆专营店") >= 0:
            df1["SHOPNAME"] = "航星彩妆专营店"
            df1["SHOPCODE"] = "hxcz"
        elif filename.find("航星个护专营店") >= 0:
            df1["SHOPNAME"] = "航星个护专营店"
            df1["SHOPCODE"] = "hxgh"
        elif filename.find("航星护肤品专营店") >= 0:
            df1["SHOPNAME"] = "航星护肤品专营店"
            df1["SHOPCODE"] = "hxhfpzy"
        elif filename.find("航星护肤专营") >= 0:
            df1["SHOPNAME"] = "航星护肤专营"
            df1["SHOPCODE"] = "hxhfzy"
        elif filename.find("航星化妆品专营") >= 0:
            df1["SHOPNAME"] = "航星化妆品专营"
            df1["SHOPCODE"] = "hxhzpzy"
        elif filename.find("航星美容专营店") >= 0:
            df1["SHOPNAME"] = "航星美容专营店"
            df1["SHOPCODE"] = "hxmrzy"
        elif filename.find("控师护肤品专营店") >= 0:
            df1["SHOPNAME"] = "控师护肤品专营店"
            df1["SHOPCODE"] = "kshf"
        elif filename.find("无极爽护肤品专营店") >= 0:
            df1["SHOPNAME"] = "无极爽护肤品专营店"
            df1["SHOPCODE"] = "wjshfp"
        elif filename.find("无极爽护肤专营店") >= 0:
            df1["SHOPNAME"] = "无极爽护肤专营店"
            df1["SHOPCODE"] = "wjshf"
        elif filename.find("无极爽美容个护专营店") >= 0:
            df1["SHOPNAME"] = "无极爽美容个护专营店"
            df1["SHOPCODE"] = "wjsmrgh"
        elif filename.find("无极爽美妆专营店") >= 0:
            df1["SHOPNAME"] = "无极爽美妆专营店"
            df1["SHOPCODE"] = "wjsmz"
        elif filename.find("戏酱美容个护专营店") >= 0:
            df1["SHOPNAME"] = "戏酱美容个护专营店"
            df1["SHOPCODE"] = "xjmrgh"
        elif filename.find("戏酱美妆专营店") >= 0:
            df1["SHOPNAME"] = "戏酱美妆专营店"
            df1["SHOPCODE"] = "xjmz"
        elif filename.find("芭葆兔美容个护专营店") >= 0:
            df1["SHOPNAME"] = "芭葆兔美容个护专营店"
            df1["SHOPCODE"] = "bbtmrgh"
        elif filename.find("一片珍芯日化专营店") >= 0:
            df1["SHOPNAME"] = "一片珍芯日化专营店"
            df1["SHOPCODE"] = "ypzxrh"
        elif filename.find("一片珍芯美妆专营店") >= 0:
            df1["SHOPNAME"] = "一片珍芯美妆专营店"
            df1["SHOPCODE"] = "ypzxmz"
        elif filename.find("一片珍芯家清专营店") >= 0:
            df1["SHOPNAME"] = "一片珍芯家清专营店"
            df1["SHOPCODE"] = "ypzxjq"
        elif filename.find("一片珍芯家居专营店") >= 0:
            df1["SHOPNAME"] = "一片珍芯家居专营店"
            df1["SHOPCODE"] = "ypzxjj"
        elif filename.find("一片珍芯家居生活专营店") >= 0:
            df1["SHOPNAME"] = "一片珍芯家居生活专营店"
            df1["SHOPCODE"] = "ypzxjjsh"
        elif filename.find("铲喜官护肤专营店") >= 0:
            df1["SHOPNAME"] = "铲喜官护肤专营店"
            df1["SHOPCODE"] = "cxghf"
        elif filename.find("铲喜官美妆专营店") >= 0:
            df1["SHOPNAME"] = "铲喜官美妆专营店"
            df1["SHOPCODE"] = "cxgmz"
        elif filename.find("铲喜官化妆品专营店") >= 0:
            df1["SHOPNAME"] = "铲喜官化妆品专营店"
            df1["SHOPCODE"] = "cxghzp"
        elif filename.find("控师美容专营店") >= 0:
            df1["SHOPNAME"] = "控师美容专营店"
            df1["SHOPCODE"] = "ksmr"
        elif filename.find("控师护肤品专营店") >= 0:
            df1["SHOPNAME"] = "控师护肤品专营店"
            df1["SHOPCODE"] = "kshf"
        elif filename.find("控师护肤专营店") >= 0:
            df1["SHOPNAME"] = "控师护肤专营店"
            df1["SHOPCODE"] = "kshfzyd"
        elif filename.find("控师化妆品专营店") >= 0:
            df1["SHOPNAME"] = "控师化妆品专营店"
            df1["SHOPCODE"] = "kshzp"
        elif filename.find("控师美容个护专营店") >= 0:
            df1["SHOPNAME"] = "控师美容个护专营店"
            df1["SHOPCODE"] = "ksmrgh"
        elif filename.find("茉小桃护肤品专营店") >= 0:
            df1["SHOPNAME"] = "茉小桃护肤品专营店"
            df1["SHOPCODE"] = "zxthfp"
        elif filename.find("茉小桃美容护肤专营店") >= 0:
            df1["SHOPNAME"] = "茉小桃美容护肤专营店"
            df1["SHOPCODE"] = "zxtmrhf"
        elif filename.find("茉小桃化妆品专营店") >= 0:
            df1["SHOPNAME"] = "茉小桃化妆品专营店"
            df1["SHOPCODE"] = "zxthzp"
        elif filename.find("茉小桃护肤专营店") >= 0:
            df1["SHOPNAME"] = "茉小桃护肤专营店"
            df1["SHOPCODE"] = "mxthfzy"
        elif filename.find("茉小桃美妆专营店") >= 0:
            df1["SHOPNAME"] = "茉小桃护肤品专营店"
            df1["SHOPCODE"] = "zxtmz"
        elif filename.find("芭葆兔美容护肤专营") >= 0:
            df1["SHOPNAME"] = "芭葆兔美容护肤专营"
            df1["SHOPCODE"] = "bbtmrhfzy"
        elif filename.find("麦凯莱美妆专营") >= 0:
            df1["SHOPNAME"] = "麦凯莱美妆专营店"
            df1["SHOPCODE"] = "mklmzzyd"
        elif filename.find("ALL NATURAL ADVICE美容个护旗舰店") >= 0:
            df1["SHOPNAME"] = "ALL NATURAL ADVICE美容个护旗舰店"
            df1["SHOPCODE"] = "amrghqjd"
        elif filename.find("宝贝魔术师个护专营店") >= 0:
            df1["SHOPNAME"] = "宝贝魔术师个护专营店"
            df1["SHOPCODE"] = "bbmssghzyd"
        elif filename.find("宝贝魔术师护肤品专营店") >= 0:
            df1["SHOPNAME"] = "宝贝魔术师护肤品专营店"
            df1["SHOPCODE"] = "bbmsshfpzyd"
        elif filename.find("宝贝魔术师美容专营店") >= 0:
            df1["SHOPNAME"] = "宝贝魔术师美容专营店"
            df1["SHOPCODE"] = "bbmssmrzyd"
        elif filename.find("不酷美妆专营店") >= 0:
            df1["SHOPNAME"] = "不酷美妆专营店"
            df1["SHOPCODE"] = "bkmzzyd"
        elif filename.find("橙意满满家居专营店") >= 0:
            df1["SHOPNAME"] = "橙意满满家居专营店"
            df1["SHOPCODE"] = "cymmjjzyd"
        elif filename.find("橙意满满家清专营店") >= 0:
            df1["SHOPNAME"] = "橙意满满家清专营店"
            df1["SHOPCODE"] = "cymmjqzyd"
        elif filename.find("橙意满满日用品专营店") >= 0:
            df1["SHOPNAME"] = "橙意满满日用品专营店"
            df1["SHOPCODE"] = "cymmrypzyd"
        elif filename.find("无极爽化妆品专营店") >= 0:
            df1["SHOPNAME"] = "无极爽化妆品专营店"
            df1["SHOPCODE"] = "wjshzpzyd"
        elif filename.find("不酷家居生活专营店") >= 0:
            df1["SHOPNAME"] = "不酷家居生活专营店"
            df1["SHOPCODE"] = "bkjjshzyd"
        elif filename.find("茱莉珂丝家居专营店") >= 0:
            df1["SHOPNAME"] = "茱莉珂丝家居专营店"
            df1["SHOPCODE"] = "zjjzyd"
        elif filename.find("谷口家居生活专营店") >= 0:
            df1["SHOPNAME"] = "谷口家居生活专营店"
            df1["SHOPCODE"] = "gkjjshzyd"
        elif filename.find("茱莉珂丝居家清洁专营店") >= 0:
            df1["SHOPNAME"] = "茱莉珂丝居家清洁专营店"
            df1["SHOPCODE"] = "zjjqjzyd"
        elif filename.find("茱莉珂丝家居生活专营店") >= 0:
            df1["SHOPNAME"] = "茱莉珂丝家居生活专营店"
            df1["SHOPCODE"] = "zjjshzyd"
        elif filename.find("樱语家居生活专营店") >= 0:
            df1["SHOPNAME"] = "樱语家居生活专营店"
            df1["SHOPCODE"] = "yyjjshzyd"
        elif filename.find("樱语日用品专营店") >= 0:
            df1["SHOPNAME"] = "樱语日用品专营店"
            df1["SHOPCODE"] = "yyrypzyd"
        elif filename.find("橙意满满居家日用专营店") >= 0:
            df1["SHOPNAME"] = "橙意满满居家日用专营店"
            df1["SHOPCODE"] = "cymmjjryzyd"
        elif filename.find("橙意满满家居生活专营店") >= 0:
            df1["SHOPNAME"] = "橙意满满家居生活专营店"
            df1["SHOPCODE"] = "cymmjjshzyd"
        elif filename.find("谷口清洁用品专营店") >= 0:
            df1["SHOPNAME"] = "谷口清洁用品专营店"
            df1["SHOPCODE"] = "gkqjypzyd"
        elif filename.find("樱语洗护专营店") >= 0:
            df1["SHOPNAME"] = "樱语洗护专营店"
            df1["SHOPCODE"] = "yyxhzyd"

        if filename.find("海外") >= 0:
            df1["currency"] = "USD"
            df1["overseas_income"] = df["收入金额（+$）"]
            df1["overseas_expend"] = df["支出金额（-$）"]
            df1["currency_cny_rate"] = df["汇率（美元兑人民币）"]
            df1["overseas_income"] = df1["overseas_income"].astype(float).abs()
            df1["overseas_expend"] = -df1["overseas_expend"].astype(float).abs()
        else:
            df1["currency"] = ""
            df1["overseas_income"] = ""
            df1["overseas_expend"] = ""
            df1["currency_cny_rate"] = ""
        print(df1.head(5).to_markdown())
        print("拼多多")

    # 抖音逻辑
    elif filename.find("抖音") >= 0:
        if filename.find("保证金") >= 0:
            if filename.find("xls") >= 0:
                df = pd.read_excel(filename, dtype=str)
            else:
                try:
                    df = pd.read_csv(filename, dtype=str)
                except Exception as e:
                    df = pd.read_csv(filename, dtype=str, encoding="gb18030")

            df = df[df["操作类型"].str.contains("售后扣款")]
            if len(df) > 0:
                for column_name in df.columns:
                    df.rename(columns={column_name: column_name.replace(" ", "").replace("\n", "").strip()},
                              inplace=True)
                print(df.head(1).to_markdown())
                df["备注"] = df["备注"].astype(str)
                df["订单号"] = df["备注"].apply(lambda x: "".join(("".join(x.split("：")[1:])).split("订单")[:-1]))
                print(df.head(1).to_markdown())

                df["操作金额"] = df["操作金额"].astype(float)

                plat = "DY"

                # 保证金
                df1 = pd.DataFrame()
                df1["TID"] = df["订单号"]
                df1["SHOPNAME"] = ""
                df1["PLATFORM"] = plat
                df1["SHOPCODE"] = ""
                df1["BILLPLATFORM"] = plat
                df1["CREATED"] = df["操作时间"]
                df1["TITLE"] = ""
                df1["TRADE_TYPE"] = df["操作类型"]
                df1["BUSINESS_NO"] = df["操作单号"]
                df1["INCOME_AMOUNT"] = 0
                df1["EXPEND_AMOUNT"] = df["操作金额"]
                df1["TRADING_CHANNELS"] = ""
                df1["BUSINESS_DESCRIPTION"] = "操作单号：" + df["操作单号"] + "。交易状态：" + df["交易状态"]
                df1["remark"] = df["备注"].apply(lambda x: x[x.find("发生") - 1:])
                df1["BUSINESS_BILL_SOURCE"] = ""
                df1["IS_REFUNDAMOUNT"] = 1
                df1["IS_AMOUNT"] = 0
                df1["OID"] = ""
                df1["SOURCEDATA"] = "EXCEL"
                df1["RECIPROCAL_ACCOUNT"] = ""
                df1["BATCHNO"] = ""
                df1["currency"] = ""
                df1["overseas_income"] = ""
                df1["overseas_expend"] = ""
                df1["currency_cny_rate"] = ""
                df1["EXPEND_AMOUNT"] = -df1["EXPEND_AMOUNT"].astype(float).abs()
                # df1 = df1.loc[df1.INCOME_AMOUNT != 0]
                print(df1.head(5).to_markdown())
            else:
                print("无售后扣款数据！")
                dict = {"TID": "nan", "SHOPNAME": "", "PLATFORM": "", "SHOPCODE": "", "BILLPLATFORM": "", "CREATED": "",
                        "TITLE": "", "TRADE_TYPE": "", "BUSINESS_NO": "", "INCOME_AMOUNT": "", "EXPEND_AMOUNT": "",
                        "TRADING_CHANNELS": "", "BUSINESS_DESCRIPTION": "", "remark": "", "BUSINESS_BILL_SOURCE": "",
                        "IS_REFUNDAMOUNT": "", "IS_AMOUNT": "", "OID": "", "SOURCEDATA": "", "RECIPROCAL_ACCOUNT": "",
                        "BATCHNO": ""}
                df1 = pd.DataFrame(dict, index=[0])
                return df1

        else:
            try:
                df = pd.read_excel(filename, sheet_name="月账单", dtype=str)
            except Exception as e:
                try:
                    df = pd.read_excel(filename, sheet_name="在线支付订单账单", dtype=str)
                except Exception as e:
                    try:
                        df = pd.read_excel(filename, dtype=str)
                    except Exception as e:
                        try:
                            df = pd.read_csv(filename, dtype=str, encoding="gb18030")
                        except Exception as e:
                            df = pd.read_csv(filename, dtype=str)

            for column_name in df.columns:
                df.rename(columns={column_name: column_name.replace(" ", "").replace("\n", "").strip()}, inplace=True)

            df["订单号"] = df["订单号"].astype(str)
            df["平台服务费(元)"] = df["平台服务费(元)"].astype(float)
            df["订单号"] = df["订单号"].apply(
                lambda x: x.replace("¥", "").replace(":", "").replace("'", "").replace('"', ''))
            if "子订单号" in df.columns:
                df["子订单号"] = df["子订单号"].apply(
                    lambda x: x.replace("¥", "").replace(":", "").replace("'", "").replace('"', ''))
                df["订单号"] = df.apply(lambda x: x["订单号"] if len(x["订单号"]) > 3 else x["子订单号"], axis=1)
            df.dropna(subset=["订单号"], inplace=True)
            df = df[~df["订单号"].str.contains("nan")]
            print(df.head(5).to_markdown())
            plat = "DY"

            # 实际支付
            df1 = pd.DataFrame()
            df1["TID"] = df["订单号"]
            df1["SHOPNAME"] = ""
            df1["PLATFORM"] = plat
            df1["SHOPCODE"] = ""
            df1["BILLPLATFORM"] = plat
            df1["CREATED"] = df["结算时间"]
            df1["TITLE"] = ""
            df1["TRADE_TYPE"] = df["业务类型"]
            df1["BUSINESS_NO"] = ""
            df1["INCOME_AMOUNT"] = df["订单实付(元)"]
            df1["EXPEND_AMOUNT"] = 0
            df1["TRADING_CHANNELS"] = df["结算账户"]
            df1["BUSINESS_DESCRIPTION"] = "实际支付"
            df1["remark"] = ""
            df1["BUSINESS_BILL_SOURCE"] = ""
            df1["IS_REFUNDAMOUNT"] = 0
            df1["IS_AMOUNT"] = 1
            df1["OID"] = ""
            df1["SOURCEDATA"] = "EXCEL"
            df1["RECIPROCAL_ACCOUNT"] = ""
            df1["BATCHNO"] = ""
            df1["currency"] = ""
            df1["overseas_income"] = ""
            df1["overseas_expend"] = ""
            df1["currency_cny_rate"] = ""
            df1["INCOME_AMOUNT"] = df1["INCOME_AMOUNT"].astype(float).abs()
            # df1 = df1.loc[df1.INCOME_AMOUNT != 0]
            print(df1.head(5).to_markdown())

            # 平台服务费
            df2 = df1.copy()
            df2["INCOME_AMOUNT"] = df.apply(lambda x: x["平台服务费(元)"] if x["平台服务费(元)"] > 0 else 0, axis=1)
            df2["EXPEND_AMOUNT"] = df.apply(lambda x: x["平台服务费(元)"] if x["平台服务费(元)"] < 0 else 0, axis=1)
            df2["BUSINESS_DESCRIPTION"] = "平台服务费"
            df2["IS_REFUNDAMOUNT"] = 0
            df2["IS_AMOUNT"] = 0
            df2["EXPEND_AMOUNT"] = df2["EXPEND_AMOUNT"].astype(float)
            # df2 = df2.loc[df2.EXPEND_AMOUNT !=0]
            print(df2.head(5).to_markdown())

            # 达人佣金
            df3 = df1.copy()
            df3["INCOME_AMOUNT"] = 0
            if "佣金(元)" in df.columns:
                df3["EXPEND_AMOUNT"] = df["佣金(元)"]
            elif "达人佣金(元)" in df.columns:
                df3["EXPEND_AMOUNT"] = df["达人佣金(元)"]
            else:
                df3["EXPEND_AMOUNT"] = 0
            df3["BUSINESS_DESCRIPTION"] = "达人佣金"
            df3["IS_REFUNDAMOUNT"] = 0
            df3["IS_AMOUNT"] = 0
            df3["EXPEND_AMOUNT"] = df3["EXPEND_AMOUNT"].astype(float)
            # df3 = df3.loc[df3.EXPEND_AMOUNT != 0]
            print(df3.head(5).to_markdown())

            # 平台补贴
            # df4 = df1.copy()
            # if "实际补贴金额(元)" in df.columns:
            #     df["实际补贴金额(元)"] = df["实际补贴金额(元)"].astype(float)
            #     df4["INCOME_AMOUNT"] = df["实际补贴金额(元)"].apply(lambda x: x if x > 0 else 0)
            #     df4["EXPEND_AMOUNT"] = df["实际补贴金额(元)"].apply(lambda x: x if x < 0 else 0)
            # elif "实际平台补贴(元)" in df.columns:
            #     df["实际平台补贴(元)"] = df["实际平台补贴(元)"].astype(float)
            #     df4["INCOME_AMOUNT"] = df["实际平台补贴(元)"].apply(lambda x: x if x > 0 else 0)
            #     df4["EXPEND_AMOUNT"] = df["实际平台补贴(元)"].apply(lambda x: x if x < 0 else 0)
            # elif "平台补贴(元)" in df.columns:
            #     df["平台补贴(元)"] = df["平台补贴(元)"].astype(float)
            #     df4["INCOME_AMOUNT"] = df["平台补贴(元)"].apply(lambda x: x if x > 0 else 0)
            #     df4["EXPEND_AMOUNT"] = df["平台补贴(元)"].apply(lambda x: x if x < 0 else 0)
            # else:
            #     df4["INCOME_AMOUNT"] = 0
            #     df4["EXPEND_AMOUNT"] = 0
            # df4["BUSINESS_DESCRIPTION"] = "平台补贴"
            # df4["IS_REFUNDAMOUNT"] = df.apply(lambda x:1 if x["结算状态"].find("退款")>=0 else 0,axis=1)
            # df4["IS_AMOUNT"] = df.apply(lambda x:1 if x["结算状态"]=="已结算" else 0,axis=1)
            # df4["INCOME_AMOUNT"] = df4["INCOME_AMOUNT"].astype(float).abs()
            # # df4 = df4.loc[df4.INCOME_AMOUNT != 0]
            # print(df4.head(5).to_markdown())

            # 平台补贴
            df4 = df1.copy()
            if (("实际平台补贴(元)" in df.columns) & ("平台补贴(元)" in df.columns)):
                df["实际平台补贴(元)"] = df["实际平台补贴(元)"].astype(float)
                df["平台补贴(元)"] = df["平台补贴(元)"].astype(float)
                df4["INCOME_AMOUNT1"] = df.apply(lambda x: x["实际平台补贴(元)"] if x["实际平台补贴(元)"] > 0 else x["平台补贴(元)"],
                                                 axis=1)
                df4["EXPEND_AMOUNT1"] = df.apply(lambda x: x["实际平台补贴(元)"] if x["实际平台补贴(元)"] < 0 else x["平台补贴(元)"],
                                                 axis=1)
            elif "实际平台补贴(元)" in df.columns:
                df["实际平台补贴(元)"] = df["实际平台补贴(元)"].astype(float)
                df4["INCOME_AMOUNT1"] = df.apply(lambda x: x["实际平台补贴(元)"] if x["实际平台补贴(元)"] > 0 else 0, axis=1)
                df4["EXPEND_AMOUNT1"] = df.apply(lambda x: x["实际平台补贴(元)"] if x["实际平台补贴(元)"] < 0 else 0, axis=1)
            if "平台补贴(元)" in df.columns:
                df["平台补贴(元)"] = df["平台补贴(元)"].astype(float)
                df4["INCOME_AMOUNT1"] = df.apply(lambda x: x["平台补贴(元)"] if x["平台补贴(元)"] > 0 else 0, axis=1)
                df4["EXPEND_AMOUNT1"] = df.apply(lambda x: x["平台补贴(元)"] if x["平台补贴(元)"] < 0 else 0, axis=1)
            else:
                df4["INCOME_AMOUNT1"] = 0
                df4["EXPEND_AMOUNT1"] = 0

            if (("实际达人补贴(元)" in df.columns) & ("达人补贴(元)" in df.columns)):
                df["实际达人补贴(元)"] = df["实际达人补贴(元)"].astype(float)
                df["达人补贴(元)"] = df["达人补贴(元)"].astype(float)
                df4["INCOME_AMOUNT2"] = df.apply(lambda x: x["实际达人补贴(元)"] if x["实际达人补贴(元)"] > 0 else x["达人补贴(元)"],
                                                 axis=1)
                df4["EXPEND_AMOUNT2"] = df.apply(lambda x: x["实际达人补贴(元)"] if x["实际达人补贴(元)"] < 0 else x["达人补贴(元)"],
                                                 axis=1)
            elif "实际达人补贴(元)" in df.columns:
                df["实际达人补贴(元)"] = df["实际达人补贴(元)"].astype(float)
                df4["INCOME_AMOUNT2"] = df.apply(lambda x: x["实际达人补贴(元)"] if x["实际达人补贴(元)"] > 0 else 0, axis=1)
                df4["EXPEND_AMOUNT2"] = df.apply(lambda x: x["实际达人补贴(元)"] if x["实际达人补贴(元)"] < 0 else 0, axis=1)
            if "达人补贴(元)" in df.columns:
                df["达人补贴(元)"] = df["达人补贴(元)"].astype(float)
                df4["INCOME_AMOUNT2"] = df.apply(lambda x: x["达人补贴(元)"] if x["达人补贴(元)"] > 0 else 0, axis=1)
                df4["EXPEND_AMOUNT2"] = df.apply(lambda x: x["达人补贴(元)"] if x["达人补贴(元)"] < 0 else 0, axis=1)
            else:
                df4["INCOME_AMOUNT2"] = 0
                df4["EXPEND_AMOUNT2"] = 0

            if (("实际抖音支付补贴(元)" in df.columns) & ("抖音支付补贴(元)" in df.columns)):
                df["实际抖音支付补贴(元)"] = df["实际抖音支付补贴(元)"].astype(float)
                df["抖音支付补贴(元)"] = df["抖音支付补贴(元)"].astype(float)
                df4["INCOME_AMOUNT3"] = df.apply(lambda x: x["实际抖音支付补贴(元)"] if x["实际抖音支付补贴(元)"] > 0 else x["抖音支付补贴(元)"],
                                                 axis=1)
                df4["EXPEND_AMOUNT3"] = df.apply(lambda x: x["实际抖音支付补贴(元)"] if x["实际抖音支付补贴(元)"] < 0 else x["抖音支付补贴(元)"],
                                                 axis=1)
            elif "实际抖音支付补贴(元)" in df.columns:
                df["实际抖音支付补贴(元)"] = df["实际抖音支付补贴(元)"].astype(float)
                df4["INCOME_AMOUNT3"] = df.apply(lambda x: x["实际抖音支付补贴(元)"] if x["实际抖音支付补贴(元)"] > 0 else 0, axis=1)
                df4["EXPEND_AMOUNT3"] = df.apply(lambda x: x["实际抖音支付补贴(元)"] if x["实际抖音支付补贴(元)"] < 0 else 0, axis=1)
            if "抖音支付补贴(元)" in df.columns:
                df["抖音支付补贴(元)"] = df["抖音支付补贴(元)"].astype(float)
                df4["INCOME_AMOUNT3"] = df.apply(lambda x: x["抖音支付补贴(元)"] if x["抖音支付补贴(元)"] > 0 else 0, axis=1)
                df4["EXPEND_AMOUNT3"] = df.apply(lambda x: x["抖音支付补贴(元)"] if x["抖音支付补贴(元)"] < 0 else 0, axis=1)
            else:
                df4["INCOME_AMOUNT3"] = 0
                df4["EXPEND_AMOUNT3"] = 0

            if (("实际DOU分期营销补贴(元)" in df.columns) & ("DOU分期营销补贴(元)" in df.columns)):
                df["实际DOU分期营销补贴(元)"] = df["实际DOU分期营销补贴(元)"].astype(float)
                df["DOU分期营销补贴(元)"] = df["DOU分期营销补贴(元)"].astype(float)
                df4["INCOME_AMOUNT4"] = df.apply(
                    lambda x: x["实际DOU分期营销补贴(元)"] if x["实际DOU分期营销补贴(元)"] > 0 else x["DOU分期营销补贴(元)"], axis=1)
                df4["EXPEND_AMOUNT4"] = df.apply(
                    lambda x: x["实际DOU分期营销补贴(元)"] if x["实际DOU分期营销补贴(元)"] < 0 else x["DOU分期营销补贴(元)"], axis=1)
            elif "实际DOU分期营销补贴(元)" in df.columns:
                df["实际DOU分期营销补贴(元)"] = df["实际DOU分期营销补贴(元)"].astype(float)
                df4["INCOME_AMOUNT4"] = df.apply(lambda x: x["实际DOU分期营销补贴(元)"] if x["实际DOU分期营销补贴(元)"] > 0 else 0,
                                                 axis=1)
                df4["EXPEND_AMOUNT4"] = df.apply(lambda x: x["实际DOU分期营销补贴(元)"] if x["实际DOU分期营销补贴(元)"] < 0 else 0,
                                                 axis=1)
            if "DOU分期营销补贴(元)" in df.columns:
                df["DOU分期营销补贴(元)"] = df["DOU分期营销补贴(元)"].astype(float)
                df4["INCOME_AMOUNT4"] = df.apply(lambda x: x["DOU分期营销补贴(元)"] if x["DOU分期营销补贴(元)"] > 0 else 0, axis=1)
                df4["EXPEND_AMOUNT4"] = df.apply(lambda x: x["DOU分期营销补贴(元)"] if x["DOU分期营销补贴(元)"] < 0 else 0, axis=1)
            else:
                df4["INCOME_AMOUNT4"] = 0
                df4["EXPEND_AMOUNT4"] = 0

            df4["INCOME_AMOUNT"] = df4["INCOME_AMOUNT1"] + df4["INCOME_AMOUNT2"] + df4["INCOME_AMOUNT3"] + df4[
                "INCOME_AMOUNT4"]
            df4["EXPEND_AMOUNT"] = df4["EXPEND_AMOUNT1"] + df4["EXPEND_AMOUNT2"] + df4["EXPEND_AMOUNT3"] + df4[
                "EXPEND_AMOUNT4"]
            df4["INCOME_AMOUNT"] = df4["INCOME_AMOUNT"].astype(float).abs()
            df4["EXPEND_AMOUNT"] = -df4["EXPEND_AMOUNT"].astype(float).abs()
            df4["BUSINESS_DESCRIPTION"] = "平台补贴"
            df4["IS_REFUNDAMOUNT"] = df4.apply(lambda x: 1 if x["EXPEND_AMOUNT"] < 0 else 0, axis=1)
            df4["IS_AMOUNT"] = df4.apply(lambda x: 1 if x["INCOME_AMOUNT"] > 0 else 0, axis=1)

            del df4["INCOME_AMOUNT1"]
            del df4["EXPEND_AMOUNT1"]
            del df4["INCOME_AMOUNT2"]
            del df4["EXPEND_AMOUNT2"]
            del df4["INCOME_AMOUNT3"]
            del df4["EXPEND_AMOUNT3"]
            del df4["INCOME_AMOUNT4"]
            del df4["EXPEND_AMOUNT4"]
            # df4 = df4.loc[df4.INCOME_AMOUNT != 0]
            print(df4.head(5).to_markdown())

            # 支付优惠
            try:
                df5 = df1.copy()
                if "实际支付优惠(元)" in df.columns:
                    df["实际支付优惠(元)"] = df["实际支付优惠(元)"].astype(float)
                    df5["INCOME_AMOUNT"] = df["实际支付优惠(元)"].apply(lambda x: x if x > 0 else 0)
                    df5["EXPEND_AMOUNT"] = df["实际支付优惠(元)"].apply(lambda x: x if x < 0 else 0)
                elif "支付优惠(元)" in df.columns:
                    df["支付优惠(元)"] = df["支付优惠(元)"].astype(float)
                    df5["INCOME_AMOUNT"] = df["支付优惠(元)"].apply(lambda x: x if x > 0 else 0)
                    df5["EXPEND_AMOUNT"] = df["支付优惠(元)"].apply(lambda x: x if x < 0 else 0)
                else:
                    df5["INCOME_AMOUNT"] = 0
                    df5["EXPEND_AMOUNT"] = 0
                df5["BUSINESS_DESCRIPTION"] = "支付优惠"
                # df5["IS_REFUNDAMOUNT"] = df.apply(lambda x:1 if x["结算状态"].find("退款")>=0 else 0,axis=1)
                # df5["IS_AMOUNT"] = df.apply(lambda x:1 if x["结算状态"]=="已结算" else 0,axis=1)
                df5["INCOME_AMOUNT"] = df5["INCOME_AMOUNT"].astype(float).abs()
                df5["EXPEND_AMOUNT"] = -df5["EXPEND_AMOUNT"].astype(float).abs()
                df5["IS_REFUNDAMOUNT"] = df5.apply(lambda x: 1 if x["EXPEND_AMOUNT"] < 0 else 0, axis=1)
                df5["IS_AMOUNT"] = df5.apply(lambda x: 1 if x["INCOME_AMOUNT"] > 0 else 0, axis=1)
                # df5 = df5.loc[df5.INCOME_AMOUNT != 0]
                print(df5.head(5).to_markdown())
            except Exception as e:
                print("无支付优惠相关字段")
                dict = {"TID": "nan", "SHOPNAME": "", "PLATFORM": "", "SHOPCODE": "", "BILLPLATFORM": "", "CREATED": "",
                        "TITLE": "", "TRADE_TYPE": "", "BUSINESS_NO": "", "INCOME_AMOUNT": "", "EXPEND_AMOUNT": "",
                        "TRADING_CHANNELS": "", "BUSINESS_DESCRIPTION": "", "remark": "", "BUSINESS_BILL_SOURCE": "",
                        "IS_REFUNDAMOUNT": "", "IS_AMOUNT": "", "OID": "", "SOURCEDATA": "", "RECIPROCAL_ACCOUNT": "",
                        "BATCHNO": ""}
                df5 = pd.DataFrame(dict, index=[0])

            # 订单退款
            try:
                df6 = df1.copy()
                df6["INCOME_AMOUNT"] = 0
                df6["EXPEND_AMOUNT"] = df["订单退款(元)"]
                df6["BUSINESS_DESCRIPTION"] = "订单退款"
                df6["IS_REFUNDAMOUNT"] = 1
                df6["IS_AMOUNT"] = 0
                df6["EXPEND_AMOUNT"] = -df6["EXPEND_AMOUNT"].astype(float).abs()
                # df6 = df6.loc[df6.EXPEND_AMOUNT !=0]
                print(df6.head(5).to_markdown())

            except Exception as e:
                print("无订单退款相关字段")
                dict = {"TID": "nan", "SHOPNAME": "", "PLATFORM": "", "SHOPCODE": "", "BILLPLATFORM": "", "CREATED": "",
                        "TITLE": "", "TRADE_TYPE": "", "BUSINESS_NO": "", "INCOME_AMOUNT": "", "EXPEND_AMOUNT": "",
                        "TRADING_CHANNELS": "", "BUSINESS_DESCRIPTION": "", "remark": "", "BUSINESS_BILL_SOURCE": "",
                        "IS_REFUNDAMOUNT": "", "IS_AMOUNT": "", "OID": "", "SOURCEDATA": "", "RECIPROCAL_ACCOUNT": "",
                        "BATCHNO": ""}
                df6 = pd.DataFrame(dict, index=[0])

            # 结算前退款
            try:
                df7 = df1.copy()
                df7["INCOME_AMOUNT"] = 0
                df7["EXPEND_AMOUNT"] = df["结算前退款(元)"]
                df7["BUSINESS_DESCRIPTION"] = "订单退款"
                df7["IS_REFUNDAMOUNT"] = 1
                df7["IS_AMOUNT"] = 0
                df7["EXPEND_AMOUNT"] = -df7["EXPEND_AMOUNT"].astype(float).abs()
                # df7 = df7.loc[df7.EXPEND_AMOUNT !=0]
                print(df7.head(5).to_markdown())

            except Exception as e:
                print("无结算前退款相关字段")
                dict = {"TID": "nan", "SHOPNAME": "", "PLATFORM": "", "SHOPCODE": "", "BILLPLATFORM": "", "CREATED": "",
                        "TITLE": "", "TRADE_TYPE": "", "BUSINESS_NO": "", "INCOME_AMOUNT": "", "EXPEND_AMOUNT": "",
                        "TRADING_CHANNELS": "", "BUSINESS_DESCRIPTION": "", "remark": "", "BUSINESS_BILL_SOURCE": "",
                        "IS_REFUNDAMOUNT": "", "IS_AMOUNT": "", "OID": "", "SOURCEDATA": "", "RECIPROCAL_ACCOUNT": "",
                        "BATCHNO": ""}
                df7 = pd.DataFrame(dict, index=[0])

            dfs = [df1, df2, df3, df4, df5, df6, df7]
            df1 = pd.concat(dfs)

        df1 = df1[~df1["TID"].str.contains("nan")]
        df1 = df1.loc[~((df1.INCOME_AMOUNT == 0) & (df1.EXPEND_AMOUNT == 0))]
        df1["INCOME_AMOUNT"] = df1["INCOME_AMOUNT"].astype(float)
        df1["INCOME_AMOUNT"] = df1["INCOME_AMOUNT"].abs()
        df1["EXPEND_AMOUNT"] = -df1["EXPEND_AMOUNT"].astype(float)
        df1["EXPEND_AMOUNT"] = -df1["EXPEND_AMOUNT"].abs()
        if ((filename.find("航星个护") >= 0) or (filename.find("mega店铺") >= 0)):
            df1["SHOPCODE"] = "hxzmd"
            df1["SHOPNAME"] = "航星个护专营店"
        elif filename.find("可瘾个护") >= 0:
            df1["SHOPCODE"] = "kyjqzyd"
            df1["SHOPNAME"] = "可瘾个护家清专营店"
        elif filename.find("麦凯莱个护") >= 0:
            df1["SHOPCODE"] = "mklghjqzyd"
            df1["SHOPNAME"] = "麦凯莱个护家清专营店"
        elif filename.find("麦凯莱美妆") >= 0:
            df1["SHOPCODE"] = "mklmzzyd"
            df1["SHOPNAME"] = "麦凯莱美妆专营店"
        elif filename.find("播地艾个护专营店") >= 0:
            df1["SHOPCODE"] = "bdaghzyd"
            df1["SHOPNAME"] = "播地艾个护专营店"
        elif filename.find("尚西专营") >= 0:
            df1["SHOPCODE"] = "sxzyd"
            df1["SHOPNAME"] = "尚西专营店"
        elif filename.find("Dentyl") >= 0:
            df1["SHOPCODE"] = "dentylactive"
            df1["SHOPNAME"] = "DentylActive官方旗舰店"
        elif ((filename.find("肯妮诗") >= 0) & (filename.find("loshi") < 0)):
            df1["SHOPCODE"] = "knsghzyd"
            df1["SHOPNAME"] = "肯妮诗个护专营店"
        elif filename.find("博滴卖家") >= 0:
            df1["SHOPCODE"] = "madesmjyxsyzmd"
            df1["SHOPNAME"] = "博滴卖家优选专卖店"
        elif filename.find("博滴旗舰店") >= 0:
            df1["SHOPCODE"] = "jlsmptd"
            df1["SHOPNAME"] = "博滴旗舰店"
        elif ((filename.find("BodyAid个护") >= 0) or (filename.find("bodyaid个护") >= 0)):
            df1["SHOPCODE"] = "nclqjd"
            df1["SHOPNAME"] = "BodyAid个护旗舰店"
        elif filename.find("睿旗") >= 0:
            df1["SHOPCODE"] = "rqghjqzyd"
            df1["SHOPNAME"] = "睿旗个护家清专营店"
        elif filename.find("loshi肯妮诗") >= 0:
            df1["SHOPCODE"] = "loshikns"
            df1["SHOPNAME"] = "loshi肯妮诗专卖店"
        elif filename.find("loshi旗舰店") >= 0:
            df1["SHOPCODE"] = "loshi"
            df1["SHOPNAME"] = "loshi旗舰店"
        elif filename.find("BodyAid旗舰店") >= 0:
            df1["SHOPCODE"] = "bodyaid"
            df1["SHOPNAME"] = "bodyaid旗舰店"
        elif filename.find("MOREI家清") >= 0:
            df1["SHOPCODE"] = "moreijq"
            df1["SHOPNAME"] = "morei家清旗舰店"
        elif filename.find("MOREI日用品") >= 0:
            df1["SHOPCODE"] = "moreirypqjd"
            df1["SHOPNAME"] = "MOREI日用品旗舰店"
        elif filename.find("芭葆兔") >= 0:
            df1["SHOPCODE"] = "bbtrypzyd"
            df1["SHOPNAME"] = "芭葆兔日用品专营店"
        elif filename.find("白卿") >= 0:
            df1["SHOPCODE"] = "bqrypzyd"
            df1["SHOPNAME"] = "白卿日用专营店"
        elif filename.find("宝贝魔术师") >= 0:
            df1["SHOPCODE"] = "bbmssgrhlzyd"
            df1["SHOPNAME"] = "宝贝魔术师个人护理专营店"
        elif filename.find("宝贝配方师") >= 0:
            df1["SHOPCODE"] = "bbpfsrypzyd"
            df1["SHOPNAME"] = "宝贝配方师日用品专营店"
        elif filename.find("贝贝港湾") >= 0:
            df1["SHOPCODE"] = "bbwgmzzyd"
            df1["SHOPNAME"] = "贝贝港湾美妆专营店"
        elif filename.find("斌闻") >= 0:
            df1["SHOPCODE"] = "bwmzzyd"
            df1["SHOPNAME"] = "斌闻美妆专营店"
        elif filename.find("冰川女神") >= 0:
            df1["SHOPCODE"] = "bcnsrhypzyd"
            df1["SHOPNAME"] = "冰川女神日化用品专营店"
        elif filename.find("不酷") >= 0:
            df1["SHOPCODE"] = "bkgrhlzyd"
            df1["SHOPNAME"] = "不酷个人护理专营店"
        elif filename.find("不是很酷") >= 0:
            df1["SHOPCODE"] = "bshkfzmyzyd"
            df1["SHOPNAME"] = "不是很酷服装贸易专营店"
        elif filename.find("航星玩具") >= 0:
            df1["SHOPCODE"] = "hxwjzyd"
            df1["SHOPNAME"] = "航星玩具专营店"
        elif filename.find("MONTOOTH旗舰店") >= 0:
            df1["SHOPCODE"] = "nuggelasulexmyzmd"
            df1["SHOPNAME"] = "MONTOOTH旗舰店"
        elif filename.find("Montooth萌洁齿旗舰店") >= 0:
            df1["SHOPCODE"] = "montoothmjcqjd"
            df1["SHOPNAME"] = "Montooth萌洁齿旗舰店"
        elif filename.find("卖家优选个人") >= 0:
            df1["SHOPCODE"] = "mjyxgrhlzyd"
            df1["SHOPNAME"] = "卖家优选个人护理专营店"
        elif filename.find("魔湾个人") >= 0:
            df1["SHOPCODE"] = "mwgrhlzyd"
            df1["SHOPNAME"] = "魔湾个人护理专营店"
        elif filename.find("配颜个护") >= 0:
            df1["SHOPCODE"] = "pyghzyd"
            df1["SHOPNAME"] = "配颜个护专营店"
        elif filename.find("配颜师个护") >= 0:
            df1["SHOPCODE"] = "pysghzyd"
            df1["SHOPNAME"] = "配颜师个护专营店"
        elif filename.find("配颜师日化品") >= 0:
            df1["SHOPCODE"] = "pysrhpzyd"
            df1["SHOPNAME"] = "配颜师日化品专营店"
        elif filename.find("尚隐日用") >= 0:
            df1["SHOPCODE"] = "syrpzyd"
            df1["SHOPNAME"] = "尚隐日用专营店"
        elif filename.find("樱语个护") >= 0:
            df1["SHOPCODE"] = "yyghqjd"
            df1["SHOPNAME"] = "樱语个护旗舰店"
        elif filename.find("宅星人日用品") >= 0:
            df1["SHOPCODE"] = "zxrrypzyd"
            df1["SHOPNAME"] = "宅星人日用品专营店"
        elif filename.find("珍芯漾肤") >= 0:
            df1["SHOPCODE"] = "zxyfgrhld"
            df1["SHOPNAME"] = "珍芯漾肤个人护理专营店"
        elif filename.find("朱莉珂丝") >= 0:
            df1["SHOPCODE"] = "zlksrhypzyd"
            df1["SHOPNAME"] = "朱莉珂丝日化用品专营店"
        elif filename.find("若蘅旗舰店") >= 0:
            df1["SHOPCODE"] = "rhqjd"
            df1["SHOPNAME"] = "若蘅旗舰店"
        elif filename.find("mega超值精选") >= 0:
            df1["SHOPCODE"] = "megaczjx"
            df1["SHOPNAME"] = "mega超值精选"
        elif filename.find("mega小店") >= 0:
            df1["SHOPCODE"] = "megaxd"
            df1["SHOPNAME"] = "魔湾游戏个护专营店"
        elif filename.find("spr小店") >= 0:
            df1["SHOPCODE"] = "sprxd"
            df1["SHOPNAME"] = "spr小店"
        elif filename.find("博滴小店") >= 0:
            df1["SHOPCODE"] = "bdxd"
            df1["SHOPNAME"] = "博滴小店"
        elif filename.find("邓特个护") >= 0:
            df1["SHOPCODE"] = "dentylactivegh"
            df1["SHOPNAME"] = "邓特个护"
        elif filename.find("精粮商贸") >= 0:
            df1["SHOPCODE"] = "jlsmptd"
            df1["SHOPNAME"] = "博滴旗舰店"
        elif filename.find("乐丝小铺") >= 0:
            df1["SHOPCODE"] = "lsxp"
            df1["SHOPNAME"] = "乐丝小铺"
        elif filename.find("玫德丝配颜师") >= 0:
            df1["SHOPCODE"] = "madespyszmd"
            df1["SHOPNAME"] = "植之璨配颜师专卖店"
        elif filename.find("魔湾小店") >= 0:
            df1["SHOPCODE"] = "mwxd"
            df1["SHOPNAME"] = "魔湾个护家清专营店"
        elif filename.find("奈萃拉多瑞") >= 0:
            df1["SHOPCODE"] = "ncldrzmd"
            df1["SHOPNAME"] = "来一泡多瑞专卖店"
        elif filename.find("奈萃拉旗舰店") >= 0:
            df1["SHOPCODE"] = "nclqjd"
            df1["SHOPNAME"] = "BodyAid个护旗舰店"
        elif filename.find("奈萃拉樱岚") >= 0:
            df1["SHOPCODE"] = "nclyfzmd"
            df1["SHOPNAME"] = "AllNaturalAdvice樱岚专卖店"
        elif filename.find("纽苏美丽") >= 0:
            df1["SHOPCODE"] = "nsmlxd"
            df1["SHOPNAME"] = "纽苏美丽小店"
        elif filename.find("微笑铺子") >= 0:
            df1["SHOPCODE"] = "wxpz"
            df1["SHOPNAME"] = "微笑铺子"
        elif filename.find("羊羊的小铺") >= 0:
            df1["SHOPCODE"] = "essxsqtghzyd"
            df1["SHOPNAME"] = "二十四小时七天个护专营店"
        elif filename.find("盈养泉旗舰店") >= 0:
            df1["SHOPCODE"] = "yyqzmd"
            df1["SHOPNAME"] = "盈养泉旗舰店"
        elif filename.find("悠尼珂丝") >= 0:
            df1["SHOPCODE"] = "ynksxd"
            df1["SHOPNAME"] = "悠尼珂丝小店"
        elif filename.find("mega精选小店") >= 0:
            df1["SHOPCODE"] = "megajxxd"
            df1["SHOPNAME"] = "Mega精选小铺"
        elif filename.find("mades个护") >= 0:
            df1["SHOPCODE"] = "madesghqjd"
            df1["SHOPNAME"] = "博滴个护旗舰店"
        elif filename.find("EC萌洁齿专卖店") >= 0:
            df1["SHOPCODE"] = "nclmjczmd"
            df1["SHOPNAME"] = "EC萌洁齿专卖店"
        elif filename.find("MOREI旗舰店") >= 0:
            df1["SHOPCODE"] = "moreiqjd"
            df1["SHOPNAME"] = "MOREI旗舰店"
        elif filename.find("橙意满满化妆品专营店") >= 0:
            df1["SHOPCODE"] = "cymmhzpzyd"
            df1["SHOPNAME"] = "橙意满满化妆品专营店"
        elif filename.find("控师化妆品专营店") >= 0:
            df1["SHOPCODE"] = "kshzpzyd"
            df1["SHOPNAME"] = "控师化妆品专营店"
        elif filename.find("萌齿洁旗舰店") >= 0:
            df1["SHOPCODE"] = "madesqjd"
            df1["SHOPNAME"] = "萌齿洁旗舰店"
        elif filename.find("魔法符号宏炽专卖店") >= 0:
            df1["SHOPCODE"] = "unixylshzzmd"
            df1["SHOPNAME"] = "魔法符号宏炽专卖店"
        elif filename.find("航星个护专营店") >= 0:
            df1["SHOPCODE"] = "hxzmd"
            df1["SHOPNAME"] = "航星个护专营店"
        elif filename.find("魔法符号宏炽专卖店") >= 0:
            df1["SHOPCODE"] = "unixylshzzmd"
            df1["SHOPNAME"] = "魔法符号宏炽专卖店"
        elif filename.find("二十四小时七天个护专营店") >= 0:
            df1["SHOPCODE"] = "essxsqtghzyd"
            df1["SHOPNAME"] = "二十四小时七天个护专营店"
        elif filename.find("无极爽日化用品专营店") >= 0:
            df1["SHOPCODE"] = "wjsrhypzyd"
            df1["SHOPNAME"] = "博滴个护旗舰店"
        elif filename.find("造白个护美妆专营店") >= 0:
            df1["SHOPCODE"] = "zbghmz"
            df1["SHOPNAME"] = "造白个护美妆专营店"
        elif filename.find("mades个护") >= 0:
            df1["SHOPCODE"] = "madesghqjd"
            df1["SHOPNAME"] = "博滴个护旗舰店"
        elif filename.find("驰骄日用品专营店") >= 0:
            df1["SHOPCODE"] = "cjrypzyd"
            df1["SHOPNAME"] = "驰骄日用品专营店"
        elif ((filename.find("补舍日用专营店") >= 0) or (filename.find("博滴官方旗舰店") >= 0)):
            df1["SHOPCODE"] = "bsrypzyd"
            df1["SHOPNAME"] = "补舍日用专营店"
        elif filename.find("博滴日用品专营店") >= 0:
            df1["SHOPCODE"] = "bdrypzyd"
            df1["SHOPNAME"] = "博滴日用品专营店"
        elif filename.find("博滴个护旗舰店") >= 0:
            df1["SHOPCODE"] = "madesghqjd"
            df1["SHOPNAME"] = "博滴个护旗舰店"
        elif filename.find("肌沫化妆品专营店") >= 0:
            df1["SHOPCODE"] = "jmhzpzyd"
            df1["SHOPNAME"] = "肌沫化妆品专营店"
        elif filename.find("肌密泉个人护理专营店") >= 0:
            df1["SHOPCODE"] = "jmqgrhlzyd"
            df1["SHOPNAME"] = "肌密泉个人护理专营店"
        elif filename.find("喝啥个护专营店") >= 0:
            df1["SHOPCODE"] = "hsghzyd"
            df1["SHOPNAME"] = "喝啥个护专营店"
        elif filename.find("航星玩具专营店") >= 0:
            df1["SHOPCODE"] = "hxwjzyd"
            df1["SHOPNAME"] = "航星玩具专营店"
        elif filename.find("商魂专营店") >= 0:
            df1["SHOPCODE"] = "shzyd"
            df1["SHOPNAME"] = "商魂专营店"
        elif filename.find("末隐师个护专营店") >= 0:
            df1["SHOPCODE"] = "mysghzyd"
            df1["SHOPNAME"] = "末隐师个护专营店"
        elif filename.find("魔湾游戏个护专营店") >= 0:
            df1["SHOPCODE"] = "megaxd"
            df1["SHOPNAME"] = "魔湾游戏个护专营店"
        elif filename.find("麦凯莱专营店") >= 0:
            df1["SHOPCODE"] = "mklzyd"
            df1["SHOPNAME"] = "抖音小店-麦凯莱专营店"
        elif filename.find("秀美颜个护专营店") >= 0:
            df1["SHOPCODE"] = "xmyghzyd"
            df1["SHOPNAME"] = "秀美颜个护专营店"
        elif filename.find("配颜师日用化妆专营店") >= 0:
            df1["SHOPCODE"] = "pysryhzzyd"
            df1["SHOPNAME"] = "配颜师日用化妆专营店"
        elif filename.find("植之璨专营店") >= 0:
            df1["SHOPCODE"] = "zzczyd"
            df1["SHOPNAME"] = "植之璨专营店"
        elif filename.find("盈养泉化妆品专营店") >= 0:
            df1["SHOPCODE"] = "yyqhzpzyd"
            df1["SHOPNAME"] = "盈养泉化妆品专营店"
        elif filename.find("配颜师家居用品专营店") >= 0:
            df1["SHOPCODE"] = "pysjjypzyd"
            df1["SHOPNAME"] = "配颜师家居用品专营店"
        elif filename.find("配颜师日用品专营店") >= 0:
            df1["SHOPCODE"] = "pysrypzyd"
            df1["SHOPNAME"] = "配颜师日用品专营店"
        elif filename.find("宝贝化妆品专营店") >= 0:
            df1["SHOPCODE"] = "bbhzpzyd"
            df1["SHOPNAME"] = "宝贝化妆品专营店"
        elif filename.find("肌先知美妆旗舰店") >= 0:
            df1["SHOPCODE"] = "mklzyd"
            df1["SHOPNAME"] = "肌先知美妆旗舰店"

        df1["SHOPNAME"] = df1["SHOPNAME"].apply(lambda x: x if pd.notnull(x) else get_shopcode(plat, x, 1))
        df1["SHOPCODE"] = df1["SHOPCODE"].apply(lambda x: x if pd.notnull(x) else get_shopcode(plat, x, 2))

        df1["TID"] = df1["TID"].astype(str)
        df1["TID"] = df1["TID"].apply(lambda x: x if len(x) > 3 else "nan")
        df1 = df1[~df1["TID"].str.contains("nan")]
        print(df1.head(5).to_markdown())
        print("抖音")
        return df1

    # 快手逻辑
    elif filename.find("快手") >= 0:
        try:
            df = pd.read_excel(filename, sheet_name="月账单", dtype=str)
        except Exception as e:
            df = pd.read_excel(filename, dtype=str)
        for column_name in df.columns:
            df.rename(columns={column_name: column_name.replace(" ", "").replace("\n", "").strip()}, inplace=True)
        # df.dropna(inplace=True)
        print(df.head(5).to_markdown())
        print(df.tail(5).to_markdown())
        if "实际结算时间" in df.columns:
            df.rename(columns={"实际结算时间": "结算时间"}, inplace=True)
        df["结算时间"].replace("nan", np.nan, inplace=True)
        df.dropna(subset=["结算时间"], inplace=True)
        df = df[~df["结算时间"].str.contains("nan")]
        df = df[~df["订单号"].str.contains("nan")]
        print(df.tail(5).to_markdown())
        df["订单实付(元)"] = df["订单实付(元)"].astype(float)
        df["订单退款(元)"] = df["订单退款(元)"].astype(float)
        plat = "KS"
        # 订单实付
        df1 = pd.DataFrame()
        df1["TID"] = df["订单号"]
        df1["SHOPNAME"] = ""
        df1["PLATFORM"] = plat
        df1["SHOPCODE"] = ""
        df1["BILLPLATFORM"] = plat
        df1["CREATED"] = df["结算时间"]
        df1["TITLE"] = df["商品名称"]
        df1["TRADE_TYPE"] = "订单实付"
        df1["BUSINESS_NO"] = df["订单号"] + df["商品ID"] + df["合计收入(元)"] + df["合计支出(元)"]
        df1["INCOME_AMOUNT"] = df["订单实付(元)"]
        df1["EXPEND_AMOUNT"] = 0
        df1["TRADING_CHANNELS"] = df["资金渠道"]
        df1["BUSINESS_DESCRIPTION"] = "订单实付"
        df1["remark"] = ""
        df1["IS_REFUNDAMOUNT"] = 0
        df1["IS_AMOUNT"] = 1
        df1["OID"] = ""
        df1["SOURCEDATA"] = "EXCEL"
        df1["RECIPROCAL_ACCOUNT"] = ""
        df1["BATCHNO"] = ""
        df1["currency"] = ""
        df1["overseas_income"] = ""
        df1["overseas_expend"] = ""
        df1["currency_cny_rate"] = ""

        # 平台补贴
        df2 = df1.copy()
        df2["TRADE_TYPE"] = "平台补贴"
        df2["INCOME_AMOUNT"] = df.apply(lambda x: x["平台补贴(元)"] if x["订单实付(元)"] > 0 else 0, axis=1)
        df2["EXPEND_AMOUNT"] = df.apply(lambda x: x["平台补贴(元)"] if ((x["订单实付(元)"] == 0) & (x["订单退款(元)"] > 0)) else 0,
                                        axis=1)
        df2["BUSINESS_DESCRIPTION"] = "平台补贴"
        df2["INCOME_AMOUNT"] = df2["INCOME_AMOUNT"].astype(float).abs()
        df2["EXPEND_AMOUNT"] = -df2["EXPEND_AMOUNT"].astype(float).abs()
        df2["IS_REFUNDAMOUNT"] = df2["EXPEND_AMOUNT"].apply(lambda x: 1 if x < 0 else 0)
        df2["IS_AMOUNT"] = df2["INCOME_AMOUNT"].apply(lambda x: 1 if x > 0 else 0)

        # 技术服务费
        df3 = df1.copy()
        df3["TRADE_TYPE"] = "技术服务费"
        df3["INCOME_AMOUNT"] = df.apply(lambda x: x["技术服务费(元)"] if ((x["订单实付(元)"] == 0) & (x["订单退款(元)"] > 0)) else 0,
                                        axis=1)
        df3["EXPEND_AMOUNT"] = df.apply(lambda x: x["技术服务费(元)"] if x["订单实付(元)"] > 0 else 0, axis=1)
        df3["BUSINESS_DESCRIPTION"] = "技术服务费"
        df3["IS_REFUNDAMOUNT"] = 0
        df3["IS_AMOUNT"] = 0

        # 订单退款
        df4 = df1.copy()
        df4["TRADE_TYPE"] = "订单退款"
        df4["INCOME_AMOUNT"] = 0
        df4["EXPEND_AMOUNT"] = df["订单退款(元)"]
        df4["BUSINESS_DESCRIPTION"] = "订单退款"
        df4["IS_REFUNDAMOUNT"] = 1
        df4["IS_AMOUNT"] = 0

        # 花呗服务费
        try:
            df5 = df1.copy()
            df5["TRADE_TYPE"] = "花呗服务费"
            df5["INCOME_AMOUNT"] = df.apply(lambda x: x["花呗服务费"] if ((x["订单实付(元)"] == 0) & (x["订单退款(元)"] > 0)) else 0,
                                            axis=1)
            df5["EXPEND_AMOUNT"] = df.apply(lambda x: x["花呗服务费"] if x["订单实付(元)"] > 0 else 0, axis=1)
            df5["BUSINESS_DESCRIPTION"] = "花呗服务费"
            df5["IS_REFUNDAMOUNT"] = 0
            df5["IS_AMOUNT"] = 0
        except Exception as e:
            dict = {"TID": "nan", "SHOPNAME": "", "PLATFORM": "", "SHOPCODE": "", "BILLPLATFORM": "", "CREATED": "",
                    "TITLE": "", "TRADE_TYPE": "", "BUSINESS_NO": "", "INCOME_AMOUNT": "", "EXPEND_AMOUNT": "",
                    "TRADING_CHANNELS": "", "BUSINESS_DESCRIPTION": "", "remark": "", "BUSINESS_BILL_SOURCE": "",
                    "IS_REFUNDAMOUNT": "", "IS_AMOUNT": "", "OID": "", "SOURCEDATA": "", "RECIPROCAL_ACCOUNT": "",
                    "BATCHNO": ""}
            df5 = pd.DataFrame(dict, index=[0])

        # 推广者佣金
        try:
            df6 = df1.copy()
            df6["TRADE_TYPE"] = "推广者佣金"
            df6["INCOME_AMOUNT"] = df.apply(
                lambda x: x["推广者佣金(元)"] if ((x["订单实付(元)"] == 0) & (x["订单退款(元)"] > 0)) else 0, axis=1)
            df6["EXPEND_AMOUNT"] = df.apply(lambda x: x["推广者佣金(元)"] if x["订单实付(元)"] > 0 else 0, axis=1)
            df6["BUSINESS_DESCRIPTION"] = "推广者佣金"
            df6["IS_REFUNDAMOUNT"] = 0
            df6["IS_AMOUNT"] = 0
        except Exception as e:
            dict = {"TID": "nan", "SHOPNAME": "", "PLATFORM": "", "SHOPCODE": "", "BILLPLATFORM": "", "CREATED": "",
                    "TITLE": "", "TRADE_TYPE": "", "BUSINESS_NO": "", "INCOME_AMOUNT": "", "EXPEND_AMOUNT": "",
                    "TRADING_CHANNELS": "", "BUSINESS_DESCRIPTION": "", "remark": "", "BUSINESS_BILL_SOURCE": "",
                    "IS_REFUNDAMOUNT": "", "IS_AMOUNT": "", "OID": "", "SOURCEDATA": "", "RECIPROCAL_ACCOUNT": "",
                    "BATCHNO": ""}
            df6 = pd.DataFrame(dict, index=[0])

        # 达人佣金
        try:
            df7 = df1.copy()
            df7["TRADE_TYPE"] = "达人佣金"
            df7["INCOME_AMOUNT"] = df.apply(
                lambda x: x["达人佣金(元)"] if ((x["订单实付(元)"] == 0) & (x["订单退款(元)"] > 0)) else 0, axis=1)
            df7["EXPEND_AMOUNT"] = df.apply(lambda x: x["达人佣金(元)"] if x["订单实付(元)"] > 0 else 0, axis=1)
            df7["BUSINESS_DESCRIPTION"] = "达人佣金"
            df7["IS_REFUNDAMOUNT"] = 0
            df7["IS_AMOUNT"] = 0
        except Exception as e:
            dict = {"TID": "nan", "SHOPNAME": "", "PLATFORM": "", "SHOPCODE": "", "BILLPLATFORM": "", "CREATED": "",
                    "TITLE": "", "TRADE_TYPE": "", "BUSINESS_NO": "", "INCOME_AMOUNT": "", "EXPEND_AMOUNT": "",
                    "TRADING_CHANNELS": "", "BUSINESS_DESCRIPTION": "", "remark": "", "BUSINESS_BILL_SOURCE": "",
                    "IS_REFUNDAMOUNT": "", "IS_AMOUNT": "", "OID": "", "SOURCEDATA": "", "RECIPROCAL_ACCOUNT": "",
                    "BATCHNO": ""}
            df7 = pd.DataFrame(dict, index=[0])

        # 团长佣金
        try:
            df8 = df1.copy()
            df8["TRADE_TYPE"] = "团长佣金"
            df8["INCOME_AMOUNT"] = df.apply(
                lambda x: x["团长佣金(元)"] if ((x["订单实付(元)"] == 0) & (x["订单退款(元)"] > 0)) else 0, axis=1)
            df8["EXPEND_AMOUNT"] = df.apply(lambda x: x["团长佣金(元)"] if x["订单实付(元)"] > 0 else 0, axis=1)
            df8["BUSINESS_DESCRIPTION"] = "团长佣金"
            df8["IS_REFUNDAMOUNT"] = 0
            df8["IS_AMOUNT"] = 0
        except Exception as e:
            dict = {"TID": "nan", "SHOPNAME": "", "PLATFORM": "", "SHOPCODE": "", "BILLPLATFORM": "", "CREATED": "",
                    "TITLE": "", "TRADE_TYPE": "", "BUSINESS_NO": "", "INCOME_AMOUNT": "", "EXPEND_AMOUNT": "",
                    "TRADING_CHANNELS": "", "BUSINESS_DESCRIPTION": "", "remark": "", "BUSINESS_BILL_SOURCE": "",
                    "IS_REFUNDAMOUNT": "", "IS_AMOUNT": "", "OID": "", "SOURCEDATA": "", "RECIPROCAL_ACCOUNT": "",
                    "BATCHNO": ""}
            df8 = pd.DataFrame(dict, index=[0])

        # 快赚客佣金
        try:
            df9 = df1.copy()
            df9["TRADE_TYPE"] = "快赚客佣金"
            df9["INCOME_AMOUNT"] = df.apply(
                lambda x: x["快赚客佣金(元)"] if ((x["订单实付(元)"] == 0) & (x["订单退款(元)"] > 0)) else 0, axis=1)
            df9["EXPEND_AMOUNT"] = df.apply(lambda x: x["快赚客佣金(元)"] if x["订单实付(元)"] > 0 else 0, axis=1)
            df9["BUSINESS_DESCRIPTION"] = "快赚客佣金"
            df9["IS_REFUNDAMOUNT"] = 0
            df9["IS_AMOUNT"] = 0
        except Exception as e:
            dict = {"TID": "nan", "SHOPNAME": "", "PLATFORM": "", "SHOPCODE": "", "BILLPLATFORM": "", "CREATED": "",
                    "TITLE": "", "TRADE_TYPE": "", "BUSINESS_NO": "", "INCOME_AMOUNT": "", "EXPEND_AMOUNT": "",
                    "TRADING_CHANNELS": "", "BUSINESS_DESCRIPTION": "", "remark": "", "BUSINESS_BILL_SOURCE": "",
                    "IS_REFUNDAMOUNT": "", "IS_AMOUNT": "", "OID": "", "SOURCEDATA": "", "RECIPROCAL_ACCOUNT": "",
                    "BATCHNO": ""}
            df9 = pd.DataFrame(dict, index=[0])

        # 平台补贴
        try:
            df10 = df1.copy()
            df10["TRADE_TYPE"] = "主播补贴"
            df10["INCOME_AMOUNT"] = df.apply(lambda x: x["主播补贴(元)"] if x["订单实付(元)"] > 0 else 0, axis=1)
            df10["EXPEND_AMOUNT"] = df.apply(
                lambda x: x["主播补贴(元)"] if ((x["订单实付(元)"] == 0) & (x["订单退款(元)"] > 0)) else 0, axis=1)
            df10["BUSINESS_DESCRIPTION"] = "主播补贴"
            df10["INCOME_AMOUNT"] = df10["INCOME_AMOUNT"].astype(float).abs()
            df10["EXPEND_AMOUNT"] = -df10["EXPEND_AMOUNT"].astype(float).abs()
            df10["IS_REFUNDAMOUNT"] = df10["EXPEND_AMOUNT"].apply(lambda x: 1 if x < 0 else 0)
            df10["IS_AMOUNT"] = df10["INCOME_AMOUNT"].apply(lambda x: 1 if x > 0 else 0)
        except Exception as e:
            dict = {"TID": "nan", "SHOPNAME": "", "PLATFORM": "", "SHOPCODE": "", "BILLPLATFORM": "", "CREATED": "",
                    "TITLE": "", "TRADE_TYPE": "", "BUSINESS_NO": "", "INCOME_AMOUNT": "", "EXPEND_AMOUNT": "",
                    "TRADING_CHANNELS": "", "BUSINESS_DESCRIPTION": "", "remark": "", "BUSINESS_BILL_SOURCE": "",
                    "IS_REFUNDAMOUNT": "", "IS_AMOUNT": "", "OID": "", "SOURCEDATA": "", "RECIPROCAL_ACCOUNT": "",
                    "BATCHNO": ""}
            df10 = pd.DataFrame(dict, index=[0])

        dfs = [df1, df2, df3, df4, df5, df6, df7, df8, df9, df10]
        df1 = pd.concat(dfs)
        df1 = df1[~df1["TID"].str.contains("nan")]
        print(df1.tail(5).to_markdown())
        df1["INCOME_AMOUNT"] = df1["INCOME_AMOUNT"].astype("float64").abs()
        df1["EXPEND_AMOUNT"] = -df1["EXPEND_AMOUNT"].astype("float64").abs()
        if ((filename.find("BODYAID播地艾") >= 0) | (filename.find("博滴播地艾") >= 0)):
            df1["SHOPNAME"] = "博滴品牌专卖店"
            df1["SHOPCODE"] = "bodyppzmd"
        elif filename.find("Bodyaid美妆") >= 0:
            df1["SHOPNAME"] = "bodyaid美妆专营店"
            df1["SHOPCODE"] = "bodyaidmzzyd"
        elif filename.find("Montooth品牌店") >= 0:
            df1["SHOPNAME"] = "Montooth品牌店"
            df1["SHOPCODE"] = "montoothppd"
        elif filename.find("Montooth品牌") >= 0:
            df1["SHOPNAME"] = "Montooth品牌"
            df1["SHOPCODE"] = "zbmklqjd"
        elif ((filename.find("MONTOOTH专营") >= 0) | (filename.find("Montooth专营") >= 0)):
            df1["SHOPNAME"] = "Montooth专营店"
            df1["SHOPCODE"] = "montoothzyd"
        elif filename.find("morei官方旗舰") >= 0:
            df1["SHOPNAME"] = "morei官方旗舰店"
            df1["SHOPCODE"] = "moreigf"
        elif filename.find("MOREI家居") >= 0:
            df1["SHOPNAME"] = "MOREI家居品牌"
            df1["SHOPCODE"] = "morijj"
        elif ((filename.find("Vitabloom美妆") >= 0) | (filename.find("播地艾美妆") >= 0)):
            df1["SHOPNAME"] = "Vitabloom美妆店"
            df1["SHOPCODE"] = "vitabloommzd"
        elif filename.find("艾法博滴品牌专营") >= 0:
            df1["SHOPNAME"] = "博滴艾法店"
            df1["SHOPCODE"] = "zzcafzyd"
        elif filename.find("艾法专营") >= 0:
            df1["SHOPNAME"] = "博滴艾法店"
            df1["SHOPCODE"] = "zzcafzyd"
        elif ((filename.find("薄滴尚西") >= 0) | (filename.find("博滴尚西") >= 0)):
            df1["SHOPNAME"] = "博滴广州尚西专卖店"
            df1["SHOPCODE"] = "bdgzsxzmd"
        elif filename.find("播地艾个护") >= 0:
            df1["SHOPNAME"] = "睿旗个护店"
            df1["SHOPCODE"] = "bdaghzyd"
        elif filename.find("博滴个护美妆") >= 0:
            df1["SHOPNAME"] = "博滴个护美妆专营店"
            df1["SHOPCODE"] = "bdghmzzyd"
        elif filename.find("博滴个护品牌") >= 0:
            df1["SHOPNAME"] = "博滴个护品牌店"
            df1["SHOPCODE"] = "bodiqjd"
        elif filename.find("博滴官方旗舰店") >= 0:
            df1["SHOPNAME"] = "博滴官方旗舰店"
            df1["SHOPCODE"] = "bdgf"
        elif (filename.find("博滴旗舰店") >= 0):
            df1["SHOPNAME"] = "博滴旗舰店"
            df1["SHOPCODE"] = "bdqjd"
        elif filename.find("博滴品牌专卖") >= 0:
            df1["SHOPNAME"] = "博滴品牌专卖店"
            df1["SHOPCODE"] = "bodyppzmd"
        elif filename.find("博滴若蘅") >= 0:
            df1["SHOPNAME"] = "博滴若蘅专卖店"
            df1["SHOPCODE"] = "bodyrhzmd"
        elif filename.find("morei魔妆专卖店") >= 0:
            df1["SHOPNAME"] = "morei魔妆专卖店"
            df1["SHOPCODE"] = "mmz"
        elif filename.find("多瑞专卖") >= 0:
            df1["SHOPNAME"] = "多瑞专卖店"
            df1["SHOPCODE"] = "drzmd"
        elif filename.find("博滴个护店") >= 0:
            df1["SHOPNAME"] = "博滴个护店"
            df1["SHOPCODE"] = "bdghd"
        elif filename.find("归于瘾美妆") >= 0:
            df1["SHOPNAME"] = "归于瘾美妆店"
            df1["SHOPCODE"] = "gyymzd"
        elif filename.find("宏炽美妆") >= 0:
            df1["SHOPNAME"] = "宏炽美妆专营店"
            df1["SHOPCODE"] = "hcmzzyd"
        elif filename.find("商魂个护专营店") >= 0:
            df1["SHOPNAME"] = "商魂个护专营店"
            df1["SHOPCODE"] = "shghzyd"
        elif filename.find("护肤达人") >= 0:
            df1["SHOPNAME"] = "护肤达人"
            df1["SHOPCODE"] = "yyqhfdr"
        elif filename.find("可瘾美妆") >= 0:
            df1["SHOPNAME"] = "可瘾惠优购"
            df1["SHOPCODE"] = "kymzzyd"
        elif filename.find("可瘾品牌") >= 0:
            df1["SHOPNAME"] = "可瘾品牌店"
            df1["SHOPCODE"] = "mjckyppd"
        elif filename.find("鲁文美妆") >= 0:
            df1["SHOPNAME"] = "鲁文美妆专营店"
            df1["SHOPCODE"] = "lwmzzyd"
        elif ((filename.find("麦凯莱专营") >= 0) | (filename.find("魔妆专场") >= 0)):
            df1["SHOPNAME"] = "麦凯莱专营店"
            df1["SHOPCODE"] = "mklmzzyd"
        elif filename.find("萌齿洁专营") >= 0:
            df1["SHOPNAME"] = "肯妮诗个护美妆店"
            df1["SHOPCODE"] = "knsghmzd"
        elif filename.find("萌洁齿个护") >= 0:
            df1["SHOPNAME"] = "DentylActive航星专营店/萌洁齿个护店"
            df1["SHOPCODE"] = "mjcgh"
        elif filename.find("萌洁齿睿旗") >= 0:
            df1["SHOPNAME"] = "萌洁齿睿旗专卖店"
            df1["SHOPCODE"] = "mjcrqzmd"
        elif filename.find("配颜师嘉兴") >= 0:
            df1["SHOPNAME"] = "配颜师嘉兴小店"
            df1["SHOPCODE"] = "pysjxmjyxzyd"
        elif filename.find("睿旗个护") >= 0:
            df1["SHOPNAME"] = "睿旗个护店"
            df1["SHOPCODE"] = "bdaghzyd"
        elif filename.find("商魂美妆个护") >= 0:
            df1["SHOPNAME"] = "商魂美妆个护专营店"
            df1["SHOPCODE"] = "shmzghzyd"
        elif filename.find("尚西美妆专营") >= 0:
            df1["SHOPNAME"] = "尚西美妆专营店"
            df1["SHOPCODE"] = "sxmzzyd"
        elif filename.find("深圳艾法商贸") >= 0:
            df1["SHOPNAME"] = "深圳艾法商贸"
            df1["SHOPCODE"] = "szafsm"
        elif filename.find("星空专营") >= 0:
            df1["SHOPNAME"] = ""
            df1["SHOPCODE"] = ""
        elif filename.find("秀美颜美妆个护") >= 0:
            df1["SHOPNAME"] = "秀美颜美妆个护专营店"
            df1["SHOPCODE"] = "xmymzghzyd"
        elif filename.find("植之璨品牌") >= 0:
            df1["SHOPNAME"] = "植之璨品牌店"
            df1["SHOPCODE"] = "zzc"
        elif filename.find("Vitabloom专营") >= 0:
            df1["SHOPNAME"] = "Vitabloom专营店"
            df1["SHOPCODE"] = "zzczyd"
        elif filename.find("montooth睿旗专卖") >= 0:
            df1["SHOPNAME"] = "montooth睿旗专卖"
            df1["SHOPCODE"] = "mothtoothrqzmd"
        elif filename.find("vitaBloom旗舰店") >= 0:
            df1["SHOPNAME"] = "vitaBloom旗舰店"
            df1["SHOPCODE"] = "vitaBloom"
        elif filename.find("博滴老板") >= 0:
            df1["SHOPNAME"] = "博滴老板"
            df1["SHOPCODE"] = "bblb"
        elif filename.find("美妆博主") >= 0:
            df1["SHOPNAME"] = "美妆博主"
            df1["SHOPCODE"] = "mzbz"
        elif filename.find("萌齿洁品牌专营店") >= 0:
            df1["SHOPNAME"] = "鲁文美妆专营店"
            df1["SHOPCODE"] = "lwmzzyd"
        elif filename.find("香姐严选") >= 0:
            df1["SHOPNAME"] = "香姐严选"
            df1["SHOPCODE"] = "xjyx"
        elif filename.find("鑫桂个护美妆店") >= 0:
            df1["SHOPNAME"] = "鑫桂个护美妆店"
            df1["SHOPCODE"] = "xgghmzd"
        elif filename.find("植之璨泡研专卖店") >= 0:
            df1["SHOPNAME"] = "植之璨泡研专卖店"
            df1["SHOPCODE"] = "zzcpyzmd"
        elif filename.find("宝贝港湾的小品牌店") >= 0:
            df1["SHOPNAME"] = "宝贝港湾的小品牌店"
            df1["SHOPCODE"] = "bbgw"
        elif filename.find("泡研泡研") >= 0:
            df1["SHOPNAME"] = "泡研泡研"
            df1["SHOPCODE"] = "pypy"
        elif filename.find("家政管理师佩佩") >= 0:
            df1["SHOPNAME"] = "家政管理师佩佩"
            df1["SHOPCODE"] = "jzglspp"
        elif filename.find("morei家清店") >= 0:
            df1["SHOPNAME"] = "MOREI家清店"
            df1["SHOPCODE"] = "mjqd"
        print(df1.head(5).to_markdown())
        print("快手")
        return df1

    # 淘宝&阿里巴巴&天猫逻辑
    elif ((filename.find("淘宝") >= 0) or (filename.find("天猫") >= 0)):
        if ((filename.find("海外") >= 0) & (filename.find("逍遥") < 0)):
            if filename.find("settle_") >= 0:
                if filename.find("csv") >= 0:
                    df = pd.read_csv(filename, dtype=str, encoding="gb18030")
                elif filename.find("xls") >= 0:
                    df = pd.read_excel(filename, dtype=str)
                else:
                    dict = {"TID": "nan", "SHOPNAME": "", "PLATFORM": "", "SHOPCODE": "", "BILLPLATFORM": "",
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
                for column_name in df.columns:
                    df.rename(columns={column_name: column_name.replace(" ", "").replace("\n", "").strip()},
                              inplace=True)
                print(df.head().to_markdown())
                df["Rmb_amount"] = df["Rmb_amount"].astype(float)
                df["Amount"] = df["Amount"].astype(float)

                if filename.find("淘宝") >= 0:
                    plat = "TAOBAO"
                elif filename.find("天猫") >= 0:
                    plat = "TMALL"

                df1 = pd.DataFrame()
                df1["TID"] = df["Original_partner_transaction_ID"].astype(str).apply(
                    lambda x: x.replace(" ", "").strip())
                df1["SHOPNAME"] = ""
                df1["PLATFORM"] = plat
                df1["SHOPCODE"] = ""
                df1["BILLPLATFORM"] = "OVERSEASZFB"
                df1["CREATED"] = df["Settlement_time"]
                df1["TITLE"] = df["Remarks"]
                df1["TRADE_TYPE"] = df.apply(lambda x: "Amount" if x["Rmb_amount"] > 0 else "Refund", axis=1)
                df1["BUSINESS_NO"] = df["Partner_transaction_id"]
                df1["INCOME_AMOUNT"] = df["Rmb_amount"].apply(lambda x: x if x > 0 else 0)
                df1["EXPEND_AMOUNT"] = df["Rmb_amount"].apply(lambda x: x if x < 0 else 0)
                df1["TRADING_CHANNELS"] = ""
                df1["BUSINESS_DESCRIPTION"] = df.apply(lambda x: "Amount" if x["Rmb_amount"] > 0 else "Refund", axis=1)
                df1["remark"] = ""
                df1["BUSINESS_BILL_SOURCE"] = ""
                df1["IS_REFUNDAMOUNT"] = df.apply(lambda x: 1 if x["Rmb_amount"] < 0 else 0, axis=1)
                df1["IS_AMOUNT"] = df.apply(lambda x: 1 if x["Rmb_amount"] > 0 else 0, axis=1)
                df1["OID"] = ""
                df1["SOURCEDATA"] = "EXCEL"
                df1["RECIPROCAL_ACCOUNT"] = ""
                df1["BATCHNO"] = ""
                df1["currency"] = df["Currency"].apply(lambda x: x.replace(" ", "").strip())
                df1["overseas_income"] = df["Amount"].apply(lambda x: x if x > 0 else 0)
                df1["overseas_expend"] = df["Amount"].apply(lambda x: x if x < 0 else 0)
                df1["currency_cny_rate"] = df["Rate"].apply(lambda x: x.replace(" ", "").strip())
                df1 = df1.loc[~((df1.INCOME_AMOUNT == 0) & (df1.EXPEND_AMOUNT == 0))]

                if ((filename.find("entyl") >= 0) & (filename.find("海外") >= 0)):
                    df1["SHOPCODE"] = "dentylactiveov"
                    df1["SHOPNAME"] = "dentylactive海外旗舰店"
                elif filename.find("loshi海外") >= 0:
                    df1["SHOPCODE"] = "loshiov"
                    df1["SHOPNAME"] = "loshi海外旗舰店"
                elif filename.find("ades海外") >= 0:
                    df1["SHOPCODE"] = "madesov"
                    df1["SHOPNAME"] = "mades海外旗舰店"
                elif ((filename.find("lcn海外") >= 0) | (filename.find("LCN海外") >= 0)):
                    df1["SHOPCODE"] = "lcnov"
                    df1["SHOPNAME"] = "lcn海外旗舰店"
                elif filename.find("met海外") >= 0:
                    df1["SHOPCODE"] = "manukascosmetov"
                    df1["SHOPNAME"] = "manukascosmet海外旗舰店"
                elif filename.find("ule海外") >= 0:
                    df1["SHOPCODE"] = "nuggelasuleov"
                    df1["SHOPNAME"] = "nuggelasule海外旗舰店"
                elif filename.find("rai海外") >= 0:
                    df1["SHOPCODE"] = "samouraiov"
                    df1["SHOPNAME"] = "samourai海外旗舰店"
                elif filename.find("ex海外") >= 0:
                    df1["SHOPCODE"] = "ultradexov"
                    df1["SHOPNAME"] = "ultradex海外旗舰店"
                elif ((filename.find("ambra海外") >= 0) | (filename.find("chiara海外") >= 0)):
                    df1["SHOPCODE"] = "chiaraambraov"
                    df1["SHOPNAME"] = "chiaraambra海外旗舰店"
                elif ((filename.find("Fit海外") >= 0) | (filename.find("cora海外") >= 0)):
                    df1["SHOPCODE"] = "dicoraurbanfitov"
                    df1["SHOPNAME"] = "dicoraurbanfit海外旗舰店"
                elif filename.find("iwhite海外") >= 0:
                    df1["SHOPCODE"] = "iwhiteovqjd"
                    df1["SHOPNAME"] = "iwhite海外旗舰店"
                elif ((filename.find("image海外") >= 0) | (filename.find("IMAGE海外") >= 0) | (
                        filename.find("imega海外") >= 0)):
                    df1["SHOPCODE"] = "swissimageov"
                    df1["SHOPNAME"] = "swissimage海外旗舰店"
                elif ((filename.find("个护海外") >= 0) | (filename.find("mage个护海外") >= 0)):
                    df1["SHOPCODE"] = "megaggov"
                    df1["SHOPNAME"] = "mega个护海外专营店"
                elif filename.find("护理海外") >= 0:
                    df1["SHOPCODE"] = "megagrhlhwzydov"
                    df1["SHOPNAME"] = "mega个人护理海外专营店"
                elif filename.find("lab海外") >= 0:
                    df1["SHOPCODE"] = "smilelabov"
                    df1["SHOPNAME"] = "smilelab海外旗舰店"
                elif filename.find("电器海外") >= 0:
                    df1["SHOPCODE"] = "unixdqov"
                    df1["SHOPNAME"] = "unix电器海外旗舰店"
                elif filename.find("lilac海外") >= 0:
                    df1["SHOPCODE"] = "lilacov"
                    df1["SHOPNAME"] = "lilac海外旗舰店"
                print(df1.head().to_markdown())

            elif filename.find("fee") >= 0:
                df = pd.read_csv(filename, dtype=str, encoding="gb18030")
                for column_name in df.columns:
                    df.rename(columns={column_name: column_name.replace(" ", "").replace("\n", "").strip()},
                              inplace=True)
                print(df.head().to_markdown())
                df["Fee_rmb_amount"] = df["Fee_rmb_amount"].astype(float)
                df["Fee_amount"] = df["Fee_amount"].astype(float)
                yearmonth = "".join("".join(filename.split("_")[-1:]).split(".")[:1])
                print(yearmonth)
                # if ((yearmonth[-2:].find("02")>=0)&(yearmonth.find("2020")>=0)):
                if ((yearmonth[-2:].find("2020-02") >= 0) | (yearmonth.find("202002") >= 0)):
                    time = yearmonth + "29 23:59:59"
                elif ((yearmonth[-2:].find("02") >= 0) & (yearmonth.find("2020") < 0)):
                    time = yearmonth + "28 23:59:59"
                elif ((yearmonth[-2:].find("04") >= 0) | (yearmonth[-2:].find("06") >= 0) | (
                        yearmonth[-2:].find("09") >= 0) | (yearmonth[-2:].find("11") >= 0)):
                    time = yearmonth + "30 23:59:59"
                else:
                    time = yearmonth + "31 23:59:59"
                print(time)
                bill_time = datetime.datetime.strptime(time, "%Y%m%d %H:%M:%S").strftime("%Y-%m-%d %H:%M:%S")
                df["Original_partner_transaction_ID"] = df["Original_partner_transaction_ID"].str.replace('\s+', '')
                df["Partner_transaction_id"] = df["Partner_transaction_id"].str.replace('\s+', '')
                df["Transaction_id"] = df["Transaction_id"].str.replace('\s+', '')

                # print(df[df["Transaction_id"].str.contains("2018102022001237645405453666")])

                if filename.find("淘宝") >= 0:
                    plat = "TAOBAO"
                elif filename.find("天猫") >= 0:
                    plat = "TMALL"

                df1 = pd.DataFrame()
                df["TID"] = df.apply(
                    lambda x: x["Original_partner_transaction_ID"] if x["Original_partner_transaction_ID"].find(
                        "nan") == 0 else x["Partner_transaction_id"], axis=1)
                df1["TID"] = df["TID"]
                # df1["TID"] = df.apply(lambda x:x["TID"] if x["TID"].find("nan")==0 else x["Transaction_id"],axis=1)
                df1["SHOPNAME"] = ""
                df1["PLATFORM"] = plat
                df1["SHOPCODE"] = ""
                df1["BILLPLATFORM"] = "OVERSEASZFB"
                df1["CREATED"] = bill_time
                df1["TITLE"] = ""
                df1["TRADE_TYPE"] = df["Fee_type"].str.replace('\d+', '')
                df1["BUSINESS_NO"] = df["Transaction_id"]
                df1["INCOME_AMOUNT"] = df["Fee_rmb_amount"].apply(lambda x: x if x < 0 else 0)
                df1["EXPEND_AMOUNT"] = df["Fee_rmb_amount"].apply(lambda x: x if x > 0 else 0)
                df1["TRADING_CHANNELS"] = ""
                df1["BUSINESS_DESCRIPTION"] = df["Fee_type"].str.replace('\d+', '')
                df1["remark"] = "Remark：" + df["Remark"] + "。Partner_transaction_id：" + df["Partner_transaction_id"]
                df1["BUSINESS_BILL_SOURCE"] = ""
                df1["IS_REFUNDAMOUNT"] = 0
                df1["IS_AMOUNT"] = 0
                df1["OID"] = ""
                df1["SOURCEDATA"] = "EXCEL"
                df1["RECIPROCAL_ACCOUNT"] = ""
                df1["BATCHNO"] = ""
                df1["currency"] = df["Currency"].apply(lambda x: x.replace(" ", "").strip())
                df1["overseas_income"] = df["Fee_amount"].apply(lambda x: x if x < 0 else 0)
                df1["overseas_expend"] = df["Fee_amount"].apply(lambda x: x if x > 0 else 0)
                df1["currency_cny_rate"] = df["Rate"].apply(lambda x: x.replace(" ", "").strip())
                df1 = df1.loc[~((df1.INCOME_AMOUNT == 0) & (df1.EXPEND_AMOUNT == 0))]
                df1["INCOME_AMOUNT"] = df1["INCOME_AMOUNT"].abs()
                df1["EXPEND_AMOUNT"] = -df1["EXPEND_AMOUNT"].abs()
                df1["overseas_income"] = df1["overseas_income"].abs()
                df1["overseas_expend"] = -df1["overseas_expend"].abs()

                if ((filename.find("entyl") >= 0) & (filename.find("海外") >= 0)):
                    df1["SHOPCODE"] = "dentylactiveov"
                    df1["SHOPNAME"] = "dentylactive海外旗舰店"
                elif filename.find("loshi海外") >= 0:
                    df1["SHOPCODE"] = "loshiov"
                    df1["SHOPNAME"] = "loshi海外旗舰店"
                elif filename.find("ades海外") >= 0:
                    df1["SHOPCODE"] = "madesov"
                    df1["SHOPNAME"] = "mades海外旗舰店"
                elif ((filename.find("lcn海外") >= 0) | (filename.find("LCN海外") >= 0)):
                    df1["SHOPCODE"] = "lcnov"
                    df1["SHOPNAME"] = "lcn海外旗舰店"
                elif filename.find("met海外") >= 0:
                    df1["SHOPCODE"] = "manukascosmetov"
                    df1["SHOPNAME"] = "manukascosmet海外旗舰店"
                elif filename.find("ule海外") >= 0:
                    df1["SHOPCODE"] = "nuggelasuleov"
                    df1["SHOPNAME"] = "nuggelasule海外旗舰店"
                elif filename.find("rai海外") >= 0:
                    df1["SHOPCODE"] = "samouraiov"
                    df1["SHOPNAME"] = "samourai海外旗舰店"
                elif filename.find("ex海外") >= 0:
                    df1["SHOPCODE"] = "ultradexov"
                    df1["SHOPNAME"] = "ultradex海外旗舰店"
                elif ((filename.find("ambra海外") >= 0) | (filename.find("chiara海外") >= 0)):
                    df1["SHOPCODE"] = "chiaraambraov"
                    df1["SHOPNAME"] = "chiaraambra海外旗舰店"
                elif ((filename.find("Fit海外") >= 0) | (filename.find("cora海外") >= 0)):
                    df1["SHOPCODE"] = "dicoraurbanfitov"
                    df1["SHOPNAME"] = "dicoraurbanfit海外旗舰店"
                elif filename.find("iwhite海外") >= 0:
                    df1["SHOPCODE"] = "iwhiteovqjd"
                    df1["SHOPNAME"] = "iwhite海外旗舰店"
                elif ((filename.find("个护海外") >= 0) | (filename.find("mage个护海外") >= 0)):
                    df1["SHOPCODE"] = "megaggov"
                    df1["SHOPNAME"] = "mega个护海外专营店"
                elif filename.find("护理海外") >= 0:
                    df1["SHOPCODE"] = "megagrhlhwzydov"
                    df1["SHOPNAME"] = "mega个人护理海外专营店"
                elif filename.find("lab海外") >= 0:
                    df1["SHOPCODE"] = "smilelabov"
                    df1["SHOPNAME"] = "smilelab海外旗舰店"
                elif ((filename.find("image海外") >= 0) | (filename.find("IMAGE海外") >= 0) | (
                        filename.find("imega海外") >= 0)):
                    df1["SHOPCODE"] = "swissimageov"
                    df1["SHOPNAME"] = "swissimage海外旗舰店"
                elif filename.find("电器海外") >= 0:
                    df1["SHOPCODE"] = "unixdqov"
                    df1["SHOPNAME"] = "unix电器海外旗舰店"
                elif filename.find("lilac海外") >= 0:
                    df1["SHOPCODE"] = "lilacov"
                    df1["SHOPNAME"] = "lilac海外旗舰店"
                print(df1.head().to_markdown())

            elif filename.find("支付宝") >= 0:
                if filename.find("xls") >= 0:
                    try:
                        df = pd.read_excel(filename, sheet_name="我的账户其他费用", dtype=str)
                    except Exception as e:
                        dict = {"TID": "nan", "SHOPNAME": "", "PLATFORM": "", "SHOPCODE": "", "BILLPLATFORM": "",
                                "CREATED": "",
                                "TITLE": "", "TRADE_TYPE": "", "BUSINESS_NO": "", "INCOME_AMOUNT": "",
                                "EXPEND_AMOUNT": "",
                                "TRADING_CHANNELS": "", "BUSINESS_DESCRIPTION": "", "remark": "",
                                "BUSINESS_BILL_SOURCE": "",
                                "IS_REFUNDAMOUNT": "", "IS_AMOUNT": "", "OID": "", "SOURCEDATA": "",
                                "RECIPROCAL_ACCOUNT": "",
                                "BATCHNO": "", "currency": "", "overseas_income": "", "overseas_expend": "",
                                "currency_cny_rate": ""}
                        df = pd.DataFrame(dict, index=[0])
                        return df
                elif filename.find("csv") >= 0:
                    try:
                        df = pd.read_csv(filename, dtype=str)
                    except Exception as e:
                        try:
                            df = pd.read_csv(filename, dtype=str, encoding="gb18030")
                        except Exception as e:
                            dict = {"TID": "nan", "SHOPNAME": "", "PLATFORM": "", "SHOPCODE": "", "BILLPLATFORM": "",
                                    "CREATED": "",
                                    "TITLE": "", "TRADE_TYPE": "", "BUSINESS_NO": "", "INCOME_AMOUNT": "",
                                    "EXPEND_AMOUNT": "",
                                    "TRADING_CHANNELS": "", "BUSINESS_DESCRIPTION": "", "remark": "",
                                    "BUSINESS_BILL_SOURCE": "",
                                    "IS_REFUNDAMOUNT": "", "IS_AMOUNT": "", "OID": "", "SOURCEDATA": "",
                                    "RECIPROCAL_ACCOUNT": "",
                                    "BATCHNO": "", "currency": "", "overseas_income": "", "overseas_expend": "",
                                    "currency_cny_rate": ""}
                            df = pd.DataFrame(dict, index=[0])
                            return df
                else:
                    dict = {"TID": "nan", "SHOPNAME": "", "PLATFORM": "", "SHOPCODE": "", "BILLPLATFORM": "",
                            "CREATED": "",
                            "TITLE": "", "TRADE_TYPE": "", "BUSINESS_NO": "", "INCOME_AMOUNT": "",
                            "EXPEND_AMOUNT": "",
                            "TRADING_CHANNELS": "", "BUSINESS_DESCRIPTION": "", "remark": "",
                            "BUSINESS_BILL_SOURCE": "",
                            "IS_REFUNDAMOUNT": "", "IS_AMOUNT": "", "OID": "", "SOURCEDATA": "",
                            "RECIPROCAL_ACCOUNT": "",
                            "BATCHNO": "", "currency": "", "overseas_income": "", "overseas_expend": "",
                            "currency_cny_rate": ""}
                    df = pd.DataFrame(dict, index=[0])
                    return df
                for column_name in df.columns:
                    df.rename(columns={column_name: column_name.replace(" ", "").replace("\n", "").strip()},
                              inplace=True)
                print(df.head().to_markdown())
                df["Amount"] = df["Amount"].astype(float)
                if "Details" in df.columns:
                    df["Details"] = df["Details"].astype(str)

                if filename.find("淘宝") >= 0:
                    plat = "TAOBAO"
                elif filename.find("天猫") >= 0:
                    plat = "TMALL"

                df1 = pd.DataFrame()
                if "OrderNo" in df.columns:
                    df["OrderNo"] = df["OrderNo"].astype(str).apply(lambda x: x.replace(" ", "").strip())
                    df1["TID"] = df.apply(
                        lambda x: "".join(x["Details"].split("淘宝订单号")[-1:]) if x["Details"].find("淘宝订单号") >= 0 else x[
                            "OrderNo"], axis=1)
                else:
                    df1["TID"] = df["Partner_transaction_id"].astype(str).apply(lambda x: x.replace(" ", "").strip())
                    if "Details" in df.columns:
                        df1["TID"] = df.apply(
                            lambda x: "".join(x["Details"].split("淘宝订单号")[-1:]) if x["Details"].find("淘宝订单号") >= 0 else
                            x["Partner_transaction_id"], axis=1)
                    else:
                        df1["TID"] = df["Partner_transaction_id"]
                df1["SHOPNAME"] = ""
                df1["PLATFORM"] = plat
                df1["SHOPCODE"] = ""
                df1["BILLPLATFORM"] = "OVERSEASZFB"
                if "Time" in df.columns:
                    df1["CREATED"] = df["Time"]
                else:
                    df1["CREATED"] = df["Settlement_time"]
                df1["TITLE"] = ""
                if "Type" in df.columns:
                    df1["TRADE_TYPE"] = df["Type"]
                else:
                    df1["TRADE_TYPE"] = df["Type"]
                df1["BUSINESS_NO"] = df["PartnerTransactionID"]
                df1["INCOME_AMOUNT"] = df["Amount"].apply(lambda x: x if x > 0 else 0)
                df1["EXPEND_AMOUNT"] = df["Amount"].apply(lambda x: x if x < 0 else 0)
                df1["TRADING_CHANNELS"] = ""
                if "Details" in df.columns:
                    if "Remarks" in df.columns:
                        df1["BUSINESS_DESCRIPTION"] = df.apply(
                            lambda x: x["Details"] if pd.notnull(x["Details"]) else x["Remarks"], axis=1)
                    else:
                        df1["BUSINESS_DESCRIPTION"] = df["Details"]
                elif "Remarks" in df.columns:
                    df1["BUSINESS_DESCRIPTION"] = df["Remarks"]
                else:
                    df1["BUSINESS_DESCRIPTION"] = ""
                df1["BUSINESS_DESCRIPTION"] = df1["BUSINESS_DESCRIPTION"].str.replace('\d+', '')
                # if "Remark" in df.columns:
                #     df1["remark"] = df["Remark"]
                # else:
                #     df1["remark"] = ""
                df1["remark"] = df["Remarks"]
                df1["BUSINESS_BILL_SOURCE"] = ""
                df1["IS_REFUNDAMOUNT"] = 0
                df1["IS_AMOUNT"] = 0
                df1["OID"] = ""
                df1["SOURCEDATA"] = "EXCEL"
                df1["RECIPROCAL_ACCOUNT"] = ""
                df1["BATCHNO"] = ""
                df1["currency"] = "CNY"
                df1["overseas_income"] = ""
                df1["overseas_expend"] = ""
                df1["currency_cny_rate"] = ""
                df1 = df1.loc[~((df1.INCOME_AMOUNT == 0) & (df1.EXPEND_AMOUNT == 0))]

                if ((filename.find("entyl") >= 0) & (filename.find("海外") >= 0)):
                    df1["SHOPCODE"] = "dentylactiveov"
                    df1["SHOPNAME"] = "dentylactive海外旗舰店"
                elif filename.find("loshi海外") >= 0:
                    df1["SHOPCODE"] = "loshiov"
                    df1["SHOPNAME"] = "loshi海外旗舰店"
                elif filename.find("ades海外") >= 0:
                    df1["SHOPCODE"] = "madesov"
                    df1["SHOPNAME"] = "mades海外旗舰店"
                elif ((filename.find("lcn海外") >= 0) | (filename.find("LCN海外") >= 0)):
                    df1["SHOPCODE"] = "lcnov"
                    df1["SHOPNAME"] = "lcn海外旗舰店"
                elif filename.find("met海外") >= 0:
                    df1["SHOPCODE"] = "manukascosmetov"
                    df1["SHOPNAME"] = "manukascosmet海外旗舰店"
                elif filename.find("ule海外") >= 0:
                    df1["SHOPCODE"] = "nuggelasuleov"
                    df1["SHOPNAME"] = "nuggelasule海外旗舰店"
                elif filename.find("rai海外") >= 0:
                    df1["SHOPCODE"] = "samouraiov"
                    df1["SHOPNAME"] = "samourai海外旗舰店"
                elif filename.find("ex海外") >= 0:
                    df1["SHOPCODE"] = "ultradexov"
                    df1["SHOPNAME"] = "ultradex海外旗舰店"
                elif ((filename.find("ambra海外") >= 0) | (filename.find("chiara海外") >= 0)):
                    df1["SHOPCODE"] = "chiaraambraov"
                    df1["SHOPNAME"] = "chiaraambra海外旗舰店"
                elif ((filename.find("Fit海外") >= 0) | (filename.find("cora海外") >= 0)):
                    df1["SHOPCODE"] = "dicoraurbanfitov"
                    df1["SHOPNAME"] = "dicoraurbanfit海外旗舰店"
                elif filename.find("iwhite海外") >= 0:
                    df1["SHOPCODE"] = "iwhiteovqjd"
                    df1["SHOPNAME"] = "iwhite海外旗舰店"
                elif ((filename.find("个护海外") >= 0) | (filename.find("mage海外") >= 0)):
                    df1["SHOPCODE"] = "megaggov"
                    df1["SHOPNAME"] = "mega个护海外专营店"
                elif filename.find("护理海外") >= 0:
                    df1["SHOPCODE"] = "megagrhlhwzydov"
                    df1["SHOPNAME"] = "mega个人护理海外专营店"
                elif filename.find("lab海外") >= 0:
                    df1["SHOPCODE"] = "smilelabov"
                    df1["SHOPNAME"] = "smilelab海外旗舰店"
                elif ((filename.find("image海外") >= 0) | (filename.find("IMAGE海外") >= 0) | (
                        filename.find("imega海外") >= 0)):
                    df1["SHOPCODE"] = "swissimageov"
                    df1["SHOPNAME"] = "swissimage海外旗舰店"
                elif filename.find("电器海外") >= 0:
                    df1["SHOPCODE"] = "unixdqov"
                    df1["SHOPNAME"] = "unix电器海外旗舰店"
                elif filename.find("lilac海外") >= 0:
                    df1["SHOPCODE"] = "lilacov"
                    df1["SHOPNAME"] = "lilac海外旗舰店"

                print(df1.head().to_markdown())

            else:
                dict = {"TID": "nan", "SHOPNAME": "", "PLATFORM": "", "SHOPCODE": "", "BILLPLATFORM": "", "CREATED": "",
                        "TITLE": "", "TRADE_TYPE": "", "BUSINESS_NO": "", "INCOME_AMOUNT": "", "EXPEND_AMOUNT": "",
                        "TRADING_CHANNELS": "", "BUSINESS_DESCRIPTION": "", "remark": "", "BUSINESS_BILL_SOURCE": "",
                        "IS_REFUNDAMOUNT": "", "IS_AMOUNT": "", "OID": "", "SOURCEDATA": "", "RECIPROCAL_ACCOUNT": "",
                        "BATCHNO": "", "currency": "", "overseas_income": "", "overseas_expend": "",
                        "currency_cny_rate": ""}
                df = pd.DataFrame(dict, index=[0])
                return df

        else:
            if filename.find("csv") >= 0:
                try:
                    df = pd.read_csv(filename, skiprows=4, dtype=str, encoding="gb18030")
                except Exception as e:
                    try:
                        df = pd.read_csv(filename, skiprows=4, dtype=str)
                    except Exception as e:
                        print("csv文件读取报错！")
                        dict = {"TID": "nan", "SHOPNAME": "", "PLATFORM": "", "SHOPCODE": "", "BILLPLATFORM": "",
                                "CREATED": "",
                                "TITLE": "", "TRADE_TYPE": "", "BUSINESS_NO": "", "INCOME_AMOUNT": "",
                                "EXPEND_AMOUNT": "",
                                "TRADING_CHANNELS": "", "BUSINESS_DESCRIPTION": "", "remark": "",
                                "BUSINESS_BILL_SOURCE": "",
                                "IS_REFUNDAMOUNT": "", "IS_AMOUNT": "", "OID": "", "SOURCEDATA": "",
                                "RECIPROCAL_ACCOUNT": "",
                                "BATCHNO": "", "currency": "", "overseas_income": "", "overseas_expend": "",
                                "currency_cny_rate": ""}
                        df = pd.DataFrame(dict, index=[0])
                        return df
            elif filename.find("xls") >= 0:
                try:
                    df = pd.read_xlsx(filename, skiprows=4, dtype=str)
                except Exception as e:
                    print("xls文件读取报错！")
                    dict = {"TID": "nan", "SHOPNAME": "", "PLATFORM": "", "SHOPCODE": "", "BILLPLATFORM": "",
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
            else:
                print("非excel文件")
                dict = {"TID": "nan", "SHOPNAME": "", "PLATFORM": "", "SHOPCODE": "", "BILLPLATFORM": "",
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
            for column_name in df.columns:
                df.rename(columns={column_name: column_name.replace(" ", "").replace("\n", "").strip()}, inplace=True)
            if "总金额（元）" in df.columns:
                print("跳过此汇总文件！")
                dict = {"TID": "nan", "SHOPNAME": "", "PLATFORM": "", "SHOPCODE": "", "BILLPLATFORM": "",
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
            df["发生时间"] = df["发生时间"].astype("datetime64[ns]")
            df = df.replace("\s+", "", regex=True)
            print(len(df))

            df["账务流水号"] = df["账务流水号"].astype(str)
            df = df[~df["账务流水号"].str.contains("#")]
            # df.dropna(inplace=True)/Users/maclove/Downloads/转换账单数据/2022/天猫/天猫账单/unix旗舰店支付宝202202
            print(len(df))
            print(df.head(5).to_markdown())
            # if "业务基础订单号" in df.columns:
            #     df.rename(columns={"业务基础订单号": "订单号"}, inplace=True)
            # elif "商户订单号" in df.columns:
            #     df["商户订单号"] = df["商户订单号"].apply(lambda x: x.replace("T200P", "").replace("CAE_CHARITY_",""))
            #     df.rename(columns={"商户订单号": "订单号"}, inplace=True)
            if "业务描述" in df.columns:
                df["业务描述"] = df["业务描述"].astype(str)
            # if "商户订单号" in df.columns:
            #     df["商户订单号"] = df["商户订单号"].apply(lambda x:np.nan if len(x)<1 else x)
            print(df.head(5).to_markdown())

            try:
                print(len(df))
                dfa = df[df["商户订单号"].str.contains("T500", na=False) | df["备注"].str.contains("诚e赊买家还款", na=False) | df[
                    "商品名称"].str.contains("诚e赊订单红包", na=False)]
                # dfa = df[df["商户订单号"].str.contains("T500") | df["备注"].str.contains("诚e赊买家还款") | df["商品名称"].str.contains("诚e赊订单红包")]
                print(len(dfa))
                df = df[~df["商户订单号"].str.contains("T500", na=False) & ~df["备注"].str.contains("诚e赊买家还款", na=False) & ~df[
                    "商品名称"].str.contains("诚e赊订单红包", na=False)]
                # df = df[~df["商户订单号"].str.contains("T500") & ~df["备注"].str.contains("诚e赊买家还款") & ~df["商品名称"].str.contains("诚e赊订单红包")]
                print(len(df))
            except Exception as e:
                print("区分阿里巴巴数据报错！")
                dfa = pd.DataFrame()

            if filename.find("淘宝") >= 0:
                plat = "TAOBAO"
            elif filename.find("天猫") >= 0:
                plat = "TMALL"
            df1 = pd.DataFrame()
            df["收入金额（+元）"] = df["收入金额（+元）"].astype(float)
            df["支出金额（-元）"] = df["支出金额（-元）"].astype(float)
            df1["TID"] = ""
            if len(df) == 0:
                pass
            elif "业务基础订单号" in df.columns:
                print("订单号1")
                # df["业务基础订单号"].replace("	", np.nan, inplace=True)
                # df["商户订单号"].replace("	", np.nan, inplace=True)
                # df["备注"].replace("	", np.nan, inplace=True)
                # df["商户订单号"] = df["商户订单号"].astype(str)
                # df["商户订单号"] = df["商户订单号"].apply(lambda x: np.nan if str(x).isspace() else x)
                # df["备注"] = df["备注"].apply(lambda x: np.nan if str(x).isspace() else x)
                # df["业务基础订单号"] = df["业务基础订单号"].astype(str)
                # df["商户订单号"] = df["商户订单号"].astype(str)
                # df["备注"] = df["备注"].astype(str)
                df["业务基础订单号"] = df["业务基础订单号"].apply(lambda x: np.nan if len(str(x)) < 1 else x)
                df["商户订单号"] = df["商户订单号"].apply(lambda x: np.nan if len(str(x)) < 1 else x)
                df["备注"] = df["备注"].apply(lambda x: np.nan if len(str(x)) < 1 else x)
                df["业务基础订单号"] = df["业务基础订单号"].astype(str)
                df["商户订单号"] = df["商户订单号"].astype(str)
                df["备注"] = df["备注"].astype(str)
                print(df.head(1).to_markdown())
                df1["TID"] = df.apply(
                    lambda x: taobao_tid(x["商户订单号"], x["业务类型"], x["备注"], x["业务流水号"]).strip() if x["业务基础订单号"].find(
                        "nan") >= 0 else x["业务基础订单号"], axis=1)
            elif "商户订单号" in df.columns:
                print("订单号2")
                # df["商户订单号"].replace("	", np.nan, inplace=True)
                # df["备注"].replace("	", np.nan, inplace=True)
                # df["商户订单号"] = df["商户订单号"].astype(str)
                # df["商户订单号"] = df["商户订单号"].apply(lambda x: np.nan if str(x).isspace() else x.strip())
                # df["备注"] = df["备注"].apply(lambda x: np.nan if str(x).isspace() else x)

                df["商户订单号"] = df["商户订单号"].apply(lambda x: np.nan if len(str(x)) < 1 else x)
                df["备注"] = df["备注"].apply(lambda x: np.nan if len(str(x)) < 1 else x)
                df["商户订单号"] = df["商户订单号"].astype(str)
                df["备注"] = df["备注"].astype(str)
                df1["TID"] = df.apply(
                    lambda x: x["业务流水号"] if x["商户订单号"].find("nan") >= 0 else taobao_tid(x["商户订单号"], x["业务类型"], x["备注"],
                                                                                        x["业务流水号"]).strip(), axis=1)
            elif "业务流水号" in df.columns:
                print("订单号3")
                df1["TID"] = df["业务流水号"].astype(str).strip()
            df1["SHOPNAME"] = ""
            df1["PLATFORM"] = plat
            df1["SHOPCODE"] = ""
            df1["BILLPLATFORM"] = "ZFB"
            df1["CREATED"] = df["发生时间"]
            df1["TITLE"] = df["商品名称"]
            df1["TRADE_TYPE"] = df["业务类型"]
            df1["BUSINESS_NO"] = df["账务流水号"]
            df1["INCOME_AMOUNT"] = df["收入金额（+元）"]
            df1["EXPEND_AMOUNT"] = df["支出金额（-元）"]
            df1["TRADING_CHANNELS"] = df["交易渠道"]
            df1["BUSINESS_DESCRIPTION"] = ""
            if len(df) == 0:
                pass
            elif "业务描述" in df.columns:
                print("业务描述1")
                # df["业务描述"].replace("nan", np.nan, inplace=True)
                df["业务描述"] = df["业务描述"].apply(lambda x: np.nan if len(str(x)) < 1 else x)
                df["业务描述"] = df["业务描述"].astype(str)
                df["备注"] = df["备注"].astype(str)
                df["商户订单号"] = df["商户订单号"].astype(str)
                df["商品名称"] = df["商品名称"].astype(str)
                df1["BUSINESS_DESCRIPTION"] = df.apply(
                    lambda x: taobao_desc(x["备注"], x["业务类型"], x["商户订单号"], x["商品名称"], x["业务描述"]), axis=1)

            else:
                print("业务描述2")
                df["备注"] = df["备注"].astype(str)
                df["商户订单号"] = df["商户订单号"].astype(str)
                df["商品名称"] = df["商品名称"].astype(str)
                df1["BUSINESS_DESCRIPTION"] = df.apply(
                    lambda x: taobao_desc(x["备注"], x["业务类型"], x["商户订单号"], x["商品名称"], "nan"), axis=1)
            df1["remark"] = df["备注"]
            if "业务账单来源" in df.columns:
                df1["BUSINESS_BILL_SOURCE"] = df["业务账单来源"]
            else:
                df1["BUSINESS_BILL_SOURCE"] = ""
            df1["IS_REFUNDAMOUNT"] = ""
            df1["IS_AMOUNT"] = ""
            if len(df) == 0:
                pass
            elif "业务描述" in df.columns:
                df1["IS_REFUNDAMOUNT"] = df.apply(
                    lambda x: taobao_is_refund(x["业务描述"], x["商品名称"], x["业务类型"], x["备注"], x["支出金额（-元）"], x["商户订单号"],
                                               filename), axis=1)
            else:
                df1["IS_REFUNDAMOUNT"] = df.apply(
                    lambda x: taobao_is_refund("nan", x["商品名称"], x["业务类型"], x["备注"], x["支出金额（-元）"], x["商户订单号"],
                                               filename), axis=1)
            if len(df) == 0:
                pass
            elif "业务描述" in df.columns:
                df1["IS_AMOUNT"] = df.apply(
                    lambda x: taobao_is_amount(x["业务描述"], x["商品名称"], x["业务类型"], x["商户订单号"], x["收入金额（+元）"], filename),
                    axis=1)
            else:
                df1["IS_AMOUNT"] = df.apply(
                    lambda x: taobao_is_amount("nan", x["商品名称"], x["业务类型"], x["商户订单号"], x["收入金额（+元）"], filename),
                    axis=1)
            df1["OID"] = ""
            df1["SOURCEDATA"] = "EXCEL"
            df1["RECIPROCAL_ACCOUNT"] = df["对方账号"]
            df1["BATCHNO"] = ""
            df1["currency"] = ""
            df1["overseas_income"] = ""
            df1["overseas_expend"] = ""
            df1["currency_cny_rate"] = ""
            df1["CREATED"] = df1["CREATED"].astype("datetime64[ns]")
            print(len(df1))
            print(df1.head(5).to_markdown())

            if ((filename.find("dentylactive") >= 0) & (filename.find("海外") < 0)):
                df1["SHOPCODE"] = "dentylactive"
                df1["SHOPNAME"] = "dentylactive旗舰店"
                df1["PLATFORM"] = "TMALL"
            elif filename.find("loshi旗舰店") >= 0:
                df1["SHOPCODE"] = "loshi"
                df1["SHOPNAME"] = "loshi旗舰店"
                df1["PLATFORM"] = "TMALL"
            elif filename.find("dentyl天猫旗舰店") >= 0:
                df1["SHOPCODE"] = "dentylactive"
                df1["SHOPNAME"] = "dentylactive旗舰店"
                df1["PLATFORM"] = "TMALL"
            elif ((filename.find("麦凯莱品牌") >= 0) & (filename.find("特价区") < 0)):
                df1["SHOPCODE"] = "mklppzydzd"
                df1["SHOPNAME"] = "麦凯莱品牌自营店"
                df1["PLATFORM"] = "TAOBAO"
            elif ((filename.find("麦凯莱品牌") >= 0) & (filename.find("特价区") >= 0)):
                df1["SHOPCODE"] = "mklppzydtjq"
                df1["SHOPNAME"] = "麦凯莱品牌自营店特价区"
                df1["PLATFORM"] = "TMALL"
            elif filename.find("优皓康") >= 0:
                df1["SHOPCODE"] = "yhk"
                df1["SHOPNAME"] = "优皓康旗舰店"
                df1["PLATFORM"] = "TMALL"
            elif filename.find("unix旗舰店") >= 0:
                df1["SHOPCODE"] = "unix"
                df1["SHOPNAME"] = "unix旗舰店"
                df1["PLATFORM"] = "TMALL"
            elif filename.find("bodyaid旗舰店") >= 0:
                df1["SHOPCODE"] = "bodyaid"
                df1["SHOPNAME"] = "bodyaid旗舰店"
                df1["PLATFORM"] = "TMALL"
            elif filename.find("allnaturaladvice旗舰店") >= 0:
                df1["SHOPCODE"] = "allnaturaladvice"
                df1["SHOPNAME"] = "allnaturaladvice旗舰店"
                df1["PLATFORM"] = "TMALL"
            elif filename.find("莎莎") >= 0:
                df1["SHOPCODE"] = "ssmzppd"
                df1["SHOPNAME"] = "莎莎美妆品牌店"
                df1["PLATFORM"] = "TAOBAO"
            elif filename.find("黑天鹅美妆") >= 0:
                df1["SHOPCODE"] = "hte"
                df1["SHOPNAME"] = "黑天鹅美妆"
                df1["PLATFORM"] = "TAOBAO"
            elif filename.find("芭莎") >= 0:
                df1["SHOPCODE"] = "bsmzg"
                df1["SHOPNAME"] = "时尚芭莎美妆店主"
                df1["PLATFORM"] = "TAOBAO"
            elif filename.find("逍遥") >= 0:
                df1["SHOPCODE"] = "xyhgkjg"
                df1["SHOPNAME"] = "逍遥海外跨境购"
                df1["PLATFORM"] = "TAOBAO"
            elif filename.find("茜茜") >= 0:
                df1["SHOPCODE"] = "cecixxmz"
                df1["SHOPNAME"] = "Ceci茜茜美妆"
                df1["PLATFORM"] = "TAOBAO"
            elif filename.find("惠优购官方") >= 0:
                df1["SHOPCODE"] = "yhg"
                df1["SHOPNAME"] = "惠优购官方旗舰店"
                df1["PLATFORM"] = "TMALL"
            elif filename.find("vitabloom") >= 0:
                df1["SHOPCODE"] = "vitabloom"
                df1["SHOPNAME"] = "vitabloom美妆店"
                df1["PLATFORM"] = "TMALL"
            elif filename.find("samourai美妆店") >= 0:
                df1["SHOPCODE"] = "samouraimzd"
                df1["SHOPNAME"] = "samourai美妆店"
                df1["PLATFORM"] = "TMALL"
            elif filename.find("惠优购企业店") >= 0:
                df1["SHOPCODE"] = "hyg"
                df1["SHOPNAME"] = "惠优购企业店"
                df1["PLATFORM"] = "TAOBAO"
            elif filename.find("Morei品牌") >= 0:
                df1["SHOPCODE"] = "moreippd"
                df1["SHOPNAME"] = "淘宝Morei品牌店"
                df1["PLATFORM"] = "TAOBAO"
            elif filename.find("樱语品牌店") >= 0:
                df1["SHOPCODE"] = "yyppd"
                df1["SHOPNAME"] = "樱语品牌店"
                df1["PLATFORM"] = "TAOBAO"
            elif filename.find("珍芯美妆") >= 0:
                df1["SHOPCODE"] = "zxmz"
                df1["SHOPNAME"] = "珍芯美妆的小店"
                df1["PLATFORM"] = "TAOBAO"
            elif filename.find("贝贝港湾") >= 0:
                df1["SHOPCODE"] = "bbgw"
                df1["SHOPNAME"] = "贝贝港湾的小店"
                df1["PLATFORM"] = "TAOBAO"
            elif filename.find("樱语品牌店") >= 0:
                df1["SHOPCODE"] = "yyppd"
                df1["SHOPNAME"] = "樱语品牌店"
                df1["PLATFORM"] = "TAOBAO"
            elif filename.find("肌密泉") >= 0:
                df1["SHOPCODE"] = "jmq"
                df1["SHOPNAME"] = "肌密泉"
                df1["PLATFORM"] = "TAOBAO"
            elif filename.find("柚选美妆店") >= 0:
                df1["SHOPCODE"] = "yxmz"
                df1["SHOPNAME"] = "柚选美妆店"
                df1["PLATFORM"] = "TAOBAO"
            elif filename.find("盈养泉旗舰店") >= 0:
                df1["SHOPCODE"] = "yyqqjd2021"
                df1["SHOPNAME"] = "盈养泉旗舰店"
                df1["PLATFORM"] = "TMALL"
            elif filename.find("montooth旗舰店") >= 0:
                df1["SHOPCODE"] = "montoothqjd"
                df1["SHOPNAME"] = "montooth旗舰店"
                df1["PLATFORM"] = "TMALL"
            elif filename.find("EC个人护理旗舰店") >= 0:
                df1["SHOPCODE"] = "ecgrhljjd"
                df1["SHOPNAME"] = "EC个人护理旗舰店"
                df1["PLATFORM"] = "TMALL"
            elif filename.find("mades天猫旗舰店") >= 0:
                df1["SHOPCODE"] = "mades"
                df1["SHOPNAME"] = "mades旗舰店"
                df1["PLATFORM"] = "TMALL"
            elif filename.find("magicsymbol美妆店") >= 0:
                df1["SHOPCODE"] = "magicsymbol"
                df1["SHOPNAME"] = "magicsymbol美妆店"
                df1["PLATFORM"] = "TMALL"
            elif filename.find("montooth品牌店") >= 0:
                df1["SHOPCODE"] = "montooth"
                df1["SHOPNAME"] = "montooth品牌店"
                df1["PLATFORM"] = "TMALL"
            elif filename.find("樱语洗护旗舰店") >= 0:
                df1["SHOPCODE"] = "yyxhjjd"
                df1["SHOPNAME"] = "樱语洗护旗舰店"
                df1["PLATFORM"] = "TMALL"
            elif filename.find("morei旗舰店") >= 0:
                df1["SHOPCODE"] = "morei"
                df1["SHOPNAME"] = "morei家清店"
                df1["PLATFORM"] = "TMALL"
            elif filename.find("tb6539190105") >= 0:
                df1["SHOPCODE"] = "tb6539190105"
                df1["SHOPNAME"] = "tb6539190105"
                df1["PLATFORM"] = "TAOBAO"
            elif filename.find("tb4557506560") >= 0:
                df1["SHOPCODE"] = "tbswwqwllwll"
                df1["SHOPNAME"] = "tb4557506560"
                df1["PLATFORM"] = "TAOBAO"
            elif filename.find("樱加美旗舰店") >= 0:
                df1["SHOPCODE"] = "yjmqjd"
                df1["SHOPNAME"] = "樱加美旗舰店"
                df1["PLATFORM"] = "TMALL"
            elif filename.find("cain旗舰店") >= 0:
                df1["SHOPCODE"] = "cain"
                df1["SHOPNAME"] = "cain旗舰店"
                df1["PLATFORM"] = "TMALL"
            elif filename.find("milelab旗舰店") >= 0:
                df1["SHOPCODE"] = "smilelab"
                df1["SHOPNAME"] = "smilelab旗舰店"
                df1["PLATFORM"] = "TMALL"
            elif filename.find("米兰站美妆") >= 0:
                df1["SHOPCODE"] = "mlzmz"
                df1["SHOPNAME"] = "米兰站美妆"
                df1["PLATFORM"] = "TAOBAO"
            elif filename.find("前男友美妆") >= 0:
                df1["SHOPCODE"] = "qnymz"
                df1["SHOPNAME"] = "前男友美妆"
                df1["PLATFORM"] = "TAOBAO"
            elif filename.find("咖思美旗舰店") >= 0:
                df1["SHOPCODE"] = "jsmqjd"
                df1["SHOPNAME"] = "咖思美旗舰店"
                df1["PLATFORM"] = "TMALL"
            elif filename.find("77897我") >= 0:
                df1["SHOPCODE"] = "qqbjqw"
                df1["SHOPNAME"] = "77897我"
                df1["PLATFORM"] = "TAOBAO"
            elif filename.find("BODI美妆馆") >= 0:
                df1["SHOPCODE"] = "bodi"
                df1["SHOPNAME"] = "BODI美妆馆"
                df1["PLATFORM"] = "TAOBAO"
            elif filename.find("tb1287472841") >= 0:
                df1["SHOPCODE"] = "tbyebqsqebsy"
                df1["SHOPNAME"] = "tb1287472841"
                df1["PLATFORM"] = "TAOBAO"
            elif filename.find("tb520670794") >= 0:
                df1["SHOPCODE"] = "tbwellqlqjs"
                df1["SHOPNAME"] = "tb520670794"
                df1["PLATFORM"] = "TAOBAO"
            elif filename.find("tb5364335213") >= 0:
                df1["SHOPCODE"] = "tbwslsssweys"
                df1["SHOPNAME"] = "tb5364335213"
                df1["PLATFORM"] = "TAOBAO"
            elif filename.find("橘子日杂店") >= 0:
                df1["SHOPCODE"] = "jzrzd"
                df1["SHOPNAME"] = "橘子日杂店"
                df1["PLATFORM"] = "TAOBAO"
            elif filename.find("泡研美妆") >= 0:
                df1["SHOPCODE"] = "pymz"
                df1["SHOPNAME"] = "泡研美妆"
                df1["PLATFORM"] = "TAOBAO"
            elif filename.find("橙子新建站") >= 0:
                df1["SHOPCODE"] = "jzxjz"
                df1["SHOPNAME"] = "橙子新建站"
                df1["PLATFORM"] = "TAOBAO"
            elif filename.find("魔法符号美妆店") >= 0:
                df1["SHOPCODE"] = "mffhmzd"
                df1["SHOPNAME"] = "魔法符号美妆店"
                df1["PLATFORM"] = "TAOBAO"
            elif filename.find("若蘅") >= 0:
                df1["SHOPCODE"] = "rhmzd"
                df1["SHOPNAME"] = "若蘅旗舰店"
                df1["PLATFORM"] = "TAOBAO"
            elif filename.find("swissimage大贸旗舰店") >= 0:
                df1["SHOPCODE"] = "swissimage"
                df1["SHOPNAME"] = "swissimage大贸旗舰店"
                df1["PLATFORM"] = "TMALL"
            elif filename.find("tb380231166") >= 0:
                df1["SHOPCODE"] = "tb380231166"
                df1["SHOPNAME"] = "tb380231166"
                df1["PLATFORM"] = "TAOBAO"
            elif filename.find("宅星人旗舰店") >= 0:
                df1["SHOPCODE"] = "zxr"
                df1["SHOPNAME"] = "宅星人旗舰店"
                df1["PLATFORM"] = "TMALL"
            elif filename.find("drbubble旗舰店") >= 0:
                df1["SHOPCODE"] = "drbubbleqjd"
                df1["SHOPNAME"] = "drbubble旗舰店"
                df1["PLATFORM"] = "TMALL"
            elif filename.find("tb0563814991") >= 0:
                df1["SHOPCODE"] = "tb0563814991"
                df1["SHOPNAME"] = "tb0563814991"
                df1["PLATFORM"] = "TAOBAO"
            elif filename.find("tb840130912") >= 0:
                df1["SHOPCODE"] = "tb840130912"
                df1["SHOPNAME"] = "tb840130912"
                df1["PLATFORM"] = "TAOBAO"
            elif filename.find("tb8999331946") >= 0:
                df1["SHOPCODE"] = "tbbjjjssyjsl"
                df1["SHOPNAME"] = "tb8999331946"
                df1["PLATFORM"] = "TAOBAO"
            elif filename.find("麦凯莱厂家直销店") >= 0:
                df1["SHOPCODE"] = "mklcjzxd"
                df1["SHOPNAME"] = "麦凯莱厂家直销店"
                df1["PLATFORM"] = "TAOBAO"
            elif filename.find("倚新妆全球购") >= 0:
                df1["SHOPCODE"] = "mklppzydzd"
                df1["SHOPNAME"] = "倚新妆全球购"
                df1["PLATFORM"] = "TAOBAO"
            elif filename.find("星蓝个护") >= 0:
                df1["SHOPCODE"] = "xlghd"
                df1["SHOPNAME"] = "星蓝个护店"
                df1["PLATFORM"] = "TAOBAO"
            elif filename.find("玛雅个护美妆") >= 0:
                df1["SHOPCODE"] = "myghmz"
                df1["SHOPNAME"] = "玛雅个护美妆店"
                df1["PLATFORM"] = "TAOBAO"
            elif filename.find("柚选美妆") >= 0:
                df1["SHOPCODE"] = "yxmz"
                df1["SHOPNAME"] = "柚选美妆店"
                df1["PLATFORM"] = "TAOBAO"
            elif filename.find("惠优购") >= 0:
                df1["SHOPCODE"] = "hyg"
                df1["SHOPNAME"] = "淘客惠优购"
                df1["PLATFORM"] = "TAOBAO"
            elif filename.find("尚美个护") >= 0:
                df1["SHOPCODE"] = "smgh"
                df1["SHOPNAME"] = "尚美个护"
                df1["PLATFORM"] = "TAOBAO"
            elif filename.find("tb862892834") >= 0:
                df1["SHOPCODE"] = "tbblebjebss"
                df1["SHOPNAME"] = "航星洗护店（tb862892834）"
                df1["PLATFORM"] = "TAOBAO"
            elif filename.find("tb584689610") >= 0:
                df1["SHOPCODE"] = "tbwbslbjlyl"
                df1["SHOPNAME"] = "盛美日化品店（tb584689610）"
                df1["PLATFORM"] = "TAOBAO"
            elif filename.find("tb547024599") >= 0:
                df1["SHOPCODE"] = "tbwsqleswjj"
                df1["SHOPNAME"] = "美丽蓓蕾（tb547024599）"
                df1["PLATFORM"] = "TAOBAO"
            elif filename.find("tb203703983") >= 0:
                df1["SHOPCODE"] = "tbelsqlsjbs"
                df1["SHOPNAME"] = "卖家优选直销店（tb203703983）"
                df1["PLATFORM"] = "TAOBAO"
            elif filename.find("tb0884087963") >= 0:
                df1["SHOPCODE"] = "tblbbslbqjls"
                df1["SHOPNAME"] = "BodyAid博滴品牌严选店（tb0884087963）"
                df1["PLATFORM"] = "TAOBAO"

            df1["INCOME_AMOUNT"] = df1["INCOME_AMOUNT"].astype(float).abs()
            df1["EXPEND_AMOUNT"] = df1["EXPEND_AMOUNT"].astype(float)
            print(len(df1))
            print(df1.head(5).to_markdown())
            print("天猫")

            if len(dfa) > 0:
                plat = "ALIBABA"
                df2 = pd.DataFrame()
                dfa["收入金额（+元）"] = dfa["收入金额（+元）"].astype(float)
                dfa["支出金额（-元）"] = dfa["支出金额（-元）"].astype(float)
                if "业务基础订单号" in dfa.columns:
                    print("订单号1")
                    # dfa["业务基础订单号"].replace("	", np.nan, inplace=True)
                    # dfa["商户订单号"].replace("	", np.nan, inplace=True)
                    # dfa["备注"].replace("	", np.nan, inplace=True)
                    # dfa["商户订单号"] = dfa["商户订单号"].astype(str)
                    # dfa["商户订单号"] = dfa["商户订单号"].apply(lambda x: np.nan if str(x).isspace() else x)
                    # dfa["备注"] = dfa["备注"].apply(lambda x: np.nan if str(x).isspace() else x)
                    dfa["业务基础订单号"] = dfa["业务基础订单号"].apply(lambda x: np.nan if len(str(x)) < 1 else x)
                    dfa["商户订单号"] = dfa["商户订单号"].apply(lambda x: np.nan if len(str(x)) < 1 else x)
                    dfa["备注"] = dfa["备注"].apply(lambda x: np.nan if len(str(x)) < 1 else x)
                    dfa["业务基础订单号"] = dfa["业务基础订单号"].astype(str)
                    dfa["商户订单号"] = dfa["商户订单号"].astype(str)
                    dfa["备注"] = dfa["备注"].astype(str)
                    print("定位")
                    print(dfa.head().to_markdown())
                    df2["TID"] = dfa.apply(
                        lambda x: taobao_tid(x["商户订单号"], x["业务类型"], x["备注"], x["业务流水号"]).strip() if x["业务基础订单号"].find(
                            "nan") >= 0 else x["业务基础订单号"], axis=1)
                elif "商户订单号" in dfa.columns:
                    print("订单号2")
                    # dfa["商户订单号"].replace("	", np.nan, inplace=True)
                    # dfa["备注"].replace("	", np.nan, inplace=True)
                    # dfa["商户订单号"] = dfa["商户订单号"].astype(str)
                    # dfa["商户订单号"] = dfa["商户订单号"].apply(lambda x: np.nan if str(x).isspace() else x.strip())
                    # dfa["备注"] = dfa["备注"].apply(lambda x: np.nan if str(x).isspace() else x)
                    dfa["商户订单号"] = dfa["商户订单号"].apply(lambda x: np.nan if len(str(x)) < 1 else x)
                    dfa["备注"] = dfa["备注"].apply(lambda x: np.nan if len(str(x)) < 1 else x)
                    dfa["商户订单号"] = dfa["商户订单号"].astype(str)
                    dfa["备注"] = dfa["备注"].astype(str)
                    df2["TID"] = dfa.apply(
                        lambda x: x["业务流水号"] if x["商户订单号"].find("nan") >= 0 else taobao_tid(x["商户订单号"], x["业务类型"],
                                                                                            x["备注"],
                                                                                            x["业务流水号"]).strip(), axis=1)
                elif "业务流水号" in dfa.columns:
                    print("订单号3")
                    df2["TID"] = dfa["业务流水号"].strip()
                # df2["TID"] = df2["TID"].apply(
                #     lambda x: "".join(x.split("-")[-1:]) if x.find("-") >= 0 else "".join(x.split("P")[-1:]))
                df2["SHOPNAME"] = "阿里巴巴卖家联合loshi总代店"
                df2["PLATFORM"] = plat
                df2["SHOPCODE"] = "mjlhzdd"
                df2["BILLPLATFORM"] = "ZFB"
                df2["CREATED"] = dfa["发生时间"]
                df2["TITLE"] = dfa["商品名称"]
                df2["TRADE_TYPE"] = dfa["业务类型"]
                df2["BUSINESS_NO"] = dfa["账务流水号"]
                df2["INCOME_AMOUNT"] = dfa["收入金额（+元）"]
                df2["EXPEND_AMOUNT"] = dfa["支出金额（-元）"]
                df2["TRADING_CHANNELS"] = dfa["交易渠道"]
                if "业务描述" in dfa.columns:
                    print("业务描述1")
                    # dfa["业务描述"].replace("nan", np.nan, inplace=True)
                    dfa["业务描述"] = dfa["业务描述"].apply(lambda x: np.nan if len(str(x)) < 1 else x)
                    dfa["业务描述"] = dfa["业务描述"].astype(str)
                    dfa["备注"] = dfa["备注"].astype(str)
                    dfa["商户订单号"] = dfa["商户订单号"].astype(str)
                    dfa["商品名称"] = dfa["商品名称"].astype(str)
                    df2["BUSINESS_DESCRIPTION"] = dfa.apply(
                        lambda x: taobao_desc(x["备注"], x["业务类型"], x["商户订单号"], x["商品名称"], x["业务描述"]),
                        axis=1)
                else:
                    print("业务描述2")
                    dfa["备注"] = dfa["备注"].astype(str)
                    dfa["商户订单号"] = dfa["商户订单号"].astype(str)
                    dfa["商品名称"] = dfa["商品名称"].astype(str)
                    df2["BUSINESS_DESCRIPTION"] = dfa.apply(
                        lambda x: taobao_desc(x["备注"], x["业务类型"], x["商户订单号"], x["商品名称"], "nan"), axis=1)
                df2["remark"] = dfa["备注"]
                if "业务账单来源" in dfa.columns:
                    df2["BUSINESS_BILL_SOURCE"] = dfa["业务账单来源"]
                else:
                    df2["BUSINESS_BILL_SOURCE"] = ""
                if "业务描述" in df.columns:
                    df2["IS_REFUNDAMOUNT"] = dfa.apply(
                        lambda x: taobao_is_refund(x["业务描述"], x["商品名称"], x["业务类型"], x["备注"], x["支出金额（-元）"], x["商户订单号"],
                                                   filename), axis=1)
                else:
                    df2["IS_REFUNDAMOUNT"] = dfa.apply(
                        lambda x: taobao_is_refund("nan", x["商品名称"], x["业务类型"], x["备注"], x["支出金额（-元）"], x["商户订单号"],
                                                   filename), axis=1)
                if "业务描述" in df.columns:
                    df2["IS_AMOUNT"] = dfa.apply(
                        lambda x: alibaba_is_amount(x["业务描述"], x["商品名称"], x["业务类型"], x["商户订单号"], x["收入金额（+元）"], x["备注"],
                                                    filename), axis=1)
                else:
                    df2["IS_AMOUNT"] = dfa.apply(
                        lambda x: alibaba_is_amount("nan", x["商品名称"], x["业务类型"], x["商户订单号"], x["收入金额（+元）"], x["备注"],
                                                    filename), axis=1)
                df2["OID"] = ""
                df2["SOURCEDATA"] = "EXCEL"
                df2["RECIPROCAL_ACCOUNT"] = dfa["对方账号"]
                df2["BATCHNO"] = ""
                df2["currency"] = ""
                df2["overseas_income"] = ""
                df2["overseas_expend"] = ""
                df2["currency_cny_rate"] = ""
                df2["CREATED"] = df2["CREATED"].astype("datetime64[ns]")
                print(len(df2))
                print(df2.head(5).to_markdown())
                print("阿里巴巴")

                if filename.find("tb6539190105") >= 0:
                    df2["SHOPCODE"] = "zzc"
                    df2["SHOPNAME"] = "植之璨（深圳）化妆品有限公司"
                elif filename.find("tb584689610") >= 0:
                    df2["SHOPCODE"] = "mjc"
                    df2["SHOPNAME"] = "萌洁齿（深圳）日用品有限公司"
                elif filename.find("tb8999331946") >= 0:
                    df2["SHOPCODE"] = "mjyx"
                    df2["SHOPNAME"] = "深圳市卖家优选实业有限公司"
                elif filename.find("tb5364335213") >= 0:
                    df2["SHOPCODE"] = "dr"
                    df2["SHOPNAME"] = "多瑞(深圳)日用品有限公司"
                elif filename.find("77897我") >= 0:
                    df2["SHOPCODE"] = "ylhfp"
                    df2["SHOPNAME"] = "深圳樱岚护肤品有限公司"

                dfs = [df1, df2]
                df1 = pd.concat(dfs)
                print("合并天猫，阿里巴巴账单")
            else:
                pass
            df1 = df1.loc[~((df1.INCOME_AMOUNT == 0) & (df1.EXPEND_AMOUNT == 0))]
            print("天猫")
            print(len(df1))
            print(df1.head(5).to_markdown())

    # 京东逻辑
    elif filename.find("京东") >= 0:
        # if (((filename.find("邓特") >= 0) | (filename.find("Dentyl") >= 0)) & (filename.find("2021") < 0)) :
        #     if filename.find("结算单") >=0:
        #         df = pd.read_excel(filename, dtype=str)
        #         print("邓特结算单")
        #         time = "".join(filename.split("结算单")[1:])
        #         time = "20"+"".join(time.split(".")[:1])
        #         print(time)
        #     else:
        #         df = pd.read_excel(filename, sheet_name="月账单", dtype=str)
        #         print("邓特月账单")
        #         time = "".join(filename.split(".")[:-1])
        #         time = "".join(time.split("自营店")[-1:])
        #         print(time)
        #     for column_name in df.columns:
        #         df.rename(columns={column_name: column_name.replace(" ", "").replace("\n", "").strip()}, inplace=True)
        #     print(df.head(1).to_markdown())
        #     df.dropna(subset=["订单编号"], inplace=True)
        #     df = df[~df["订单编号"].str.contains("nan")]
        #     df["金额"] = df["金额"].astype(float)
        #     # df["完成时间"] = df["完成时间"].astype(str)
        #     plat = "JD"
        #     df1 = pd.DataFrame()
        #     df1["TID"] = df["订单编号"]
        #     df1["SHOPNAME"] = "dentylactive京东自营旗舰店"
        #     df1["PLATFORM"] = plat
        #     df1["SHOPCODE"] = "dentylactivejdzyd"
        #     df1["BILLPLATFORM"] = plat
        #     print("定位1")
        #     df1["CREATED"] = jd_dentyl(time)
        #     print("定位2")
        #     if "商品名称" in df.columns:
        #         df1["TITLE"] = df["商品名称"]
        #     else:
        #         df1["TITLE"] = ""
        #     df1["TRADE_TYPE"] = df["费用项"]
        #     if "商品编号" in df.columns:
        #         df1["BUSINESS_NO"] = df["商品编号"] + df["完成时间"]
        #     elif "商品编码" in df.columns:
        #         df1["BUSINESS_NO"] = df["商品编码"] + df["完成时间"]
        #     else:
        #         df1["BUSINESS_NO"] = ""
        #     df1["INCOME_AMOUNT"] = df["金额"].apply(lambda x: x if x > 0 else 0)
        #     df1["EXPEND_AMOUNT"] = df["金额"].apply(lambda x: x if x < 0 else 0)
        #     df1["TRADING_CHANNELS"] = ""
        #     df1["BUSINESS_DESCRIPTION"] = df["费用项"]
        #     df1["remark"] = "原账单文件时间：" + time
        #     df1["BUSINESS_BILL_SOURCE"] = ""
        #     # df1["IS_REFUNDAMOUNT"] = df.apply(lambda x:1 if ((x["费用项"]=="FCS退款")|((x["费用项"]=="FCS货款")&(x["金额"]<0))|((x["费用项"]=="代收配送费")&(x["金额"]<0))) else 0,axis=1)
        #     # df1["IS_AMOUNT"] = df.apply(lambda x:1 if (((x["费用项"]=="FCS货款")&(x["金额"]>0))|((x["费用项"]=="代收配送费")&(x["金额"]>0))) else 0,axis=1)
        #     df1["IS_REFUNDAMOUNT"] = df.apply(lambda x: 1 if x["金额"] < 0 else 0, axis=1)
        #     df1["IS_AMOUNT"] = df.apply(lambda x: 1 if x["金额"] > 0 else 0,axis=1)
        #     df1["OID"] = ""
        #     df1["SOURCEDATA"] = "EXCEL"
        #     df1["RECIPROCAL_ACCOUNT"] = ""
        #     df1["BATCHNO"] = ""
        #     df1["currency"] = ""
        #     df1["overseas_income"] = ""
        #     df1["overseas_expend"] = ""
        #     df1["currency_cny_rate"] = ""
        #     df1["INCOME_AMOUNT"] = df1["INCOME_AMOUNT"].astype(float)
        #     df1["EXPEND_AMOUNT"] = df1["EXPEND_AMOUNT"].astype(float)
        #     df1["CREATED"] = df1["CREATED"].astype("datetime64[ns]")
        #     # df1["iyear"] = df1["CREATED"].apply(lambda x: x.year)
        #     # df1["imonth"] = df1["CREATED"].apply(lambda x: x.month)
        #     print(df1.head(5).to_markdown())
        #
        #     return df1
        # elif ((filename.find("VitaBloom") >= 0) & (filename.find("2021") >= 0)):
        #     if ((filename.find("核销清单") >= 0)&(filename.find("xls") >= 0)):
        #         df = pd.read_excel(filename, dtype=str)
        #         print("VitaBloom京东-xls")
        #         time = "".join(filename.split("核销清单")[1:])
        #         time = "".join(time.split(".")[:1])
        #         print(time)
        #     elif ((filename.find("核销清单") >= 0)&(filename.find("csv") >= 0)):
        #         df = pd.read_csv(filename, dtype=str, encoding="gb18030")
        #         print("VitaBloom京东-csv")
        #         time = "".join(filename.split("核销清单")[1:])
        #         time = "".join(time.split(".")[:1])
        #         print(time)
        #     else:
        #         dict = {"TID": "nan", "SHOPNAME": "", "PLATFORM": "", "SHOPCODE": "", "BILLPLATFORM": "", "CREATED": "",
        #                 "TITLE": "", "TRADE_TYPE": "", "BUSINESS_NO": "", "INCOME_AMOUNT": "", "EXPEND_AMOUNT": "",
        #                 "TRADING_CHANNELS": "", "BUSINESS_DESCRIPTION": "", "remark": "", "BUSINESS_BILL_SOURCE": "",
        #                 "IS_REFUNDAMOUNT": "", "IS_AMOUNT": "", "OID": "", "SOURCEDATA": "", "RECIPROCAL_ACCOUNT": "",
        #                 "BATCHNO": ""}
        #         df = pd.DataFrame(dict, index=[0])
        #         return df
        #     for column_name in df.columns:
        #         df.rename(columns={column_name: column_name.replace(" ", "").replace("\n", "").strip()}, inplace=True)
        #     print(df.head(1).to_markdown())
        #     df.dropna(subset=["订单编号"], inplace=True)
        #     df = df[~df["订单编号"].str.contains("nan")]
        #     df["金额"] = df["金额"].astype(float)
        #     # df["完成时间"] = df["完成时间"].astype(str)
        #     plat = "JD"
        #     df1 = pd.DataFrame()
        #     df1["TID"] = df["订单编号"]
        #     df1["SHOPNAME"] = "VitaBloom京东官方自营旗舰店"
        #     df1["PLATFORM"] = plat
        #     df1["SHOPCODE"] = "vbqjd"
        #     df1["BILLPLATFORM"] = plat
        #     print("定位1")
        #     if ((time.find("04")>=0)|(time.find("06")>=0)):
        #         df1["CREATED"] = time + "30 23:59:59"
        #         df1["CREATED"] = df1["CREATED"].astype("datetime64[ns]")
        #     elif ((time.find("05")>=0)|(time.find("08")>=0)):
        #         df1["CREATED"] = time + "31 23:59:59"
        #         df1["CREATED"] = df1["CREATED"].astype("datetime64[ns]")
        #     print("定位2")
        #     if "商品名称" in df.columns:
        #         df1["TITLE"] = df["商品名称"]
        #     else:
        #         df1["TITLE"] = ""
        #     df1["TRADE_TYPE"] = df["费用项"]
        #     if "商品编号" in df.columns:
        #         df1["BUSINESS_NO"] = df["商品编号"] + df["完成时间"]
        #     elif "商品编码" in df.columns:
        #         df1["BUSINESS_NO"] = df["商品编码"] + df["完成时间"]
        #     else:
        #         df1["BUSINESS_NO"] = ""
        #     df1["INCOME_AMOUNT"] = df["金额"].apply(lambda x: x if x > 0 else 0)
        #     df1["EXPEND_AMOUNT"] = df["金额"].apply(lambda x: x if x < 0 else 0)
        #     df1["TRADING_CHANNELS"] = ""
        #     df1["BUSINESS_DESCRIPTION"] = df["费用项"]
        #     df1["remark"] = "原账单文件时间：" + time
        #     df1["BUSINESS_BILL_SOURCE"] = ""
        #     # df1["IS_REFUNDAMOUNT"] = df.apply(lambda x:1 if ((x["费用项"]=="FCS退款")|((x["费用项"]=="FCS货款")&(x["金额"]<0))|((x["费用项"]=="代收配送费")&(x["金额"]<0))) else 0,axis=1)
        #     # df1["IS_AMOUNT"] = df.apply(lambda x:1 if (((x["费用项"]=="FCS货款")&(x["金额"]>0))|((x["费用项"]=="代收配送费")&(x["金额"]>0))) else 0,axis=1)
        #     df1["IS_REFUNDAMOUNT"] = df.apply(lambda x: 1 if x["金额"] < 0 else 0, axis=1)
        #     df1["IS_AMOUNT"] = df.apply(lambda x: 1 if x["金额"] > 0 else 0, axis=1)
        #     df1["OID"] = ""
        #     df1["SOURCEDATA"] = "EXCEL"
        #     df1["RECIPROCAL_ACCOUNT"] = ""
        #     df1["BATCHNO"] = ""
        #     df1["currency"] = ""
        #     df1["overseas_income"] = ""
        #     df1["overseas_expend"] = ""
        #     df1["currency_cny_rate"] = ""
        #     df1["INCOME_AMOUNT"] = df1["INCOME_AMOUNT"].astype(float)
        #     df1["EXPEND_AMOUNT"] = df1["EXPEND_AMOUNT"].astype(float)
        #     df1["CREATED"] = df1["CREATED"].astype("datetime64[ns]")
        #     # df1["iyear"] = df1["CREATED"].apply(lambda x: x.year)
        #     # df1["imonth"] = df1["CREATED"].apply(lambda x: x.month)
        #     print(df1.head(5).to_markdown())
        # else:
        print("非邓特账单")
        if filename.find("xls") >= 0:
            try:
                df = pd.read_excel(filename, sheet_name="月账单", dtype=str)
            except Exception as e:
                try:
                    df = pd.read_excel(filename, dtype=str)
                    if "金额" not in df.columns:
                        dict = {"TID": "nan", "SHOPNAME": "", "PLATFORM": "", "SHOPCODE": "", "BILLPLATFORM": "",
                                "CREATED": "",
                                "TITLE": "", "TRADE_TYPE": "", "BUSINESS_NO": "", "INCOME_AMOUNT": "",
                                "EXPEND_AMOUNT": "",
                                "TRADING_CHANNELS": "", "BUSINESS_DESCRIPTION": "", "remark": "",
                                "BUSINESS_BILL_SOURCE": "",
                                "IS_REFUNDAMOUNT": "", "IS_AMOUNT": "", "OID": "", "SOURCEDATA": "",
                                "RECIPROCAL_ACCOUNT": "", "BATCHNO": ""}
                        df = pd.DataFrame(dict, index=[0])
                        return df
                except Exception as e:
                    dict = {"TID": "nan", "SHOPNAME": "", "PLATFORM": "", "SHOPCODE": "", "BILLPLATFORM": "",
                            "CREATED": "",
                            "TITLE": "", "TRADE_TYPE": "", "BUSINESS_NO": "", "INCOME_AMOUNT": "",
                            "EXPEND_AMOUNT": "",
                            "TRADING_CHANNELS": "", "BUSINESS_DESCRIPTION": "", "remark": "",
                            "BUSINESS_BILL_SOURCE": "",
                            "IS_REFUNDAMOUNT": "", "IS_AMOUNT": "", "OID": "", "SOURCEDATA": "",
                            "RECIPROCAL_ACCOUNT": "", "BATCHNO": ""}
                    df = pd.DataFrame(dict, index=[0])
                    return df
        elif filename.find("csv") >= 0:
            try:
                df = pd.read_csv(filename, dtype=str, encoding="gb18030")
            except Exception as e:
                try:
                    df = pd.read_csv(filename, dtype=str)
                except Exception as e:
                    dict = {"TID": "nan", "SHOPNAME": "", "PLATFORM": "", "SHOPCODE": "", "BILLPLATFORM": "",
                            "CREATED": "",
                            "TITLE": "", "TRADE_TYPE": "", "BUSINESS_NO": "", "INCOME_AMOUNT": "", "EXPEND_AMOUNT": "",
                            "TRADING_CHANNELS": "", "BUSINESS_DESCRIPTION": "", "remark": "",
                            "BUSINESS_BILL_SOURCE": "",
                            "IS_REFUNDAMOUNT": "", "IS_AMOUNT": "", "OID": "", "SOURCEDATA": "",
                            "RECIPROCAL_ACCOUNT": "", "BATCHNO": ""}
                    df = pd.DataFrame(dict, index=[0])
                    return df
        else:
            dict = {"TID": "nan", "SHOPNAME": "", "PLATFORM": "", "SHOPCODE": "", "BILLPLATFORM": "", "CREATED": "",
                    "TITLE": "", "TRADE_TYPE": "", "BUSINESS_NO": "", "INCOME_AMOUNT": "", "EXPEND_AMOUNT": "",
                    "TRADING_CHANNELS": "", "BUSINESS_DESCRIPTION": "", "remark": "", "BUSINESS_BILL_SOURCE": "",
                    "IS_REFUNDAMOUNT": "", "IS_AMOUNT": "", "OID": "", "SOURCEDATA": "", "RECIPROCAL_ACCOUNT": "",
                    "BATCHNO": ""}
            df = pd.DataFrame(dict, index=[0])
            return df
        for column_name in df.columns:
            df.rename(columns={column_name: column_name.replace(" ", "").replace("\n", "").strip()}, inplace=True)
        # if (("订单号" in df.columns)&("订单编号" not in df.columns)):
        #     df.rename(columns={"订单号":"订单编号","应付金额":"金额"},inplace=True)
        df = df.replace('[=]', '', regex=True).replace('["]', '', regex=True)
        print(df.head().to_markdown())
        df.dropna(subset=["订单编号"], inplace=True)
        df = df[~df["订单编号"].str.contains("nan")]
        df["金额"] = df["金额"].astype(float)
        # df["账单日期"] = df["账单日期"].astype("datetime64[ns]")
        print(df.head(5).to_markdown())
        plat = "JD"
        if "钱包结算备注" in df.columns:
            pass
        else:
            print("不符合账单格式！")
            dict = {"TID": "nan", "SHOPNAME": "", "PLATFORM": "", "SHOPCODE": "", "BILLPLATFORM": "", "CREATED": "",
                    "TITLE": "", "TRADE_TYPE": "", "BUSINESS_NO": "", "INCOME_AMOUNT": "", "EXPEND_AMOUNT": "",
                    "TRADING_CHANNELS": "", "BUSINESS_DESCRIPTION": "", "remark": "", "BUSINESS_BILL_SOURCE": "",
                    "IS_REFUNDAMOUNT": "", "IS_AMOUNT": "", "OID": "", "SOURCEDATA": "", "RECIPROCAL_ACCOUNT": "",
                    "BATCHNO": ""}
            df = pd.DataFrame(dict, index=[0])
            return df
        print("逻辑处理中...")
        df1 = pd.DataFrame()
        df1["TID"] = df["订单编号"]
        df1["SHOPNAME"] = ""
        # if "店铺名称" in df.columns:
        #     df1.rename(columns={"店铺名称","SHOPNAME"},inplace=True)
        # else:
        #     if filename.find("2019") >= 0:
        #         shop = "".join(filename.split(os.sep)[-2:-1]).lower()
        #         # df1["SHOPNAME"] = df["SHOPNAME"].str.lower()
        #         shop = shop.replace("京东", "").replace(" ", "")
        #         df1["SHOPNAME"] = shop
        #     elif filename.find("2020") >=0:
        #         shop = "".join(filename.split(os.sep)[-1:]).lower()
        #         shop = "".join(shop.split("2020")[:1])
        #         shop = shop.replace("京东", "").replace(" ", "")
        #         df1["SHOPNAME"] = shop
        #     elif filename.find("2021") >=0:
        #         shop = "".join(filename.split(os.sep)[-1:]).lower()
        #         if shop.find("订单")>=0:
        #             shop = "".join(shop.split("订单")[:1])
        #         else:
        #             shop = "".join(shop.split("2021")[:1])
        #         shop = shop.replace("京东", "").replace(" ", "")
        #         df1["SHOPNAME"] = shop
        #     print(shop)
        df1["PLATFORM"] = plat
        print("定位1")
        # df1["SHOPCODE"] = get_shopcode(plat,shop)
        # df1["SHOPCODE"] = df1.apply(lambda x:get_shopcode(x["PLATFORM"],x["SHOPNAME"]),axis=1)
        df1["SHOPCODE"] = ""
        print("定位2")
        df1["BILLPLATFORM"] = "JD"
        # df1["CREATED"] = df.apply(lambda x:x["账单日期"]+datetime.timedelta(days=30) if pd.isnull(x["费用结算时间"]) else x["费用结算时间"],axis=1)
        df["结算时间"] = df.apply(
            lambda x: datetime.datetime.strptime(x["账单日期"][:4] + "-" + str(int(x["账单日期"][4:6]) + 1) + '-01 00:00:00',
                                                 '%Y-%m-%d %H:%M:%S') if int(
                x["账单日期"][4:6]) < 12 else datetime.datetime.strptime(str(int(x["账单日期"][:4]) + 1) + '-01-01 00:00:00',
                                                                     '%Y-%m-%d %H:%M:%S'), axis=1)
        # df1["CREATED"] = df.apply(lambda x:datetime.datetime.strptime(x["账单日期"][:4]+"-"+str(int(x["账单日期"][4:6])+1)+'-01 00:00:00','%Y-%m-%d %H:%M:%S') if ((pd.isnull(x["费用结算时间"]))&(int(x["账单日期"][4:6])<12)) else x["费用结算时间"],axis=1)
        # df1["CREATED"] = df.apply(lambda x:datetime.datetime.strptime(str(int(x["账单日期"][:4])+1)+'-01-01 00:00:00','%Y-%m-%d %H:%M:%S ') if ((pd.isnull(x["费用结算时间"]))&(x["账单日期"][:6].find("12")>=0)) else x["费用结算时间"],axis=1)
        # df1["CREATED"] = df.apply(lambda x:datetime.datetime.strptime(x["账单日期"][:4]+"-"+str(int(x["账单日期"][4:6])+1)+'-01 00:00:00','%Y-%m-%d %H:%M:%S') if pd.isnull(x["费用结算时间"]) else x["费用结算时间"],axis=1)
        df1["CREATED"] = df.apply(lambda x: x["结算时间"] if pd.isnull(x["费用结算时间"]) else x["费用结算时间"], axis=1)
        df1["TITLE"] = df["商品名称"]
        df1["TRADE_TYPE"] = df["钱包结算备注"].apply(lambda x: "".join(x.split("日")[-1:]) if x.find("日") >= 0 else x)
        df1["BUSINESS_NO"] = ""
        df1["INCOME_AMOUNT"] = df["金额"].apply(lambda x: x if x > 0 else 0)
        df1["EXPEND_AMOUNT"] = df["金额"].apply(lambda x: x if x < 0 else 0)
        df1["TRADING_CHANNELS"] = ""
        df1["BUSINESS_DESCRIPTION"] = df["费用项"]
        df1["remark"] = ""
        df1["BUSINESS_BILL_SOURCE"] = ""
        df1["IS_REFUNDAMOUNT"] = df.apply(lambda x: JD_IS_REFUNDAMOUNT(x["费用项"], x["收支方向"], x["金额"], x["钱包结算备注"]),
                                          axis=1)
        df1["IS_AMOUNT"] = df.apply(lambda x: JD_IS_AMOUNT(x["费用项"], x["收支方向"], x["金额"], x["钱包结算备注"]), axis=1)
        df1["OID"] = ""
        df1["SOURCEDATA"] = "EXCEL"
        df1["RECIPROCAL_ACCOUNT"] = ""
        df1["BATCHNO"] = ""
        df1["currency"] = ""
        df1["overseas_income"] = ""
        df1["overseas_expend"] = ""
        df1["currency_cny_rate"] = ""
        df["CREATED"] = df1["CREATED"].astype(str)
        df["INCOME_AMOUNT"] = df1["INCOME_AMOUNT"].astype(str)
        df["EXPEND_AMOUNT"] = df1["EXPEND_AMOUNT"].astype(str)
        df["IS_AMOUNT"] = df1["IS_AMOUNT"].astype(str)
        df["IS_REFUNDAMOUNT"] = df1["IS_REFUNDAMOUNT"].astype(str)
        df1["BUSINESS_NO"] = df["单据编号"] + df1["TID"] + df["INCOME_AMOUNT"] + df["EXPEND_AMOUNT"] + df["CREATED"] + df1[
            "TRADE_TYPE"] + df["IS_AMOUNT"] + df["IS_REFUNDAMOUNT"]

        df1["INCOME_AMOUNT"] = df1["INCOME_AMOUNT"].astype(float)
        df1["EXPEND_AMOUNT"] = df1["EXPEND_AMOUNT"].astype(float)
        # df1 = df1[
        #     ["TID", "SHOPNAME", "PLATFORM", "SHOPCODE", "BILLPLATFORM", "CREATED", "TITLE", "TRADE_TYPE", "BUSINESS_NO",
        #      "INCOME_AMOUNT", "EXPEND_AMOUNT", "TRADING_CHANNELS", "BUSINESS_DESCRIPTION", "remark",
        #      "BUSINESS_BILL_SOURCE", "IS_REFUNDAMOUNT", "IS_AMOUNT", "OID", "SOURCEDATA", "RECIPROCAL_ACCOUNT",
        #      "BATCHNO"]]
        df1["CREATED"] = df1["CREATED"].astype("datetime64[ns]")
        df1["iyear"] = df1["CREATED"].apply(lambda x: x.year)
        df1["imonth"] = df1["CREATED"].apply(lambda x: x.month)
        if ((filename.find("美妆店") >= 0) | (filename.find("麦凯莱美妆") >= 0)):
            df1["SHOPCODE"] = "mklmzzyd"
            df1["SHOPNAME"] = "麦凯莱美妆专营店"
        elif (((filename.find("loshi") >= 0) | (filename.find("Loshi") >= 0) | (filename.find("LoShi") >= 0)) & (
                filename.find("拼购") >= 0)):
            df1["SHOPCODE"] = "loshipgqjd"
            df1["SHOPNAME"] = "loshi拼购旗舰店"
        elif (((filename.find("loshi") >= 0) | (filename.find("Loshi") >= 0) | (filename.find("LoShi") >= 0)) & (
                filename.find("拼购") < 0)):
            df1["SHOPCODE"] = "loshijdzy"
            df1["SHOPNAME"] = "loshi京东自营旗舰店"
        elif filename.find("优丽氏旗舰店") >= 0:
            df1["SHOPCODE"] = "ylsjjd"
            df1["SHOPNAME"] = "优丽氏旗舰店"
        elif filename.find("Goat") >= 0:
            df1["SHOPCODE"] = "goatsoaphczmd"
            df1["SHOPNAME"] = "goatsoap宏炽专卖店"
        elif ((filename.find("Ultardex") >= 0) or (filename.find("Ultradex") >= 0)):
            df1["SHOPCODE"] = "ultradex"
            df1["SHOPNAME"] = "ultradex旗舰店"
        elif filename.find("atural") >= 0:
            df1["SHOPCODE"] = "allnaturaladvicegfqjd"
            df1["SHOPNAME"] = "All Natural Advice官方旗舰店"
            # df1 = df1.loc[((df1.imonth == 4) | (df1.imonth == 5))]
        elif ((filename.find("entyl") >= 0) & (filename.find("拼购旗舰店") >= 0)):
            df1["SHOPCODE"] = "dentylactivepgjjd"
            df1["SHOPNAME"] = "dentylactive拼购旗舰店"
        elif ((filename.find("entyl") >= 0) & (filename.find("自营店") >= 0)):
            df1["SHOPCODE"] = "dentylactivejdzyd"
            df1["SHOPNAME"] = "dentylactive京东自营旗舰店"
        elif ((filename.find("entyl") >= 0) & (filename.find("旗舰店") >= 0)):
            df1["SHOPCODE"] = "dentylactive"
            df1["SHOPNAME"] = "dentylactive旗舰店"
        elif filename.find("MOREI旗舰店") >= 0:
            df1["SHOPCODE"] = "moriqjd"
            df1["SHOPNAME"] = "MOREI旗舰店"
        elif filename.find("博滴官方旗舰店") >= 0:
            df1["SHOPCODE"] = "bdgfqjd"
            df1["SHOPNAME"] = "博滴官方旗舰店"
        elif ((filename.find("VitaBloom") >= 0) | (filename.find("vitabloom") >= 0)):
            df1["SHOPCODE"] = "vbqjd"
            df1["SHOPNAME"] = "VitaBloom京东官方自营旗舰店"
        elif ((filename.find("entyl") >= 0) & (filename.find("旗舰店") >= 0)):
            df1["SHOPCODE"] = "dentylactive"
            df1["SHOPNAME"] = "dentylactive旗舰店"
        elif ((filename.find("lab旗舰店") >= 0)):
            df1["SHOPCODE"] = "smilelab"
            df1["SHOPNAME"] = "smilelab旗舰店"
        elif filename.find("LCN麦凯莱") >= 0:
            df1["SHOPCODE"] = "lcnmklzmd"
            df1["SHOPNAME"] = "lcn麦凯莱专卖店"
        elif filename.find("惠优购官方旗舰店") >= 0:
            df1["SHOPCODE"] = "hyggfqjd"
            df1["SHOPNAME"] = "惠优购官方旗舰店"
        elif filename.find("mades京东旗舰店") >= 0:
            df1["SHOPCODE"] = "jdmades"
            df1["SHOPNAME"] = "京东mades旗舰店"
        del df1["iyear"]
        del df1["imonth"]
        print(df1.head(5).to_markdown())

    # 微盟逻辑
    elif filename.find("微盟") >= 0:
        df = pd.read_excel(filename, dtype=str)
        df["交易金额"] = df["交易金额"].astype(float)
        df["交易单号"] = df["交易单号"].apply(lambda x: x.replace(" ", "").strip())
        df["第三方交易单号"] = df["第三方交易单号"].apply(lambda x: x.replace(" ", "").strip())
        print(df.head(5).to_markdown())
        plat = "WM"
        df1 = pd.DataFrame()
        df1["TID"] = df["交易单号"]
        if filename.find("万加脉伽") >= 0:
            df1["SHOPNAME"] = "万加脉伽优选"
        elif filename.find("依娜心选") >= 0:
            df1["SHOPNAME"] = "依娜心选好物"
        df1["PLATFORM"] = plat
        if filename.find("万加脉伽") >= 0:
            df1["SHOPCODE"] = "wjmjyx"
        elif filename.find("依娜心选") >= 0:
            df1["SHOPCODE"] = "ynxxhw"
        df1["BILLPLATFORM"] = plat
        df1["CREATED"] = df["交易时间"]
        df1["TITLE"] = ""
        df1["TRADE_TYPE"] = df["交易场景"]
        df1["BUSINESS_NO"] = df["第三方交易单号"]
        df1["INCOME_AMOUNT"] = df.apply(lambda x: x["交易金额"] if x["交易金额"] > 0 else 0, axis=1)
        df1["EXPEND_AMOUNT"] = df.apply(lambda x: x["交易金额"] if x["交易金额"] < 0 else 0, axis=1)
        df1["TRADING_CHANNELS"] = df["支付方式"]
        df1["BUSINESS_DESCRIPTION"] = df["收支类型"]
        df1["IS_REFUNDAMOUNT"] = df.apply(lambda x: 1 if x["交易场景"] == "订单退款" else 0, axis=1)
        df1["IS_AMOUNT"] = df.apply(lambda x: 1 if x["交易场景"] == "网店订单" else 0, axis=1)
        df1["OID"] = ""
        df1["SOURCEDATA"] = "EXCEL"
        df1["RECIPROCAL_ACCOUNT"] = ""
        df1["BATCHNO"] = ""

        df1["INCOME_AMOUNT"] = df1["INCOME_AMOUNT"].astype(float).abs()
        df1["EXPEND_AMOUNT"] = df1["EXPEND_AMOUNT"].astype(float)
        print(df1.head(5).to_markdown())
        print("微盟")

    # 网易考拉逻辑
    elif filename.find("考拉") >= 0:
        if filename.find("其他费用") >= 0:
            df = pd.read_excel(filename, sheet_name="汇总账单", dtype=str)
            print(df.head(5).to_markdown())
            for column_name in df.columns:
                df.rename(columns={column_name: column_name.replace(" ", "").replace("\n", "").strip()}, inplace=True)
            plat = "KAOLA"
            # 优惠券赔付金额
            df1 = pd.DataFrame()
            df1["TID"] = df["商家ID"]
            df1["TID"] = ""
            df1["SHOPNAME"] = ""
            df1["PLATFORM"] = plat
            df1["SHOPCODE"] = ""
            df1["BILLPLATFORM"] = plat
            df1["CREATED"] = df["结算日期"]
            df1["TITLE"] = ""
            df1["TRADE_TYPE"] = "其它费用"
            df1["BUSINESS_NO"] = ""
            df1["INCOME_AMOUNT"] = df["优惠券赔付金额"].astype(float).apply(lambda x: x if x < 0 else 0)
            df1["EXPEND_AMOUNT"] = df["优惠券赔付金额"].astype(float).apply(lambda x: x if x > 0 else 0)
            df1["TRADING_CHANNELS"] = ""
            df1["BUSINESS_DESCRIPTION"] = "其它金额-优惠券赔付金额"
            df1["remark"] = ""
            df1["BUSINESS_BILL_SOURCE"] = ""
            df1["IS_REFUNDAMOUNT"] = 0
            df1["IS_AMOUNT"] = 0
            df1["OID"] = ""
            df1["SOURCEDATA"] = "EXCEL"
            df1["RECIPROCAL_ACCOUNT"] = ""
            df1["BATCHNO"] = ""
            df1["currency"] = ""
            df1["overseas_income"] = ""
            df1["overseas_expend"] = ""
            df1["currency_cny_rate"] = ""
            df1["index"] = "G"

            # 违规扣款金额
            df2 = df1.copy()
            df2["INCOME_AMOUNT"] = df["违规扣款金额"].astype(float).apply(lambda x: x if x < 0 else 0)
            df2["EXPEND_AMOUNT"] = df["违规扣款金额"].astype(float).apply(lambda x: x if x > 0 else 0)
            df2["BUSINESS_DESCRIPTION"] = "其它金额-违规扣款金额"
            df2["index"] = "H"
            # 短信费
            df3 = df1.copy()
            df3["INCOME_AMOUNT"] = df["短信费"].astype(float).apply(lambda x: x if x < 0 else 0)
            df3["EXPEND_AMOUNT"] = df["短信费"].astype(float).apply(lambda x: x if x > 0 else 0)
            df3["BUSINESS_DESCRIPTION"] = "其它金额-短信费"
            df3["index"] = "I"
            # 退款申诉金额
            df4 = df1.copy()
            df4["INCOME_AMOUNT"] = df["退款申诉金额"].astype(float).apply(lambda x: x if x < 0 else 0)
            df4["EXPEND_AMOUNT"] = df["退款申诉金额"].astype(float).apply(lambda x: x if x > 0 else 0)
            df4["BUSINESS_DESCRIPTION"] = "其它金额-退款申诉金额"
            df4["index"] = "J"
            # 现金赔付申诉金额
            df5 = df1.copy()
            df5["INCOME_AMOUNT"] = df["现金赔付申诉金额"].astype(float).apply(lambda x: x if x < 0 else 0)
            df5["EXPEND_AMOUNT"] = df["现金赔付申诉金额"].astype(float).apply(lambda x: x if x > 0 else 0)
            df5["BUSINESS_DESCRIPTION"] = "其它金额-现金赔付申诉金额"
            df5["index"] = "K"
            # CPS返佣
            df6 = df1.copy()
            df6["INCOME_AMOUNT"] = df["CPS返佣"].astype(float).apply(lambda x: x if x < 0 else 0)
            df6["EXPEND_AMOUNT"] = df["CPS返佣"].astype(float).apply(lambda x: x if x > 0 else 0)
            df6["BUSINESS_DESCRIPTION"] = "其它金额-CPS返佣"
            df6["index"] = "L"
            # 分享赚红包
            df7 = df1.copy()
            df7["INCOME_AMOUNT"] = df["分享赚红包"].astype(float).apply(lambda x: x if x < 0 else 0)
            df7["EXPEND_AMOUNT"] = df["分享赚红包"].astype(float).apply(lambda x: x if x > 0 else 0)
            df7["BUSINESS_DESCRIPTION"] = "其它金额-分享赚红包"
            df7["index"] = "M"
            # 商家综合服务费
            df8 = df1.copy()
            df8["INCOME_AMOUNT"] = df["商家综合服务费"].astype(float).apply(lambda x: x if x < 0 else 0)
            df8["EXPEND_AMOUNT"] = df["商家综合服务费"].astype(float).apply(lambda x: x if x > 0 else 0)
            df8["BUSINESS_DESCRIPTION"] = "其它金额-商家综合服务费"
            df8["index"] = "N"
            # 电子发票费
            df9 = df1.copy()
            df9["INCOME_AMOUNT"] = df["电子发票费"].astype(float).apply(lambda x: x if x < 0 else 0)
            df9["EXPEND_AMOUNT"] = df["电子发票费"].astype(float).apply(lambda x: x if x > 0 else 0)
            df9["BUSINESS_DESCRIPTION"] = "其它金额-电子发票费"
            df9["index"] = "O"
            # 售后退款补回
            df10 = df1.copy()
            try:
                df10["INCOME_AMOUNT"] = df["售后退款补回"].astype(float).apply(lambda x: x if x < 0 else 0)
                df10["EXPEND_AMOUNT"] = df["售后退款补回"].astype(float).apply(lambda x: x if x > 0 else 0)
                df10["BUSINESS_DESCRIPTION"] = "其它金额-售后退款补回"
                df10["index"] = "P"
            except Exception as e:
                df10["INCOME_AMOUNT"] = 0
                df10["EXPEND_AMOUNT"] = 0
                df10["BUSINESS_DESCRIPTION"] = "其它金额-售后退款补回"
                df10["index"] = "P"
            # 广告营销费用
            df11 = df1.copy()
            try:
                df11["INCOME_AMOUNT"] = df["广告营销费用"].astype(float).apply(lambda x: x if x < 0 else 0)
                df11["EXPEND_AMOUNT"] = df["广告营销费用"].astype(float).apply(lambda x: x if x > 0 else 0)
                df11["BUSINESS_DESCRIPTION"] = "其它金额-广告营销费用"
                df11["index"] = "Q"
            except Exception as e:
                df11["INCOME_AMOUNT"] = 0
                df11["EXPEND_AMOUNT"] = 0
                df11["BUSINESS_DESCRIPTION"] = "其它金额-广告营销费用"
                df11["index"] = "Q"
            # 入仓技术服务费
            df12 = df1.copy()
            try:
                df12["INCOME_AMOUNT"] = df["入仓技术服务费"].astype(float).apply(lambda x: x if x < 0 else 0)
                df12["EXPEND_AMOUNT"] = df["入仓技术服务费"].astype(float).apply(lambda x: x if x > 0 else 0)
                df12["BUSINESS_DESCRIPTION"] = "其它金额-入仓技术服务费"
                df12["index"] = "R"
            except Exception as e:
                df12["INCOME_AMOUNT"] = 0
                df12["EXPEND_AMOUNT"] = 0
                df12["BUSINESS_DESCRIPTION"] = "其它金额-入仓技术服务费"
                df12["index"] = "R"
            # 仓储赔付
            df13 = df1.copy()
            try:
                df13["INCOME_AMOUNT"] = df["仓储赔付"].astype(float).apply(lambda x: x if x < 0 else 0)
                df13["EXPEND_AMOUNT"] = df["仓储赔付"].astype(float).apply(lambda x: x if x > 0 else 0)
                df13["BUSINESS_DESCRIPTION"] = "其它金额-仓储赔付"
                df13["index"] = "S"
            except Exception as e:
                df13["INCOME_AMOUNT"] = 0
                df13["EXPEND_AMOUNT"] = 0
                df13["BUSINESS_DESCRIPTION"] = "其它金额-仓储赔付"
                df13["index"] = "S"

            dfs = [df1, df2, df3, df4, df5, df6, df7, df8, df9, df10, df11, df12, df13]
            df1 = pd.concat(dfs)
            df1["INCOME_AMOUNT"] = df1["INCOME_AMOUNT"].abs()
            df1["EXPEND_AMOUNT"] = -df1["EXPEND_AMOUNT"].abs()
            df1 = df1.loc[~((df1.INCOME_AMOUNT == 0) & (df1.EXPEND_AMOUNT == 0))]
            # df1 = pd.DataFrame(df1).reset_index()
            # df1["index"] = df1.index + 1
            df1["TID"] = "OF" + df1["SHOPCODE"] + df1["CREATED"].str.replace("-", "") + df1["index"]
            df1["BUSINESS_NO"] = df1["TID"]
            del df1["index"]
            print(df1.head().to_markdown())

        else:
            df = pd.read_excel(filename, sheet_name="销售明细", dtype=str)
            print(df.head(5).to_markdown())
            for column_name in df.columns:
                df.rename(columns={column_name: column_name.replace(" ", "").replace("\n", "").strip()}, inplace=True)
            # df.dropna(subset=["订单编号"], inplace=True)
            # df = df[~df["订单编号"].str.contains("nan")]
            df["商品销售额"] = df["商品销售额"].astype(float)
            if "商家承担折扣优惠总金额" in df.columns:
                rfee = "商家承担折扣优惠总金额"
                df[rfee] = df[rfee].astype(float)
            else:
                rfee = "应扣商家优惠总金额"
                df[rfee] = df[rfee].astype(float)
            # if "应付平台服务费" in df.columns:
            #     sfee = "应付平台服务费"
            #     df[sfee] = df[sfee].astype(float)
            # else:
            #     sfee = "收取的平台技术服务费"
            #     df[sfee] = df[sfee].astype(float)
            df["商品运费"] = df["商品运费"].astype(float)
            print(df.head(5).to_markdown())
            plat = "KAOLA"

            # 商品实付
            df1 = pd.DataFrame()
            df1["TID"] = df["销售订单号"]
            df1["SHOPNAME"] = ""
            df1["PLATFORM"] = plat
            df1["SHOPCODE"] = ""
            df1["BILLPLATFORM"] = plat
            df1["CREATED"] = df["结算日期"]
            df1["TITLE"] = df["商品名称"]
            df1["TRADE_TYPE"] = "商品实付"
            df1["BUSINESS_NO"] = ""
            df1["INCOME_AMOUNT"] = df["商品销售额"] - df[rfee]
            df1["EXPEND_AMOUNT"] = 0
            df1["TRADING_CHANNELS"] = ""
            df1["BUSINESS_DESCRIPTION"] = "商品实付"
            df1["remark"] = ""
            df1["BUSINESS_BILL_SOURCE"] = ""
            df1["IS_REFUNDAMOUNT"] = 0
            df1["IS_AMOUNT"] = 1
            df1["OID"] = ""
            df1["SOURCEDATA"] = "EXCEL"
            df1["RECIPROCAL_ACCOUNT"] = ""
            df1["BATCHNO"] = ""
            df1["INCOME_AMOUNT"] = df1["INCOME_AMOUNT"].astype(float).abs()
            print(df1.head(5).to_markdown())

            # 收取的平台技术服务费 or 应付平台服务费
            if "应付平台服务费" in df.columns:
                fee_name = "应付平台服务费"
            elif "收取的平台技术服务费" in df.columns:
                fee_name = "收取的平台技术服务费"
            df2 = df1.copy()
            df2["TRADE_TYPE"] = fee_name
            df2["INCOME_AMOUNT"] = 0
            df2["EXPEND_AMOUNT"] = df[fee_name]
            df2["BUSINESS_DESCRIPTION"] = fee_name
            df2["IS_REFUNDAMOUNT"] = 0
            df2["IS_AMOUNT"] = 0
            df2["EXPEND_AMOUNT"] = -df2["EXPEND_AMOUNT"].astype(float)
            print(df2.head(5).to_markdown())

            # 商品运费
            df3 = df1.copy()
            df3["TRADE_TYPE"] = "商品运费"
            df3["INCOME_AMOUNT"] = 0
            df3["EXPEND_AMOUNT"] = df["商品运费"]
            df3["BUSINESS_DESCRIPTION"] = "商品运费"
            df3["IS_REFUNDAMOUNT"] = 0
            df3["IS_AMOUNT"] = 0
            df3["EXPEND_AMOUNT"] = -df3["EXPEND_AMOUNT"].astype(float)
            df3 = df3.loc[df3["EXPEND_AMOUNT"] != 0]
            print(df3.head(5).to_markdown())

            # 商品税费
            df7 = df1.copy()
            df7["TRADE_TYPE"] = "商品税费"
            df7["INCOME_AMOUNT"] = df["商品税费"].apply(lambda x: x if filename.find("海外") >= 0 else 0)
            df7["EXPEND_AMOUNT"] = df["商品税费"].apply(lambda x: x if filename.find("海外") < 0 else 0)
            df7["BUSINESS_DESCRIPTION"] = "商品税费"
            df7["IS_REFUNDAMOUNT"] = 0
            df7["IS_AMOUNT"] = df["商品税费"].apply(lambda x: 1 if filename.find("海外") >= 0 else 0)
            df7["INCOME_AMOUNT"] = df7["INCOME_AMOUNT"].astype(float)
            df7["EXPEND_AMOUNT"] = -df7["EXPEND_AMOUNT"].astype(float)
            df7 = df7.loc[~((df1.INCOME_AMOUNT == 0) & (df7.EXPEND_AMOUNT == 0))]
            print(df7.head(5).to_markdown())

            # try:
            df = pd.read_excel(filename, sheet_name="退款明细", dtype=str)
            print(df.head(5).to_markdown())
            for column_name in df.columns:
                df.rename(columns={column_name: column_name.replace(" ", "").replace("\n", "").strip()}, inplace=True)
            if "实退现金" in df.columns:
                ramt = "实退现金"
                df[ramt] = df[ramt].astype(float)
            else:
                ramt = "商品实退金额（含税含发货运费）"
                df[ramt] = df[ramt].astype(float)

            # 实退金额
            df4 = pd.DataFrame()
            df4["TID"] = df["销售订单号"]
            df4["SHOPNAME"] = df1["SHOPNAME"]
            df4["PLATFORM"] = plat
            df4["SHOPCODE"] = df1["SHOPCODE"]
            df4["BILLPLATFORM"] = plat
            df4["CREATED"] = df["退款时间"]
            df4["TITLE"] = ""
            df4["TRADE_TYPE"] = ramt
            df4["BUSINESS_NO"] = ""
            df4["INCOME_AMOUNT"] = 0
            df4["EXPEND_AMOUNT"] = df[ramt]
            df4["TRADING_CHANNELS"] = ""
            df4["BUSINESS_DESCRIPTION"] = ramt
            df4["remark"] = ""
            df4["BUSINESS_BILL_SOURCE"] = ""
            df4["IS_REFUNDAMOUNT"] = 1
            df4["IS_AMOUNT"] = 0
            df4["OID"] = ""
            df4["SOURCEDATA"] = "EXCEL"
            df4["RECIPROCAL_ACCOUNT"] = ""
            df4["BATCHNO"] = ""
            df4["EXPEND_AMOUNT"] = -df4["EXPEND_AMOUNT"].astype(float)
            print(df4.head(5).to_markdown())

            # 退还平台技术服务费
            df5 = df4.copy()
            df5["TRADE_TYPE"] = "退还平台技术服务费"
            df5["INCOME_AMOUNT"] = df["退还平台技术服务费"]
            df5["EXPEND_AMOUNT"] = 0
            df5["BUSINESS_DESCRIPTION"] = "退还平台技术服务费"
            df5["IS_REFUNDAMOUNT"] = 0
            df5["IS_AMOUNT"] = 0
            df5["INCOME_AMOUNT"] = df5["INCOME_AMOUNT"].astype(str)
            df5 = df5[~df5["INCOME_AMOUNT"].str.contains("收入")]
            df5["INCOME_AMOUNT"] = df5["INCOME_AMOUNT"].astype(float).abs()
            print(df5.head(5).to_markdown())

            # 平台优惠金额补还
            df6 = df4.copy()
            df6["TRADE_TYPE"] = "平台优惠金额补还"
            df6["INCOME_AMOUNT"] = 0
            df6["EXPEND_AMOUNT"] = df["平台优惠金额补还"]
            df6["BUSINESS_DESCRIPTION"] = "平台优惠金额补还"
            df6["IS_REFUNDAMOUNT"] = 1
            df6["IS_AMOUNT"] = 0
            df6["EXPEND_AMOUNT"] = -df6["EXPEND_AMOUNT"].astype(float).abs()
            print(df6.head(5).to_markdown())

            # 合并账单
            dfs = [df1, df2, df3, df4, df5, df6, df7]
            df1 = pd.concat(dfs)
            df1["INCOME_AMOUNT"].fillna(0, inplace=True)
            df1["EXPEND_AMOUNT"].fillna(0, inplace=True)
            df1 = df1.loc[~((df1.INCOME_AMOUNT == 0) & (df1.EXPEND_AMOUNT == 0))]
            print(df1.head(5).to_markdown())

            # df1 = df1.sort_values(by=["TID","CREATED"])
            # df1.drop_duplicates(inplace=True)
            # except Exception as e:
            #     print("无退款明细分表")
            #     # 合并账单
            #     dfs = [df1, df2, df3]
            #     df1 = pd.concat(dfs)
            #     print(df1.head(5).to_markdown())

        if filename.find("博滴旗舰店") >= 0:
            df1["SHOPNAME"] = "bodyaid博滴旗舰店"
            df1["SHOPCODE"] = "bodyaidqjd"
        elif filename.find("麦凯莱个护") >= 0:
            df1["SHOPNAME"] = "麦凯莱个护专营店"
            df1["SHOPCODE"] = "mklghzyd"
        elif filename.find("drbubble旗舰店") >= 0:
            df1["SHOPNAME"] = "drbubble旗舰店"
            df1["SHOPCODE"] = "dqjd"
        elif filename.find("allnaturaladvice旗舰店") >= 0:
            df1["SHOPNAME"] = "allnaturaladvice旗舰店"
            df1["SHOPCODE"] = "aqjd"
        elif filename.find("VitaBloom旗舰店") >= 0:
            df1["SHOPNAME"] = "VitaBloom旗舰店"
            df1["SHOPCODE"] = "klhgvqjd"
        elif ((filename.find("Mades海外") >= 0) or (filename.find("mades海外") >= 0)):
            df1["SHOPNAME"] = "mades海外旗舰店"
            df1["SHOPCODE"] = "madesovqjd"
        elif ((filename.find("Dentyl") >= 0) or (filename.find("dentyl") >= 0)):
            df1["SHOPNAME"] = "dentylactive旗舰店"
            df1["SHOPCODE"] = "dentylactive"
        elif ((filename.find("MOREI") >= 0) or (filename.find("morei") >= 0)):
            df1["SHOPNAME"] = "morei旗舰店"
            df1["SHOPCODE"] = "moreiqjd"


    # 唯品会逻辑
    elif filename.find("唯品会") >= 0:
        df = pd.read_excel(filename, dtype=str)
        print(df.head(5).to_markdown())
        df["唯品会佣金(E)"] = df["唯品会佣金(E)"].astype(float)
        df["商家应收(F=D-E)"] = df["商家应收(F=D-E)"].astype(float)
        df["账单期间"] = df["账单期间"].astype(str)
        df = df[~df["账单期间"].str.contains("nan")]
        print(df["账单期间"])
        plat = "WPH"
        # 唯品会佣金
        df1 = pd.DataFrame()
        df1["TID"] = df["订单编号"]
        df1["SHOPNAME"] = "优丽氏个护旗舰店"
        df1["PLATFORM"] = plat
        df1["SHOPCODE"] = "ylsghqjd"
        df1["BILLPLATFORM"] = plat
        # df1["CREATED"] = df["订单签收或退款时间"]
        # df["账单期间"] = df["账单期间"].apply(lambda x: datetime.datetime.strptime(x + "-01 00:00:00", '%Y-%m-%d %H:%M:%S'))
        df1["CREATED"] = df["账单期间"].apply(lambda x: datetime.datetime.strptime(x + "-01 00:00:00", '%Y-%m-%d %H:%M:%S'))
        df1["TITLE"] = ""
        df1["TRADE_TYPE"] = df["订单类型名称"]
        df1["BUSINESS_NO"] = df["全局ID"]
        df1["INCOME_AMOUNT"] = df.apply(lambda x: x["唯品会佣金(E)"] if x["唯品会佣金(E)"] > 0 else 0, axis=1)
        df1["EXPEND_AMOUNT"] = df.apply(lambda x: x["唯品会佣金(E)"] if x["唯品会佣金(E)"] < 0 else 0, axis=1)
        df1["TRADING_CHANNELS"] = ""
        df1["BUSINESS_DESCRIPTION"] = "唯品会佣金"
        df1["IS_REFUNDAMOUNT"] = df.apply(lambda x: 1 if x["唯品会佣金(E)"] < 0 else 0, axis=1)
        df1["IS_AMOUNT"] = df.apply(lambda x: 1 if x["唯品会佣金(E)"] > 0 else 0, axis=1)
        df1["OID"] = ""
        df1["SOURCEDATA"] = "EXCEL"
        df1["RECIPROCAL_ACCOUNT"] = ""
        df1["BATCHNO"] = ""
        df1["INCOME_AMOUNT"] = df1["INCOME_AMOUNT"].astype(float).abs()
        df1["EXPEND_AMOUNT"] = df1["EXPEND_AMOUNT"].astype(float)
        print(df1.head(5).to_markdown())

        # 商家应收
        df2 = df1.copy()
        df2["INCOME_AMOUNT"] = df.apply(lambda x: x["商家应收(F=D-E)"] if x["商家应收(F=D-E)"] > 0 else 0, axis=1)
        df2["EXPEND_AMOUNT"] = df.apply(lambda x: x["商家应收(F=D-E)"] if x["商家应收(F=D-E)"] < 0 else 0, axis=1)
        df2["BUSINESS_DESCRIPTION"] = "商家应收"
        df2["IS_REFUNDAMOUNT"] = df.apply(lambda x: 1 if x["商家应收(F=D-E)"] < 0 else 0, axis=1)
        df2["IS_AMOUNT"] = df.apply(lambda x: 1 if x["商家应收(F=D-E)"] > 0 else 0, axis=1)
        df2["INCOME_AMOUNT"] = df2["INCOME_AMOUNT"].astype(float).abs()
        df2["EXPEND_AMOUNT"] = df2["EXPEND_AMOUNT"].astype(float)
        print(df2.head(5).to_markdown())

        dfs = [df1, df2]
        df1 = pd.concat(dfs)
        print(df1.head(5).to_markdown())
        print("唯品会")

    # 金牛逻辑
    elif filename.find("金牛") >= 0:
        df = pd.read_excel(filename, dtype=str)
        print(df.head(5).to_markdown())
        # df = df[~df["结算时间"].str.contains("nan")]
        df.dropna(subset=["结算时间"], inplace=True)
        # print(df["结算时间"].head(10).to_markdown())
        plat = "JN"
        # 订单实付
        df1 = pd.DataFrame()
        df1["TID"] = df["订单号"]
        df1["SHOPNAME"] = ""
        df1["PLATFORM"] = plat
        df1["SHOPCODE"] = ""
        df1["BILLPLATFORM"] = plat
        df1["CREATED"] = df["结算时间"]
        df1["TITLE"] = ""
        df1["TRADE_TYPE"] = df["业务类型"]
        df1["BUSINESS_NO"] = ""
        df1["INCOME_AMOUNT"] = df["订单实付"]
        df1["EXPEND_AMOUNT"] = 0
        df1["TRADING_CHANNELS"] = df["结算账户"]
        df1["BUSINESS_DESCRIPTION"] = "订单实付"
        df1["remark"] = ""
        df1["IS_REFUNDAMOUNT"] = 0
        df1["IS_AMOUNT"] = 1
        df1["OID"] = ""
        df1["SOURCEDATA"] = "EXCEL"
        df1["RECIPROCAL_ACCOUNT"] = ""
        df1["BATCHNO"] = ""
        df1["currency"] = ""
        df1["overseas_income"] = ""
        df1["overseas_expend"] = ""
        df1["currency_cny_rate"] = ""

        df1["INCOME_AMOUNT"] = df1["INCOME_AMOUNT"].astype(float).abs()
        df1["EXPEND_AMOUNT"] = -df1["EXPEND_AMOUNT"].astype(float).abs()
        if filename.find("勃狄") >= 0:
            df1["SHOPCODE"] = "bdghqjd"
            df1["SHOPNAME"] = "勃狄店"
        elif filename.find("鑫桂") >= 0:
            df1["SHOPCODE"] = "bdxgqjd"
            df1["SHOPNAME"] = "博滴鑫桂旗舰店"
        elif filename.find("尚西") >= 0:
            df1["SHOPCODE"] = "gzsx"
            df1["SHOPNAME"] = "广州尚西"
        elif filename.find("肯妮诗") >= 0:
            df1["SHOPCODE"] = "knszyd"
            df1["SHOPNAME"] = "肯妮诗专营店"
        elif ((filename.find("Mega2020") >= 0) | (filename.find("mega2020") >= 0)):
            df1["SHOPCODE"] = "mega2020"
            df1["SHOPNAME"] = "Mega2020"
        elif filename.find("萌洁齿") >= 0:
            df1["SHOPCODE"] = "megajn"
            df1["SHOPNAME"] = "萌洁齿个护店"
        elif filename.find("mega卖场") >= 0:
            df1["SHOPCODE"] = "megamcqjd"
            df1["SHOPNAME"] = "mega卖场旗舰店"
        elif filename.find("morei多瑞") >= 0:
            df1["SHOPCODE"] = "moreidr"
            df1["SHOPNAME"] = "morei多瑞"
        elif filename.find("若蘅") >= 0:
            df1["SHOPCODE"] = "nochernrh"
            df1["SHOPNAME"] = "nochern若蘅"
        elif filename.find("泡研") >= 0:
            df1["SHOPCODE"] = "pyxn"
            df1["SHOPNAME"] = "泡研小牛"
        elif filename.find("商魂") >= 0:
            df1["SHOPCODE"] = "shzyd"
            df1["SHOPNAME"] = "商魂专营店"
        print(df1.head(5).to_markdown())

        # 平台补贴
        df2 = df1.copy()
        df2["INCOME_AMOUNT"] = df["平台补贴"]
        df2["EXPEND_AMOUNT"] = 0
        df2["BUSINESS_DESCRIPTION"] = "平台补贴"
        df2["IS_REFUNDAMOUNT"] = 0
        df2["IS_AMOUNT"] = 1
        df2["INCOME_AMOUNT"] = df2["INCOME_AMOUNT"].astype(float).abs()
        df2["EXPEND_AMOUNT"] = -df2["EXPEND_AMOUNT"].astype(float).abs()
        print(df2.head(5).to_markdown())

        # 订单退款
        df3 = df1.copy()
        df3["INCOME_AMOUNT"] = 0
        df3["EXPEND_AMOUNT"] = df["订单退款"]
        df3["BUSINESS_DESCRIPTION"] = "订单退款"
        df3["IS_REFUNDAMOUNT"] = 1
        df3["IS_AMOUNT"] = 0
        df3["INCOME_AMOUNT"] = df3["INCOME_AMOUNT"].astype(float).abs()
        df3["EXPEND_AMOUNT"] = -df3["EXPEND_AMOUNT"].astype(float).abs()
        print(df3.head(5).to_markdown())

        # 平台服务费
        df4 = df1.copy()
        df4["INCOME_AMOUNT"] = 0
        df4["EXPEND_AMOUNT"] = df["平台服务费"]
        df4["BUSINESS_DESCRIPTION"] = "平台服务费"
        df4["IS_REFUNDAMOUNT"] = 0
        df4["IS_AMOUNT"] = 0
        df4["INCOME_AMOUNT"] = df4["INCOME_AMOUNT"].astype(float).abs()
        df4["EXPEND_AMOUNT"] = -df4["EXPEND_AMOUNT"].astype(float).abs()
        print(df4.head(5).to_markdown())

        # 达人佣金
        df5 = df1.copy()
        df5["INCOME_AMOUNT"] = 0
        df5["EXPEND_AMOUNT"] = df["达人佣金"]
        df5["BUSINESS_DESCRIPTION"] = "达人佣金"
        df5["IS_REFUNDAMOUNT"] = 0
        df5["IS_AMOUNT"] = 0
        df5["INCOME_AMOUNT"] = df5["INCOME_AMOUNT"].astype(float).abs()
        df5["EXPEND_AMOUNT"] = -df5["EXPEND_AMOUNT"].astype(float).abs()
        print(df5.head(5).to_markdown())

        # 渠道分成
        df6 = df1.copy()
        df6["INCOME_AMOUNT"] = 0
        df6["EXPEND_AMOUNT"] = df["渠道分成"]
        df6["BUSINESS_DESCRIPTION"] = "渠道分成"
        df6["IS_REFUNDAMOUNT"] = 0
        df6["IS_AMOUNT"] = 0
        df6["INCOME_AMOUNT"] = df6["INCOME_AMOUNT"].astype(float).abs()
        df6["EXPEND_AMOUNT"] = -df6["EXPEND_AMOUNT"].astype(float).abs()
        print(df6.head(5).to_markdown())

        dfs = [df1, df2, df3, df4, df5, df6]
        df1 = pd.concat(dfs)
        df1 = df1.loc[~((df1.INCOME_AMOUNT == 0) & (df1.EXPEND_AMOUNT == 0))]
        print(df1.head(5).to_markdown())
        print("金牛")

    # 做梦吧逻辑
    elif filename.find("做梦吧") >= 0:
        if ((filename.find("支付宝") >= 0) & (filename.find("博滴官方旗舰店") >= 0) & (filename.find("汇总") < 0)):
            if filename.find("xls") >= 0:
                df = pd.read_excel(filename, skiprows=4, dtype=str)
            else:
                try:
                    df = pd.read_csv(filename, skiprows=4, dtype=str)
                except Exception as e:
                    df = pd.read_csv(filename, skiprows=4, dtype=str, encoding="gb18030")
            for column_name in df.columns:
                df.rename(columns={column_name: column_name.replace("\n", "").replace("\t", "").replace(" ", "")},
                          inplace=True)
            df = df[~df["账务流水号"].str.contains("#")]
            print(df.head(5).to_markdown())
            if "业务基础订单号" in df.columns:
                df["业务基础订单号"] = df["业务基础订单号"].str.replace("\s+", "")
            df["账务流水号"] = df["账务流水号"].str.replace("\s+", "")
            df["业务流水号"] = df["业务流水号"].str.replace("\s+", "")
            df["商户订单号"] = df["商户订单号"].str.replace("\s+", "")
            df = df.replace("", np.nan)
            print(df.head(5).to_markdown())
            plat = "ZMB"
            # 订单实付
            df1 = pd.DataFrame()
            if "业务基础订单号" in df.columns:
                df["TID"] = df.apply(lambda x: x["业务基础订单号"] if pd.notnull(x["业务基础订单号"]) else x["商户订单号"], axis=1)
                df1["TID"] = df.apply(lambda x: x["TID"] if pd.notnull(x["TID"]) else x["业务流水号"], axis=1)
            else:
                df1["TID"] = df.apply(lambda x: x["商户订单号"] if pd.notnull(x["商户订单号"]) else x["业务流水号"], axis=1)
            df1["SHOPNAME"] = "博滴官方旗舰店"
            df1["PLATFORM"] = plat
            df1["SHOPCODE"] = "bdghqjd"
            df1["BILLPLATFORM"] = plat
            df1["CREATED"] = df["发生时间"]
            df1["TITLE"] = ""
            df1["TRADE_TYPE"] = df["业务类型"]
            df1["BUSINESS_NO"] = df["账务流水号"]
            df1["INCOME_AMOUNT"] = df["收入金额（+元）"]
            df1["EXPEND_AMOUNT"] = df["支出金额（-元）"]
            df1["TRADING_CHANNELS"] = df["交易渠道"]
            df1["BUSINESS_DESCRIPTION"] = df["业务类型"]
            df1["remark"] = df["备注"]
            if "业务描述" in df.columns:
                df["支出金额（-元）"] = df["支出金额（-元）"].astype(float)
                df1["IS_REFUNDAMOUNT"] = df.apply(
                    lambda x: taobao_is_refund(x["业务描述"], x["商品名称"], x["业务类型"], x["备注"], x["支出金额（-元）"], x["商户订单号"],
                                               filename), axis=1)
            else:
                df["支出金额（-元）"] = df["支出金额（-元）"].astype(float)
                df1["IS_REFUNDAMOUNT"] = df.apply(
                    lambda x: taobao_is_refund("nan", x["商品名称"], x["业务类型"], x["备注"], x["支出金额（-元）"], x["商户订单号"],
                                               filename), axis=1)
            if "业务描述" in df.columns:
                df["收入金额（+元）"] = df["收入金额（+元）"].astype(float)
                df1["IS_AMOUNT"] = df.apply(
                    lambda x: taobao_is_amount(x["业务描述"], x["商品名称"], x["业务类型"], x["商户订单号"], x["收入金额（+元）"], filename),
                    axis=1)
            else:
                df["收入金额（+元）"] = df["收入金额（+元）"].astype(float)
                df1["IS_AMOUNT"] = df.apply(
                    lambda x: taobao_is_amount("nan", x["商品名称"], x["业务类型"], x["商户订单号"], x["收入金额（+元）"], filename),
                    axis=1)
            df1["OID"] = ""
            df1["SOURCEDATA"] = "EXCEL"
            df1["RECIPROCAL_ACCOUNT"] = ""
            df1["BATCHNO"] = ""
            df1["currency"] = ""
            df1["overseas_income"] = ""
            df1["overseas_expend"] = ""
            df1["currency_cny_rate"] = ""
            df1["INCOME_AMOUNT"] = df1["INCOME_AMOUNT"].astype(float).abs()
            df1["EXPEND_AMOUNT"] = -df1["EXPEND_AMOUNT"].astype(float).abs()
            print(df1.head(5).to_markdown())
            print("做梦吧支付宝账单")

        else:
            dict = {"TID": "nan", "SHOPNAME": "", "PLATFORM": "", "SHOPCODE": "", "BILLPLATFORM": "",
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

    # 百度逻辑
    elif filename.find("百度") >= 0:
        # if ((filename.find("2020")>=0)&(filename.find("睿旗")>=0)):
        #     df = pd.read_excel(filename, sheet_name="月账单", dtype=str)
        #     df = df[df["订单编号"].str.contains("nan|订单编号")]
        #     plat = "BD"
        #     df["金额"] = df["金额"].astype(float)
        #     df1 = pd.DataFrame()
        #     df1["TID"] = df["订单编号"]
        #     df1["SHOPNAME"] = "莳天-睿旗科技"
        #     df1["PLATFORM"] = plat
        #     df1["SHOPCODE"] = "ruiqi"
        #     df1["BILLPLATFORM"] = plat
        #     df1["CREATED"] = df["账单日期"]
        #     df1["TITLE"] = ""
        #     df1["TRADE_TYPE"] = "货款"
        #     df1["BUSINESS_NO"] = df["第三方订单ID"]
        #     df1["INCOME_AMOUNT"] = df.apply(lambda x: x["金额"] if x["金额"] > 0 else 0, axis=1)
        #     df1["EXPEND_AMOUNT"] = df.apply(lambda x: x["金额"] if x["金额"] < 0 else 0, axis=1)
        #     df1["TRADING_CHANNELS"] = ""
        #     df1["BUSINESS_DESCRIPTION"] = "货款"
        #     df1["remark"] = "商品ID：" + df["商品ID"] + "。操作类型：" + df["操作类型"]
        #     df1["IS_REFUNDAMOUNT"] = df.apply(lambda x: 1 if x["金额"] < 0 else 0, axis=1)
        #     df1["IS_AMOUNT"] = df.apply(lambda x: 1 if x["金额"] > 0 else 0, axis=1)
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

        if filename.find("货款") >= 0:
            try:
                df = pd.read_excel(filename, sheet_name="月账单", dtype=str)
            except Exception as e:
                # try:
                #     df = pd.read_excel(filename, sheet_name="总表", dtype=str)
                # except Exception as e:
                #     try:
                #         df = pd.read_excel(filename, sheet_name="汇总", dtype=str)
                #     except Exception as e:
                #         try:
                #             df = pd.read_excel(filename, sheet_name="麦凯莱科技", dtype=str)
                #         except Exception as e:
                try:
                    df = pd.read_excel(filename, sheet_name=None, dtype=str)
                    sheet_list = list(df)
                    print(sheet_list)
                    if len(sheet_list) > 1:
                        df = None
                        for sheet in sheet_list:
                            if ((sheet == "汇总") or (sheet == "总表")):
                                continue
                            df1 = pd.read_excel(filename, sheet_name=sheet, dtype=str)
                            df1["sheet"] = sheet
                            if df is None:
                                df = df1
                            else:
                                # print(f"df1:\n{df.head(1).to_markdown()}")
                                # print(f"df:\n{df1.head(1).to_markdown()}")
                                df = pd.concat([df, df1])
                                print(len(df))
                    else:
                        sheet = sheet_list[0]
                        df = pd.read_excel(filename, sheet_name=sheet, dtype=str)
                        df["sheet"] = sheet
                except Exception as e:
                    dict = {"TID": "nan", "SHOPNAME": "", "PLATFORM": "", "SHOPCODE": "",
                            "BILLPLATFORM": "",
                            "CREATED": "",
                            "TITLE": "", "TRADE_TYPE": "", "BUSINESS_NO": "", "INCOME_AMOUNT": "",
                            "EXPEND_AMOUNT": "",
                            "TRADING_CHANNELS": "", "BUSINESS_DESCRIPTION": "", "remark": "",
                            "BUSINESS_BILL_SOURCE": "",
                            "IS_REFUNDAMOUNT": "", "IS_AMOUNT": "", "OID": "", "SOURCEDATA": "",
                            "RECIPROCAL_ACCOUNT": "", "BATCHNO": ""}
                    df = pd.DataFrame(dict, index=[0])
                    return df

            print(df.head(5).to_markdown())
            plat = "BD"
            df["金额"] = df["金额"].astype(float)
            if "商品id" in df.columns:
                df.rename(columns={"商品id": "商品ID"}, inplace=True)

            shop_df = pd.read_pickle("data/百度-商品ID与店铺名称关系0(3)(2).pkl")

            # 货款
            df1 = pd.DataFrame()
            # df1["TID"] = df["百度订单ID"]
            df1["TID"] = df["第三方订单ID"]
            df1["SHOPNAME"] = ""
            if "麦凯莱科技" in df.columns:
                print("定位1")
                df1["SHOPNAME"] = df["麦凯莱科技"]
            elif "麦凯莱小店序号" in df.columns:
                print("定位2")
                df1["SHOPNAME"] = df["麦凯莱小店序号"]
            elif "序列号" in df.columns:
                print("定位3")
                df1["SHOPNAME"] = df["序列号"]
            elif "商品ID" in df.columns:
                print("定位5")
                df["商品ID"] = df["商品ID"].astype(str)
                df1["SHOPNAME"] = df["商品ID"].apply(lambda x: get_baidushop(x, filename, 1, shop_df))
            elif "sheet" in df.columns:
                print("定位4")
                df1["SHOPNAME"] = df["sheet"]
            else:
                print("定位6")
                shop = "".join(filename.split(os.sep)[-1:])
                if ((shop.find("广分电商") >= 0) | (shop.find("广西电商") >= 0)):
                    df1["SHOPNAME"] = "".join(shop.split("-")[:2])
                else:
                    df1["SHOPNAME"] = "".join(shop.split("-")[:1])
            df1["PLATFORM"] = plat
            df1["SHOPCODE"] = df1["SHOPNAME"].apply(lambda x: get_baidushop(x, filename, 2, shop_df))
            df1["BILLPLATFORM"] = plat
            df1["CREATED"] = df["账单日期"]
            df1["TITLE"] = ""
            df1["TRADE_TYPE"] = "货款"
            df1["BUSINESS_NO"] = df["第三方订单ID"]
            df1["INCOME_AMOUNT"] = df.apply(lambda x: x["金额"] if x["金额"] > 0 else 0, axis=1)
            df1["EXPEND_AMOUNT"] = df.apply(lambda x: x["金额"] if x["金额"] < 0 else 0, axis=1)
            df1["TRADING_CHANNELS"] = ""
            df1["BUSINESS_DESCRIPTION"] = "货款"
            df1["remark"] = "商品ID：" + df["商品ID"] + "。操作类型：" + df["操作类型"]
            df1["IS_REFUNDAMOUNT"] = df.apply(lambda x: 1 if x["金额"] < 0 else 0, axis=1)
            df1["IS_AMOUNT"] = df.apply(lambda x: 1 if x["金额"] > 0 else 0, axis=1)
            df1["OID"] = ""
            df1["SOURCEDATA"] = "EXCEL"
            df1["RECIPROCAL_ACCOUNT"] = ""
            df1["BATCHNO"] = ""
            df1["currency"] = ""
            df1["overseas_income"] = ""
            df1["overseas_expend"] = ""
            df1["currency_cny_rate"] = ""
            df1["INCOME_AMOUNT"] = df1["INCOME_AMOUNT"].astype(float).abs()
            df1["EXPEND_AMOUNT"] = -df1["EXPEND_AMOUNT"].astype(float).abs()

            df2 = df1.copy()
            df2["TRADE_TYPE"] = "佣金"
            df2["INCOME_AMOUNT"] = df.apply(lambda x: x["金额"] * 0.006 if x["金额"] < 0 else 0, axis=1)
            df2["EXPEND_AMOUNT"] = df.apply(lambda x: x["金额"] * 0.006 if x["金额"] > 0 else 0, axis=1)
            df2["BUSINESS_DESCRIPTION"] = "佣金"
            df2["IS_REFUNDAMOUNT"] = 0
            df2["IS_AMOUNT"] = 0
            df2["INCOME_AMOUNT"] = df2["INCOME_AMOUNT"].astype(float).abs()
            df2["EXPEND_AMOUNT"] = -df2["EXPEND_AMOUNT"].astype(float).abs()

            dfs = [df1, df2]
            df1 = pd.concat(dfs)
            df1 = df1.loc[~((df1.INCOME_AMOUNT == 0) & (df1.EXPEND_AMOUNT == 0))]
            # if filename.find("宏炽") >= 0:
            #     df1["SHOPCODE"] = "hzzt"
            #     df1["SHOPNAME"] = "宏炽主体"
            # elif filename.find("卖家联合") >= 0:
            #     df1["SHOPCODE"] = "mjlh"
            #     df1["SHOPNAME"] = "卖家联合主体"
            # elif filename.find("麦凯莱") >= 0:
            #     df1["SHOPCODE"] = "mklzt"
            #     df1["SHOPNAME"] = "麦凯莱主体"
            # elif filename.find("睿旗") >= 0:
            #     df1["SHOPCODE"] = "ruiqi"
            #     df1["SHOPNAME"] = "莳天-睿旗科技"
            # elif filename.find("尚西") >= 0:
            #     df1["SHOPCODE"] = "sxzt"
            #     df1["SHOPNAME"] = "尚西主体"
            # elif filename.find("鑫桂") >= 0:
            #     df1["SHOPCODE"] = "xgzt"
            #     df1["SHOPNAME"] = "鑫桂主体"
            # elif filename.find("可瘾") >= 0:
            #     df1["SHOPCODE"] = "xgkyzt"
            #     df1["SHOPNAME"] = "鑫桂/可瘾主体"

            print(df1.head(5).to_markdown())
            print("百度账单")

        else:
            dict = {"TID": "nan", "SHOPNAME": "", "PLATFORM": "", "SHOPCODE": "", "BILLPLATFORM": "", "CREATED": "",
                    "TITLE": "", "TRADE_TYPE": "", "BUSINESS_NO": "", "INCOME_AMOUNT": "", "EXPEND_AMOUNT": "",
                    "TRADING_CHANNELS": "", "BUSINESS_DESCRIPTION": "", "remark": "", "BUSINESS_BILL_SOURCE": "",
                    "IS_REFUNDAMOUNT": "", "IS_AMOUNT": "", "OID": "", "SOURCEDATA": "", "RECIPROCAL_ACCOUNT": "",
                    "BATCHNO": ""}
            df = pd.DataFrame(dict, index=[0])
            return df

    return df1


def get_baidushop(id, filename, type, df):
    # df = pd.read_excel("data/百度-商品ID与店铺名称关系0(3)(2).xlsx",dtype=str)
    shop_name = ""
    # print(df.to_markdown())
    # print(id)
    # print(df.head().to_markdown())
    if type == 1:
        if df[df["商品ID"].str.contains(id)].shape[0] > 0:
            # print("test1")
            # print(id)
            df = df[df["商品ID"].str.contains(id)]
            shop_name = df.iloc[0]["店铺名称"]
        else:
            # print("test2")
            # print(filename)
            shop = "".join(filename.split(os.sep)[-1:])
            # print(shop)
            if ((shop.find("广分电商") >= 0) | (shop.find("广西电商") >= 0)):
                shop_name = "-".join(shop.split("-")[:2])
                # print(shop_name)
            else:
                shop_name = "-".join(shop.split("-")[:1])
                # print(shop_name)
    else:
        if df[(df["店铺名称"].str.contains(id))].shape[0] > 0:
            # print("test2")
            df = df[df["店铺名称"].str.contains(id)]
            shop_name = df.iloc[0]["店铺代码"]
        else:
            shop_name = ""

    return shop_name


def jd_dentyl(time):
    # print("读取时间")
    # print(time)
    if ((time.find("201908") >= 0) | (time.find("201909") >= 0) | (time.find("202001") >= 0) | (
            time.find("202007") >= 0) | (time.find("202008") >= 0) | (time.find("202009") >= 0)):
        time = datetime.datetime.strptime(str(int(time) + 2) + "01", '%Y%m%d')
        # try:
        #     time = datetime.datetime.strptime(time[:5] + str(int(time[5:7]) + 2) + time[7:], '%Y-%m-%d %H:%M:%S')
        # except Exception as e:
        #     time = datetime.datetime.strptime(time[:5] + str(int(time[5:7]) + 2) + "-01" + time[10:], '%Y-%m-%d %H:%M:%S')
        # print("时间更新后：")
        # print(time)
        return time
    elif ((time.find("201910") >= 0) | (time.find("201911") >= 0) | (time.find("202002") >= 0) | (
            time.find("202003") >= 0) | (time.find("202004") >= 0) | (time.find("202005") >= 0) | (
                  time.find("202006") >= 0) | (time.find("202010") >= 0)):
        time = datetime.datetime.strptime(str(int(time) + 1) + "01", '%Y%m%d')
        # try:
        #     time = datetime.datetime.strptime(time[:5] + str(int(time[5:7]) + 1) + time[7:], '%Y-%m-%d %H:%M:%S')
        # except Exception as e:
        #     time = datetime.datetime.strptime(time[:5] + str(int(time[5:7]) + 1) + "-01" + time[10:], '%Y-%m-%d %H:%M:%S')
        # print("时间更新后：")
        # print(time)
        return time
    elif time.find("201912") >= 0:
        time = datetime.datetime.strptime("20200101", '%Y%m%d')
        # print("时间更新后：")
        # print(time)
        return time
    else:
        # print("无需更新时间：")
        # print(time)
        time = datetime.datetime.strptime(time + "01", '%Y%m%d')
        return time


def taobao_is_refund(desc, title, type, rmark, amount, btid, filename):
    # 天猫的loshi旗舰店、淘宝的麦凯莱品牌自营店，只有这两家店才需要区分线上线下的账单的回款和退款
    if (((filename.find("麦凯莱品牌") >= 0) & (filename.find("特价区") < 0) & (filename.find("淘宝") >= 0)) | (
            (filename.find("loshi旗舰店") >= 0) & (filename.find("天猫") >= 0))):
        tid = btid
        if rmark.find("-") >= 0:
            rmark_tid = "".join(rmark.split("-")[-1:])
            if rmark_tid.startswith("T") >= 0:
                tid = rmark_tid
    else:
        tid = "T"

    if title.find("自动售货机") >= 0:
        return 0

    if desc.find("nan") < 0:
        if ((desc.find("0020001|交易退款-余额退款") >= 0) & (tid.startswith("T")) & (amount < 0)):
            return 1
        elif ((desc.find("0020002|交易退款-保证金退款") >= 0) & (tid.startswith("T")) & (amount < 0)):
            return 1
        elif ((desc.find("0020005|交易退款-售中退款（极速回款）") >= 0) & (tid.startswith("T")) & (amount < 0)):
            return 1
        elif ((desc.find("0020011|交易退款-交易退款") >= 0) & (tid.startswith("T")) & (amount < 0)):
            return 1
        elif ((desc.find("064000200001|交易还款-提前收款-花呗交易") >= 0) & (tid.startswith("T")) & (amount < 0)):
            return 1
        elif ((desc.find("064000200002|交易还款-提前收款-售中退款") >= 0) & (tid.startswith("T")) & (amount < 0)):
            return 1
        else:
            return 0
    else:
        if ((type == "交易退款") & (tid.startswith("T")) & (amount < 0)):
            return 1
        elif ((rmark.find("售后支付") >= 0) & (tid.startswith("T")) & (amount < 0)):
            return 1
        elif ((rmark.find("保证金退款") >= 0) & (tid.startswith("T")) & (amount < 0)):
            return 1
        else:
            return 0


def taobao_is_amount(desc, title, type, btid, amount, filename):
    # 天猫的loshi旗舰店、淘宝的麦凯莱品牌自营店，只有这两家店才需要区分线上线下的账单的回款和退款
    if (((filename.find("麦凯莱品牌") >= 0) & (filename.find("特价区") < 0) & (filename.find("淘宝") >= 0)) | (
            (filename.find("loshi旗舰店") >= 0) & (filename.find("天猫") >= 0))):
        tid = btid
    else:
        tid = "T"

    if title.find("自动售货机") >= 0:
        return 0

    if desc.find("nan") < 0:
        if ((desc.find("0010001|交易收款-交易收款") >= 0) & (tid.startswith("T")) & (amount > 0)):
            return 1
        elif ((desc.find("0010002|交易收款-预售定金（买家责任不退还）") >= 0) & (tid.startswith("T")) & (amount > 0)):
            return 1
        elif ((desc.find("0010022|交易收款-提前收款") >= 0) & (tid.startswith("T")) & (amount > 0)):
            return 1
        elif ((desc.find("001002200001|交易收款-提前收款-花呗交易") >= 0) & (tid.startswith("T")) & (amount > 0)):
            return 1
        else:
            return 0
    else:
        if ((type == "交易付款") & (tid.startswith("T")) & (amount > 0)):
            return 1
        elif ((type == "在线支付") & (tid.startswith("T")) & (amount > 0)):
            return 1
        elif ((type == "交易付款") & (tid.startswith("T")) & (amount > 0)):
            return 1
        elif ((type == "在线支付") & (tid.startswith("T")) & (amount > 0)):
            return 1
        else:
            return 0


def alibaba_is_amount(desc, title, type, btid, amount, remark, filename):
    # 天猫的loshi旗舰店、淘宝的麦凯莱品牌自营店，只有这两家店才需要区分线上线下的账单的回款和退款
    if (((filename.find("麦凯莱品牌") >= 0) & (filename.find("特价区") < 0) & (filename.find("淘宝") >= 0)) | (
            (filename.find("loshi旗舰店") >= 0) & (filename.find("天猫") >= 0))):
        tid = btid
    else:
        tid = "T"

    if title.find("自动售货机") >= 0:
        return 0

    if ((remark.find("诚e赊买家还款") >= 0) | (title.find("诚e赊订单红包") >= 0)):
        return 1

    if desc.find("nan") < 0:
        if ((desc.find("0010001|交易收款-交易收款") >= 0) & (tid.startswith("T")) & (amount > 0)):
            return 1
        elif ((desc.find("0010002|交易收款-预售定金（买家责任不退还）") >= 0) & (tid.startswith("T")) & (amount > 0)):
            return 1
        elif ((desc.find("0010022|交易收款-提前收款") >= 0) & (tid.startswith("T")) & (amount > 0)):
            return 1
        elif ((desc.find("001002200001|交易收款-提前收款-花呗交易") >= 0) & (tid.startswith("T")) & (amount > 0)):
            return 1
        else:
            return 0
    else:
        if ((type == "交易付款") & (tid.startswith("T")) & (amount > 0)):
            return 1
        elif ((type == "在线支付") & (tid.startswith("T")) & (amount > 0)):
            return 1
        elif ((type == "交易付款") & (tid.startswith("T")) & (amount > 0)):
            return 1
        elif ((type == "在线支付") & (tid.startswith("T")) & (amount > 0)):
            return 1
        else:
            return 0


def taobao_desc(rmark, type, btid, title, description):
    # print("业务描述函数：")
    if title.find("自动售货机") >= 0:
        desc = "自动售货机-" + type
        # print(f"业务描述1：{desc}")
        return desc
    if description.find("nan") < 0:
        if ((description.find("0010001|交易收款-交易收款") >= 0) & (btid.startswith("T"))):
            desc = "线上" + type
        else:
            desc = description
        # print(f"业务描述1.01：{desc}")
        return desc
    elif type == "交易付款":
        desc = type
        # print(f"业务描述1.1：{desc}")
        return desc
    elif type == "在线支付":
        if rmark.find("[") > 0:
            if btid.find("T200P") >= 0:
                desc = "".join(rmark.split("[")[:1])
                desc = "T200P" + desc
                # print(f"业务描述1.11：{desc}")
                return desc
            elif btid.find("T100") >= 0:
                desc = "".join(rmark.split("[")[:1])
                desc = "T100" + desc
                # print(f"业务描述1.12：{desc}")
                return desc
            else:
                desc = "".join(rmark.split("[")[:1])
                # print(f"业务描述1.13：{desc}")
                return desc
        else:
            if btid.find("T200P") >= 0:
                desc = "T200P" + type
                # print(f"业务描述1.14：{desc}")
                return desc
            elif btid.find("T100") >= 0:
                desc = "T100" + type
                # print(f"业务描述1.15：{desc}")
                return desc
            else:
                # print(f"业务描述1.16：{type}")
                return type
    elif rmark.find("[") > 0:
        desc = "".join(rmark.split("[")[:1])
        if desc.find("-") >= 0:
            desc = "".join(desc.split("-")[:1])
            # print(f"业务描述1.2：{desc}")
            return desc
        else:
            # print(f"业务描述1.2：{desc}")
            return desc
    elif ((rmark.find("[") < 0) & (rmark.find("扣款用途：") >= 0)):
        desc = "".join(rmark.split("扣款用途：")[1:])
        if desc.find("，") >= 0:
            desc = "".join(desc.split("，")[:1])
            # print(f"业务描述2.1：{desc}")
            return desc
        elif desc.find("业务交易号") >= 0:
            desc = "".join(desc.split("业务交易号")[:1])
            # print(f"业务描述2.2：{desc}")
            return desc
        elif desc.find("-") >= 0:
            desc = "".join(desc.split("-")[:1])
            # print(f"业务描述2.3：{desc}")
            return desc
        elif desc.find("(") >= 0:
            desc = "".join(desc.split("(")[:1])
            # print(f"业务描述2.4：{desc}")
            return desc
        elif desc.find(")") >= 0:
            desc = "".join(desc.split(")")[:1])
            # print(f"业务描述2.5：{desc}")
            return desc
        elif desc.find("[") >= 0:
            desc = "".join(desc.split("[")[:1])
            # print(f"业务描述2.6：{desc}")
            return desc
        elif desc.find("]") >= 0:
            desc = "".join(desc.split("]")[:1])
            # print(f"业务描述2.7：{desc}")
            return desc
        elif desc.find("{") >= 0:
            desc = "".join(desc.split("{")[:1])
            # print(f"业务描述2.8：{desc}")
            return desc
        elif desc.find("}") >= 0:
            desc = "".join(desc.split("}")[:1])
            # print(f"业务描述2.9：{desc}")
            return desc
        elif desc.find("（") >= 0:
            desc = "".join(desc.split("（")[:1])
            # print(f"业务描述2.10：{desc}")
            return desc
        elif desc.find("）") >= 0:
            desc = "".join(desc.split("）")[:1])
            # print(f"业务描述2.11：{desc}")
            return desc
        elif desc.find("_") >= 0:
            desc = "".join(desc.split("_")[:1])
            # print(f"业务描述2.13：{desc}")
            return desc
        elif desc.find("。") >= 0:
            desc = "".join(desc.split("。")[:1])
            # print(f"业务描述2.14：{desc}")
            return desc
        else:
            # print(f"业务描述2.15：{desc}")
            return desc
    elif rmark.find("售后支付") >= 0:
        desc = "售后支付"
        # print(f"业务描述3：{desc}")
        return desc
    elif rmark.find("保证金退款") >= 0:
        desc = "保证金退款"
        # print(f"业务描述4：{desc}")
        return desc
    elif ((rmark.find("nan") >= 0) & (btid.find("T200P") >= 0)):
        desc = "T200P" + type
        # print(f"业务描述5：{desc}")
        return desc
    elif ((rmark.find("nan") >= 0) & (btid.find("T100") >= 0)):
        desc = "T100" + type
        # print(f"业务描述5：{desc}")
        return desc
    else:
        # print(f"业务描述6：{type}")
        return type


def taobao_tid(btid, type, rmark, stid):
    # print("订单号函数：")
    if ((btid.find("T") == 0) & (btid.find("0P") >= 0)):
        tid = "".join(btid.split("P")[-1:])
        # print(f"订单号1：{tid}")
        return tid
    elif btid.find("T500") >= 0:
        tid = "".join(btid.split("P")[-1:])
        # print(f"订单号1.1：{tid}")
        return tid
    elif btid.find("CAE_CHARITY_") >= 0:
        tid = "".join(btid.split("CAE_CHARITY_")[-1:])
        if tid.find("_") >= 0:
            tid = "".join(tid.split("_")[:1])
            # print(f"订单号2.1：{tid}")
            return tid
        else:
            # print(f"订单号2.2：{tid}")
            return tid
    elif btid.find("CAE_POINT_") >= 0:
        tid = "".join(btid.split("CAE_POINT_")[-1:])
        if tid.find("_") >= 0:
            tid = "".join(tid.split("_")[:1])
            # print(f"订单号3.1：{tid}")
            return tid
        else:
            # print(f"订单号3.2：{tid}")
            return tid
    elif ((type == "交易退款") & (rmark.find("T200P") >= 0)):
        tid = "".join(rmark.split("T200P")[-1:])
        # print(f"订单号4：{tid}")
        return tid
    elif ((type == "交易退款") & (rmark.find("T500P") >= 0)):
        tid = "".join(rmark.split("T500P")[-1:])
        # print(f"订单号5：{tid}")
        return tid
    elif ((rmark.find("[") >= 0) & (rmark.find("]") >= 0)):
        tid = "".join(rmark.split("[")[-1:])
        tid = "".join(tid.split("]")[:1])
        if str(tid).isnumeric():
            # print(f"订单号6.1：{tid}")
            return tid
        elif btid.find("nan") < 0:
            # print(f"订单号6.2：{btid}")
            return btid
        elif stid.find("nan") < 0:
            # print(f"订单号6.3：{stid}")
            return stid
    elif rmark.find("-") >= 0:
        tid = "".join(rmark.split("-")[-1:])
        if tid.find("T200") >= 0:
            tid = "".join(tid.split("P")[-1:])
            # print(f"订单号7.1：{tid}")
            return tid
        else:
            if btid.find("nan") < 0:
                tid = btid
                # print(f"订单号7.2：{tid}")
                return tid
            else:
                tid = stid
                # print(f"订单号7.3：{tid}")
                return tid
    elif btid.find("HJCAE==") >= 0:
        tid = "".join(btid.split("==")[-1:])
        if tid.find("nan") < 0:
            # print(f"订单号7.4：{tid}")
            return tid
    elif btid.find("nan") < 0:
        return btid
    else:
        # print(f"订单号8：{stid}")
        return stid


def JD_IS_REFUNDAMOUNT(BUSINESS_DESCRIPTION, ICOME_OUTCOME, AMOUNT, TRADE_TYPE):
    if ((BUSINESS_DESCRIPTION == "货款") & (ICOME_OUTCOME == "支出") & (AMOUNT < 0)):
        return 1
    elif ((TRADE_TYPE.find("退货") >= 0) & (ICOME_OUTCOME == "支出") & (AMOUNT < 0)):
        return 1
    elif ((TRADE_TYPE.find("代收配送费") >= 0) & (ICOME_OUTCOME == "支出") & (AMOUNT < 0)):
        return 1
    else:
        return 0


def JD_IS_AMOUNT(BUSINESS_DESCRIPTION, ICOME_OUTCOME, AMOUNT, TRADE_TYPE):
    if ((BUSINESS_DESCRIPTION == "货款") & (ICOME_OUTCOME == "收入") & (AMOUNT > 0)):
        return 1
    elif ((TRADE_TYPE.find("货款") >= 0) & (ICOME_OUTCOME == "收入") & (AMOUNT > 0)):
        return 1
    elif ((TRADE_TYPE.find("代收配送费") >= 0) & (ICOME_OUTCOME == "收入") & (AMOUNT > 0)):
        return 1
    else:
        return 0


def get_shopcode(PLATFORM, SHOPNAME, type):
    df = pd.read_excel("data/shopcode.xlsx")
    for index, row in df.iterrows():
        if PLATFORM.find(row["店铺平台"]) >= 0:
            # print(row["店铺平台"])
            # print("111")
            if SHOPNAME.find(row["店铺名称"]) >= 0:
                if type == 1:
                    shopname = row["店铺名称"]
                    return shopname
                else:
                    shopcode = row["店铺代码"]
                    # print("平台名:", PLATFORM, "店铺名称为：", SHOPNAME, "店铺代码为：", shopcode)
                    return shopcode
            else:
                # print("未找到店铺名称")
                pass
        else:
            # print("未找到店铺平台")
            pass
    return "N/A"
    # df["店铺代码"] = df["店铺代码"].loc[(df["店铺平台"].str.contains(PLATFORM) & df["店铺名称"].str.contains(SHOPNAME))]
    # df["店铺代码"] = ~df["店铺代码"].str.contains("nan")
    # print(df["店铺代码"])
    # return shop
    # shopcode = df["店铺代码"]
    # return shop
    # else:
    #     print("未找到店铺平台或者店铺名称")
    #     return "N/A"


def get_all_files(rootdir, filekey):
    filelist = list_all_files(rootdir, filekey)
    # print(filelist)
    mySeries = pd.Series(filelist)
    df = pd.DataFrame(mySeries)
    df.columns = ["filename"]
    # df["filename"] = df["filename"].apply(lambda x: x.strip().replace(" ", ""))

    df = df[~df["filename"].str.contains("汇总")]
    df = df[~df["filename"].str.contains("业务")]
    df = df[~df["filename"].str.contains("订单")]
    df = df[~df["filename"].str.contains("无数据")]

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
    table["SHOPNAME"] = table.apply(lambda x: x["SHOPNAME"] if len(x["SHOPNAME"]) > 2 else x["filename"], axis=1)
    del table["filename"]
    # 删除路径

    # if table.shape[0] < 800000:
    #     table.to_excel(default_dir + "/处理后的账单.xlsx", index=False)
    # else:
    #     table.to_csv(default_dir + "/处理后的账单.csv", index=False)
    index = 0
    # if "TID" in table.columns:
    if "PDD" in table["PLATFORM"].values.tolist():
        table.replace("nan", np.nan, inplace=True)
    # if "KAOLA" in table["PLATFORM"].values.tolist():
    #     table = table.loc[~((table.INCOME_AMOUNT == 0) & (table.EXPEND_AMOUNT == 0))]
    # table = pd.DataFrame(table).reset_index()
    # table["index"] = table.index + 1
    # table["TID"] = "OF" + table["SHOPCODE"] + table["CREATED"].str.replace("-","") + table["index"].map(lambda x: "{:0>6d}".format(x))
    # del table["index"]
    else:
        table["TID"] = table["TID"].astype(str)
        # table["TID"] = table["TID"].apply(lambda x:x.str.replace(" ",np.nan).str.replace("\n",np.nan))
        table.replace("nan", np.nan, inplace=True)
        table.dropna(subset=["TID"], inplace=True)
        # table.drop_duplicates(inplace=True)
    table = table.sort_values(by=["PLATFORM", "SHOPNAME", "TID", "CREATED"])
    table = table.loc[~((table.INCOME_AMOUNT == 0) & (table.EXPEND_AMOUNT == 0))]
    # print(table.head().to_markdown())
    # table["SHOPNAME"] = table.apply(lambda x: x["SHOPNAME"] if pd.notnull(x["SHOPNAME"]) else x["filename"], axis=1)
    table["CREATED"] = table["CREATED"].astype("datetime64[ns]")
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

    # 输出按月和店铺统计金额
    # table["SHOPNAME"] = table.apply(lambda x:x["SHOPNAME"] if pd.notnull(x["SHOPNAME"]) else x["filename"],axis=1)
    # table["CREATED"] = table["CREATED"].astype("datetime64[ns]")
    # table["month"] = table["CREATED"].dt.month
    # table["date"] = table["CREATED"].dt.date
    # table["date"] = table["date"].astype(str)
    # table["date"] = table["date"].apply(lambda x:x[:7])
    # groupby_table = table.groupby(["SHOPNAME","date","IS_REFUNDAMOUNT","IS_AMOUNT"]).agg({"INCOME_AMOUNT":"sum","EXPEND_AMOUNT":"sum",})
    # groupby_table = pd.DataFrame(groupby_table).reset_index()
    # groupby_table.to_excel(default_dir + "\分组统计后的账单.xlsx",index=False)

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
