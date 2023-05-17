# __coding=utf8__
# /** 作者：zengyanghui **/

import sys
import os
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
    if ((filename.find("海外") >= 0)|(filename.find("AMBRA") >= 0)):
        if filename.find("妥投结算") >= 0:
            df = pd.read_csv(filename, skiprows=1, dtype=str, encoding="gb18030")
            for column_name in df.columns:
                df.rename(columns={column_name: column_name.replace(" ", "").replace("\n", "").strip()},
                          inplace=True)
            df.dropna(subset=["订单编号"],inplace=True)
            print(df.head(1).to_markdown())
            df["货款"] = df["货款"].astype(float)
            df["佣金"] = df["佣金"].astype(float)
            df["汇率(美元/人民币)"] = df["汇率(美元/人民币)"].astype(float)
            df["订单编号"] = df["订单编号"].astype(str)

            plat = "JD"

            # 货款
            df1 = pd.DataFrame()
            df1["TID"] = df["订单编号"].apply(lambda x: x.replace(" ", "").strip())
            df1["SHOPNAME"] = ""
            df1["PLATFORM"] = plat
            df1["SHOPCODE"] = ""
            df1["BILLPLATFORM"] = "JDOVERSEAS"
            df1["CREATED"] = df["完成时间"]
            df1["TITLE"] = ""
            df1["TRADE_TYPE"] = "货款"
            df1["BUSINESS_NO"] = ""
            df1["INCOME_AMOUNT"] = df["货款"] * df["汇率(美元/人民币)"]
            df1["EXPEND_AMOUNT"] = 0
            df1["TRADING_CHANNELS"] = ""
            df1["BUSINESS_DESCRIPTION"] = "货款"
            df1["remark"] = ""
            df1["BUSINESS_BILL_SOURCE"] = ""
            df1["IS_REFUNDAMOUNT"] = 0
            df1["IS_AMOUNT"] = 1
            df1["OID"] = ""
            df1["SOURCEDATA"] = "EXCEL"
            df1["RECIPROCAL_ACCOUNT"] = ""
            df1["BATCHNO"] = ""
            df1["currency"] = "USD"
            df1["overseas_income"] = df["货款"]
            df1["overseas_expend"] = 0
            df1["currency_cny_rate"] = df["汇率(美元/人民币)"]

            if filename.find("Fit海外") >= 0:
                df1["SHOPCODE"] = "dicoraurbanfitov"
                df1["SHOPNAME"] = "dicoraurbanfit海外旗舰店"
            elif filename.find("mades海外") >= 0:
                df1["SHOPCODE"] = "madesovqjd"
                df1["SHOPNAME"] = "mades海外旗舰店"
            elif filename.find("MET海外") >= 0:
                df1["SHOPCODE"] = "manukascosmetov"
                df1["SHOPNAME"] = "manukascosmet海外旗舰店"
            elif filename.find("rai海外") >= 0:
                df1["SHOPCODE"] = "samouraiov"
                df1["SHOPNAME"] = "samourai海外旗舰店"
            elif ((filename.find("image海外") >= 0) | (filename.find("wiss海外") >= 0)):
                df1["SHOPCODE"] = "swissimageov"
                df1["SHOPNAME"] = "swissimage海外旗舰店"
            elif ((filename.find("icora") >= 0) & (filename.find("海外") >= 0)):
                df1["SHOPCODE"] = "dicoraurbanfitov"
                df1["SHOPNAME"] = "dicoraurbanfit海外旗舰店"
            elif filename.find("LILAC海外") >= 0:
                df1["SHOPCODE"] = "lilacovqjd"
                df1["SHOPNAME"] = "lilac海外旗舰店"
            elif filename.find("AMBRA") >= 0:
                df1["SHOPCODE"] = "ambrajdqjdov"
                df1["SHOPNAME"] = "AMBRA京东海外旗舰店"
            elif ((filename.find("MANUKA") >= 0) & (filename.find("海外") >= 0)):
                df1["SHOPCODE"] = "manukascosmetov"
                df1["SHOPNAME"] = "manukascosmet海外旗舰店"

            # 佣金
            df2 = df1.copy()
            df2["TRADE_TYPE"] = "佣金"
            df2["INCOME_AMOUNT"] = 0
            df2["EXPEND_AMOUNT"] = df["佣金"] * df["汇率(美元/人民币)"]
            df2["BUSINESS_DESCRIPTION"] = "佣金"
            df2["IS_REFUNDAMOUNT"] = 0
            df2["IS_AMOUNT"] = 0
            df2["overseas_income"] = 0
            df2["overseas_expend"] = df["佣金"]
            df2["currency_cny_rate"] = df["汇率(美元/人民币)"]

            dfs = [df1, df2]
            df1 = pd.concat(dfs)
            print(df1.head().to_markdown())

            return df1

        elif filename.find("退货结算") >= 0 :
            df = pd.read_csv(filename, skiprows=1, dtype=str, encoding="gb18030")
            for column_name in df.columns:
                df.rename(columns={column_name: column_name.replace(" ", "").replace("\n", "").strip()},
                          inplace=True)
            df = df[~df["返修单号"].str.contains("合计")]
            df = df[~df["订单编号"].str.contains("合计")]
            print(df.head().to_markdown())
            if len(df) > 0:
                plat = "JD"
                df1 = pd.DataFrame()
                df1["TID"] = df["订单编号"].apply(lambda x: x.replace(" ", "").strip())
                df1["SHOPNAME"] = ""
                df1["PLATFORM"] = plat
                df1["SHOPCODE"] = ""
                df1["BILLPLATFORM"] = "JDOVERSEAS"
                df1["CREATED"] = df["计费时间"]
                df1["TITLE"] = df["商品名称"]
                df1["TRADE_TYPE"] = "退款"
                df1["BUSINESS_NO"] = df["返修单号"]
                df1["INCOME_AMOUNT"] = 0
                df1["EXPEND_AMOUNT"] = df["退款金额"].astype(float) * df["汇率(美元/人民币)"].astype(float)
                df1["TRADING_CHANNELS"] = ""
                df1["BUSINESS_DESCRIPTION"] = "退款"
                df1["remark"] = ""
                df1["BUSINESS_BILL_SOURCE"] = ""
                df1["IS_REFUNDAMOUNT"] = 1
                df1["IS_AMOUNT"] = 0
                df1["OID"] = ""
                df1["SOURCEDATA"] = "EXCEL"
                df1["RECIPROCAL_ACCOUNT"] = ""
                df1["BATCHNO"] = ""
                df1["currency"] = "USD"
                df1["overseas_income"] = 0
                df1["overseas_expend"] = df["退款金额"]
                df1["currency_cny_rate"] = df["汇率(美元/人民币)"]

                df2 = df1.copy()
                df2["TRADE_TYPE"] = "佣金"
                df2["INCOME_AMOUNT"] = df["返还佣金"].astype(float) * df["汇率(美元/人民币)"].astype(float)
                df2["EXPEND_AMOUNT"] = 0
                df2["BUSINESS_DESCRIPTION"] = "佣金"
                df2["IS_REFUNDAMOUNT"] = 0
                df2["overseas_income"] = df["返还佣金"]
                df2["overseas_expend"] = 0

                dfs = [df1,df2]
                df1 = pd.concat(dfs)

                if filename.find("Fit海外") >= 0:
                    df1["SHOPCODE"] = "dicoraurbanfitov"
                    df1["SHOPNAME"] = "dicoraurbanfit海外旗舰店"
                elif filename.find("mades海外") >= 0:
                    df1["SHOPCODE"] = "madesovqjd"
                    df1["SHOPNAME"] = "mades海外旗舰店"
                elif filename.find("MET海外") >= 0:
                    df1["SHOPCODE"] = "manukascosmetov"
                    df1["SHOPNAME"] = "manukascosmet海外旗舰店"
                elif filename.find("rai海外") >= 0:
                    df1["SHOPCODE"] = "samouraiov"
                    df1["SHOPNAME"] = "samourai海外旗舰店"
                elif ((filename.find("image海外") >= 0) | (filename.find("wiss海外") >= 0)):
                    df1["SHOPCODE"] = "swissimageov"
                    df1["SHOPNAME"] = "swissimage海外旗舰店"
                elif ((filename.find("icora") >= 0) & (filename.find("海外") >= 0)):
                    df1["SHOPCODE"] = "dicoraurbanfitov"
                    df1["SHOPNAME"] = "dicoraurbanfit海外旗舰店"
                elif filename.find("LILAC海外") >= 0:
                    df1["SHOPCODE"] = "lilacovqjd"
                    df1["SHOPNAME"] = "lilac海外旗舰店"
                elif filename.find("AMBRA") >= 0:
                    df1["SHOPCODE"] = "ambrajdqjdov"
                    df1["SHOPNAME"] = "AMBRA京东海外旗舰店"
                elif ((filename.find("MANUKA") >= 0) & (filename.find("海外") >= 0)):
                    df1["SHOPCODE"] = "manukascosmetov"
                    df1["SHOPNAME"] = "manukascosmet海外旗舰店"
                print(df1.head().to_markdown())

                return df1

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

        elif filename.find("妥投销货清单") >= 0:
            if filename.find("csv")>=0:
                df = pd.read_csv(filename, skiprows=1, dtype=str, encoding="gb18030")
            else:
                df = pd.read_excel(filename, dtype=str)
            for column_name in df.columns:
                df.rename(columns={column_name: column_name.replace(" ", "").replace("\n", "").strip()},
                          inplace=True)
            df.dropna(subset=["订单编号"],inplace=True)
            df = df.loc[df["结算金额"]!="合计"]

            print(df.head(1).to_markdown())
            df["商品应结金额"] = df["商品应结金额"].astype(float)
            df["商品佣金"] = df["商品佣金"].astype(float)
            df["汇率(美元/人民币)"] = df["汇率(美元/人民币)"].astype(float)
            df["订单编号"] = df["订单编号"].astype(str)

            yearmonth = "".join("".join(filename.split("结算单")[1:]).split(os.sep)[:1])
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
            print(bill_time)

            plat = "JD"

            # 货款
            df1 = pd.DataFrame()
            df1["TID"] = df["订单编号"].apply(lambda x: x.replace(" ", "").strip())
            df1["SHOPNAME"] = ""
            df1["PLATFORM"] = plat
            df1["SHOPCODE"] = ""
            df1["BILLPLATFORM"] = "JDOVERSEAS"
            df1["CREATED"] = bill_time
            df1["TITLE"] = ""
            df1["TRADE_TYPE"] = "货款"
            df1["BUSINESS_NO"] = ""
            df1["INCOME_AMOUNT"] = df["商品应结金额"] * df["汇率(美元/人民币)"]
            df1["EXPEND_AMOUNT"] = 0
            df1["TRADING_CHANNELS"] = ""
            df1["BUSINESS_DESCRIPTION"] = "货款"
            df1["remark"] = ""
            df1["BUSINESS_BILL_SOURCE"] = ""
            df1["IS_REFUNDAMOUNT"] = 0
            df1["IS_AMOUNT"] = 1
            df1["OID"] = ""
            df1["SOURCEDATA"] = "EXCEL"
            df1["RECIPROCAL_ACCOUNT"] = ""
            df1["BATCHNO"] = ""
            df1["currency"] = "USD"
            df1["overseas_income"] = df["商品应结金额"]
            df1["overseas_expend"] = 0
            df1["currency_cny_rate"] = df["汇率(美元/人民币)"]

            if filename.find("Fit海外") >= 0:
                df1["SHOPCODE"] = "dicoraurbanfitov"
                df1["SHOPNAME"] = "dicoraurbanfit海外旗舰店"
            elif filename.find("mades海外") >= 0:
                df1["SHOPCODE"] = "madesovqjd"
                df1["SHOPNAME"] = "mades海外旗舰店"
            elif filename.find("MET海外") >= 0:
                df1["SHOPCODE"] = "manukascosmetov"
                df1["SHOPNAME"] = "manukascosmet海外旗舰店"
            elif filename.find("rai海外") >= 0:
                df1["SHOPCODE"] = "samouraiov"
                df1["SHOPNAME"] = "samourai海外旗舰店"
            elif ((filename.find("image海外") >= 0) | (filename.find("wiss海外") >= 0)):
                df1["SHOPCODE"] = "swissimageov"
                df1["SHOPNAME"] = "swissimage海外旗舰店"
            elif ((filename.find("icora") >= 0) & (filename.find("海外") >= 0)):
                df1["SHOPCODE"] = "dicoraurbanfitov"
                df1["SHOPNAME"] = "dicoraurbanfit海外旗舰店"
            elif filename.find("LILAC海外") >= 0:
                df1["SHOPCODE"] = "lilacovqjd"
                df1["SHOPNAME"] = "lilac海外旗舰店"
            elif filename.find("AMBRA") >= 0:
                df1["SHOPCODE"] = "ambrajdqjdov"
                df1["SHOPNAME"] = "AMBRA京东海外旗舰店"
            elif ((filename.find("MANUKA") >= 0) & (filename.find("海外") >= 0)):
                df1["SHOPCODE"] = "manukascosmetov"
                df1["SHOPNAME"] = "manukascosmet海外旗舰店"

            # 佣金
            df2 = df1.copy()
            df2["TRADE_TYPE"] = "佣金"
            df2["INCOME_AMOUNT"] = 0
            df2["EXPEND_AMOUNT"] = df["商品佣金"] * df["汇率(美元/人民币)"]
            df2["BUSINESS_DESCRIPTION"] = "佣金"
            df2["IS_REFUNDAMOUNT"] = 0
            df2["IS_AMOUNT"] = 0
            df2["overseas_income"] = 0
            df2["overseas_expend"] = df["商品佣金"]
            df2["currency_cny_rate"] = df["汇率(美元/人民币)"]

            dfs = [df1, df2]
            df1 = pd.concat(dfs)
            print(df1.head().to_markdown())

            return df1

        else:
            try:
                if filename.find("xls")>=0:
                    df = pd.read_excel(filename, sheet_name=None, dtype=str)
                else:
                    df = pd.read_csv(filename, sheet_name=None, dtype=str)
                sheet_list = list(df)
                print(sheet_list)
                if len(sheet_list) > 1:
                    df = None
                    for sheet in sheet_list:
                        if sheet.find("结算单")>=0:
                            if filename.find("xls") >= 0:
                                print(sheet)
                                df1 = pd.read_excel(filename, sheet_name=sheet, dtype=str)
                            else:
                                df1 = pd.read_csv(filename, sheet_name=sheet, dtype=str)

                            # 结算金额
                            plat = "JD"
                            df1["TID"] = df["订单编号"].apply(lambda x: x.replace(" ", "").strip())
                            df1["SHOPNAME"] = ""
                            df1["PLATFORM"] = plat
                            df1["SHOPCODE"] = ""
                            df1["BILLPLATFORM"] = "JDOVERSEAS"
                            df1["CREATED"] = df["完成时间"]
                            df1["TITLE"] = ""
                            df1["TRADE_TYPE"] = "货款"
                            df1["BUSINESS_NO"] = ""
                            df1["INCOME_AMOUNT"] = df["货款"] * df["汇率(美元/人民币)"]
                            df1["EXPEND_AMOUNT"] = 0
                            df1["TRADING_CHANNELS"] = ""
                            df1["BUSINESS_DESCRIPTION"] = "货款"
                            df1["remark"] = ""
                            df1["BUSINESS_BILL_SOURCE"] = ""
                            df1["IS_REFUNDAMOUNT"] = 0
                            df1["IS_AMOUNT"] = 1
                            df1["OID"] = ""
                            df1["SOURCEDATA"] = "EXCEL"
                            df1["RECIPROCAL_ACCOUNT"] = ""
                            df1["BATCHNO"] = ""
                            df1["currency"] = "USD"
                            df1["overseas_income"] = df["货款"]
                            df1["overseas_expend"] = 0
                            df1["currency_cny_rate"] = df["汇率(美元/人民币)"]

                            if filename.find("Ambra京东海外旗舰店") >= 0:
                                df1["SHOPCODE"] = "ambrajdqjdov"
                                df1["SHOPNAME"] = "AMBRA京东海外旗舰店"
                            elif filename.find("mades海外") >= 0:
                                df1["SHOPCODE"] = "madesovqjd"
                                df1["SHOPNAME"] = "mades海外旗舰店"

                            df1["sheet"] = sheet
                            if df is None:
                                df = df1
                            else:
                                print(f"df1:\n{df.head(1).to_markdown()}")
                                print(f"df:\n{df1.head(1).to_markdown()}")
                                df = pd.concat([df, df1])
                                print(len(df))
                        else:
                            pass
                    return df
                else:
                    sheet = sheet_list[0]
                    print(sheet)
                    if filename.find("xls") >= 0:
                        df = pd.read_excel(filename, sheet_name=sheet, dtype=str)
                    else:
                        df = pd.read_csv(filename, sheet_name=sheet, dtype=str)
                    df["sheet"] = sheet
                    return df
            except Exception as e:
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


def read_bill2(filename):
    print(filename)
    if ((filename.find("海外") >= 0)|(filename.find("AMBRA") >= 0)):
        if filename.find("非销售单结算") >= 0:
            df = pd.read_csv(filename, skiprows=1, dtype=str, encoding="gb18030")
            for column_name in df.columns:
                df.rename(columns={column_name: column_name.replace(" ", "").replace("\n", "").strip()},inplace=True)
            df = df[~df["费用类型"].str.contains("合计")]
            if len(df) > 0:
                plat = "JD"
                df["费用金额"] = df["费用金额"].astype(float)
                df["汇率(美元/人民币)"] = df["汇率(美元/人民币)"].astype(float)
                df["订单编号"] = df["订单编号"].astype(str)

                df1 = pd.DataFrame()
                df1["TID"] = df["订单编号"].apply(lambda x: x.replace(" ", "").strip())
                df1["SHOPNAME"] = ""
                df1["PLATFORM"] = plat
                df1["SHOPCODE"] = ""
                df1["BILLPLATFORM"] = "JDOVERSEAS"
                df1["CREATED"] = df["计费时间"]
                df1["TITLE"] = ""
                df1["TRADE_TYPE"] = df["费用类型"]
                df1["BUSINESS_NO"] = df["单据编号"]
                df1["INCOME_AMOUNT"] = df.apply(lambda x: x["费用金额"] * x["汇率(美元/人民币)"] if x["费用金额"] > 0 else 0, axis=1)
                df1["EXPEND_AMOUNT"] = df.apply(lambda x: x["费用金额"] * x["汇率(美元/人民币)"] if x["费用金额"] < 0 else 0, axis=1)
                df1["TRADING_CHANNELS"] = ""
                df1["BUSINESS_DESCRIPTION"] = df["费用类型"]
                df1["remark"] = ""
                df1["BUSINESS_BILL_SOURCE"] = ""
                df1["IS_REFUNDAMOUNT"] = 0
                df1["IS_AMOUNT"] = 0
                df1["OID"] = ""
                df1["SOURCEDATA"] = "EXCEL"
                df1["RECIPROCAL_ACCOUNT"] = ""
                df1["BATCHNO"] = ""
                df1["currency"] = "USD"
                df1["overseas_income"] = df.apply(lambda x: x["费用金额"] if x["费用金额"] > 0 else 0, axis=1)
                df1["overseas_expend"] = df.apply(lambda x: x["费用金额"] if x["费用金额"] < 0 else 0, axis=1)
                df1["currency_cny_rate"] = df["汇率(美元/人民币)"]

                if filename.find("Fit海外") >= 0:
                    df1["SHOPCODE"] = "dicoraurbanfitov"
                    df1["SHOPNAME"] = "dicoraurbanfit海外旗舰店"
                elif filename.find("mades海外") >= 0:
                    df1["SHOPCODE"] = "madesovqjd"
                    df1["SHOPNAME"] = "mades海外旗舰店"
                elif filename.find("MET海外") >= 0:
                    df1["SHOPCODE"] = "manukascosmetov"
                    df1["SHOPNAME"] = "manukascosmet海外旗舰店"
                elif filename.find("rai海外") >= 0:
                    df1["SHOPCODE"] = "samouraiov"
                    df1["SHOPNAME"] = "samourai海外旗舰店"
                elif ((filename.find("image海外") >= 0) | (filename.find("wiss海外") >= 0)):
                    df1["SHOPCODE"] = "swissimageov"
                    df1["SHOPNAME"] = "swissimage海外旗舰店"
                elif ((filename.find("icora") >= 0) & (filename.find("海外") >= 0)):
                    df1["SHOPCODE"] = "dicoraurbanfitov"
                    df1["SHOPNAME"] = "dicoraurbanfit海外旗舰店"
                elif filename.find("LILAC海外") >= 0:
                    df1["SHOPCODE"] = "lilacovqjd"
                    df1["SHOPNAME"] = "lilac海外旗舰店"
                elif filename.find("AMBRA") >= 0:
                    df1["SHOPCODE"] = "ambrajdqjdov"
                    df1["SHOPNAME"] = "AMBRA京东海外旗舰店"
                elif ((filename.find("MANUKA") >= 0) & (filename.find("海外") >= 0)):
                    df1["SHOPCODE"] = "manukascosmetov"
                    df1["SHOPNAME"] = "manukascosmet海外旗舰店"

                return df1

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
    # del table2["filename"]

    tables = [table1,table2]
    table = pd.concat(tables)
    print(f"合并表行数：{len(table)}")

    # if table.shape[0] < 800000:
    #     table.to_excel(default_dir + "/处理后的账单.xlsx", index=False)
    # else:
    #     table.to_csv(default_dir + "/处理后的账单.csv", index=False)
    index = 0
    if "TID" in table.columns:
        table["TID"] = table["TID"].astype(str)
        # table["TID"] = table["TID"].apply(lambda x:x.str.replace(" ",np.nan).str.replace("\n",np.nan))
        table.replace("nan", np.nan, inplace=True)
        table.dropna(subset=["TID"], axis=0, inplace=True)
        # table.drop_duplicates(inplace=True)
        table = table.sort_values(by=["TID", "CREATED"])
        # table["TID"] = table["TID"].astype(str)
    table = table.loc[~((table.overseas_income == 0) & (table.overseas_expend == 0))]
    table["CREATED"] = table["CREATED"].astype("datetime64[ns]")
    table["currency_cny_rate"] = table["currency_cny_rate"].astype(float)
    table["INCOME_AMOUNT"] = table["INCOME_AMOUNT"].astype(float)
    table["EXPEND_AMOUNT"] = table["EXPEND_AMOUNT"].astype(float)
    table["time"] = pd.to_datetime(table["CREATED"]).dt.date
    table["month"] = pd.to_datetime(table["CREATED"]).dt.month
    table["year"] = pd.to_datetime(table["CREATED"]).dt.year
    tabletl = table.loc[table["currency_cny_rate"]>0][["month","currency_cny_rate"]]
    tabletl.drop_duplicates(subset=["month"],inplace=True)
    tabletl.columns = ["month","rate1"]
    tabletlr = table.loc[table["currency_cny_rate"] > 0][["year", "currency_cny_rate"]]
    tabletlr.drop_duplicates(subset=["year"], inplace=True)
    tabletlr.columns = ["year", "rate2"]
    # table.to_excel(default_dir + "\处理后的账单.xlsx")
    tablet = table.groupby(["time"]).agg({"currency_cny_rate":"max"})
    tablet = pd.DataFrame(tablet).reset_index()
    # tablet["month"] = pd.to_datetime(tablet["time"]).dt.month
    tablet.columns = ["time","rate"]
    table = pd.merge(table,tablet,how="left",on="time")
    table = pd.merge(table, tabletl, how="left", on="month")
    table = pd.merge(table, tabletlr, how="left", on="year")
    print(table.head().to_markdown())
    table["currency_cny_rate"] = table.apply(lambda x:x["currency_cny_rate"] if x["currency_cny_rate"]!=0 else x["rate"],axis=1)
    table["currency_cny_rate"] = table.apply(lambda x:x["rate1"] if ((x["currency_cny_rate"]==0)&(x["rate"]==0)) else x["currency_cny_rate"],axis=1)
    table["currency_cny_rate"] = table.apply(lambda x:x["currency_cny_rate"] if x["currency_cny_rate"]>0 else x["rate2"],axis=1)
    table["INCOME_AMOUNT"] = table.apply(lambda x:x["overseas_income"]*x["currency_cny_rate"] if x["INCOME_AMOUNT"] == 0 else x["INCOME_AMOUNT"],axis=1)
    table["EXPEND_AMOUNT"] = table.apply(lambda x:x["overseas_expend"]*x["currency_cny_rate"] if x["EXPEND_AMOUNT"] == 0 else x["EXPEND_AMOUNT"],axis=1)
    table["CREATED"] = table["CREATED"].astype("datetime64[ns]")
    table["INCOME_AMOUNT"] = table["INCOME_AMOUNT"].astype(float).abs()
    table["EXPEND_AMOUNT"] = -table["EXPEND_AMOUNT"].astype(float).abs()
    table["overseas_income"] = table["overseas_income"].astype(float).abs()
    table["overseas_expend"] = -table["overseas_expend"].astype(float).abs()
    del table["time"]
    del table["month"]
    del table["year"]
    del table["rate"]
    del table["rate1"]
    del table["rate2"]

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