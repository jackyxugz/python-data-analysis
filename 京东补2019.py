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
    # 京东逻辑
    if filename.find("京东") >= 0:
        df = pd.read_excel(filename, sheet_name=None, dtype=str)
        sheet_list = list(df)
        print(sheet_list)
        if len(sheet_list) > 1:
            df = None
            for sheet in sheet_list:
                if sheet.find("结算单") >= 0:
                    print(sheet)
                    df1 = pd.read_excel(filename, sheet_name=sheet, dtype=str)
                    print(df1.head().to_markdown())
                    # print(df1["订单编号"].sort_values())
                    df1["订单编号"] = df1["订单编号"].astype(str)
                    df1["订单编号"] = df1["订单编号"].apply(lambda x:x if len(x)>3 else "nan")

                    df1 = df1[~df1["订单编号"].str.contains("nan")]
                    plat = "JD"
                    df2 = pd.DataFrame()
                    df2["TID"] = df1["订单编号"]
                    df2["SHOPNAME"] = ""
                    df2["PLATFORM"] = plat
                    df2["SHOPCODE"] = ""
                    df2["BILLPLATFORM"] = "JDOVERSEAS"
                    df2["CREATED"] = df1["完成时间"]
                    df2["TITLE"] = ""
                    df2["TRADE_TYPE"] = "货款"
                    df2["BUSINESS_NO"] = ""
                    df2["INCOME_AMOUNT"] = df1["商品应结金额1"]
                    df2["EXPEND_AMOUNT"] = 0
                    df2["TRADING_CHANNELS"] = ""
                    df2["BUSINESS_DESCRIPTION"] = "货款"
                    df2["remark"] = ""
                    df2["BUSINESS_BILL_SOURCE"] = ""
                    df2["IS_REFUNDAMOUNT"] = 0
                    df2["IS_AMOUNT"] = 1
                    df2["OID"] = ""
                    df2["SOURCEDATA"] = "EXCEL"
                    df2["RECIPROCAL_ACCOUNT"] = ""
                    df2["BATCHNO"] = ""
                    df2["currency"] = "USD"
                    df2["overseas_income"] = df1["结算金额"]
                    df2["overseas_expend"] = 0
                    df2["currency_cny_rate"] = df1["汇率(美元/人民币)"]

                    if filename.find("Fit海外") >= 0:
                        df2["SHOPCODE"] = "dicoraurbanfitov"
                        df2["SHOPNAME"] = "dicoraurbanfit海外旗舰店"
                    elif filename.find("mades海外") >= 0:
                        df2["SHOPCODE"] = "madesovqjd"
                        df2["SHOPNAME"] = "mades海外旗舰店"

                    # 佣金
                    df3 = pd.DataFrame()
                    df3 = df1.copy()
                    df3["TRADE_TYPE"] = "佣金"
                    df3["INCOME_AMOUNT"] = 0
                    df3["EXPEND_AMOUNT"] = df1["商品佣金1"]
                    df3["BUSINESS_DESCRIPTION"] = "佣金"
                    df3["IS_REFUNDAMOUNT"] = 0
                    df3["IS_AMOUNT"] = 0
                    df3["overseas_income"] = 0
                    df3["overseas_expend"] = df1["商品佣金"]
                    # df1["sheet"] = sheet

                    df1 = pd.concat([df2,df3])

                    if df is None:
                        df = df1
                    else:
                        print(f"df1:\n{df.head(1).to_markdown()}")
                        print(f"df:\n{df1.head(1).to_markdown()}")
                        df = pd.concat([df, df1])
                        print(len(df))
                elif ((sheet.find("退款")>= 0) or (sheet.find("退货")>= 0)):
                    print(sheet)
                    df1 = pd.read_excel(filename, sheet_name=sheet, dtype=str)
                    plat = "JD"
                    df2 = pd.DataFrame()
                    df2["TID"] = df1["订单编号"]
                    df2["SHOPNAME"] = ""
                    df2["PLATFORM"] = plat
                    df2["SHOPCODE"] = ""
                    df2["BILLPLATFORM"] = "JDOVERSEAS"
                    df2["CREATED"] = df1["退货时间"]
                    df2["TITLE"] = ""
                    df2["TRADE_TYPE"] = "退款"
                    df2["BUSINESS_NO"] = ""
                    df2["INCOME_AMOUNT"] = 0
                    df2["EXPEND_AMOUNT"] = df1["退款金额"].astype(float) * df1["汇率(美元/人民币)"].astype(float)
                    df2["TRADING_CHANNELS"] = ""
                    df2["BUSINESS_DESCRIPTION"] = "退款"
                    df2["remark"] = ""
                    df2["BUSINESS_BILL_SOURCE"] = ""
                    df2["IS_REFUNDAMOUNT"] = 1
                    df2["IS_AMOUNT"] = 0
                    df2["OID"] = ""
                    df2["SOURCEDATA"] = "EXCEL"
                    df2["RECIPROCAL_ACCOUNT"] = ""
                    df2["BATCHNO"] = ""
                    df2["currency"] = "USD"
                    df2["overseas_income"] = 0
                    df2["overseas_expend"] = df1["退款金额"]
                    df2["currency_cny_rate"] = df1["汇率(美元/人民币)"]

                    if filename.find("Fit海外") >= 0:
                        df2["SHOPCODE"] = "dicoraurbanfitov"
                        df2["SHOPNAME"] = "dicoraurbanfit海外旗舰店"
                    elif filename.find("mades海外") >= 0:
                        df2["SHOPCODE"] = "madesovqjd"
                        df2["SHOPNAME"] = "mades海外旗舰店"

                    df1 = df2.copy()

                    if df is None:
                        df = df1
                    else:
                        print(f"df1:\n{df.head(1).to_markdown()}")
                        print(f"df:\n{df1.head(1).to_markdown()}")
                        df = pd.concat([df, df1])
                        print(len(df))
                else:
                    pass
            print(df.head().to_markdown())
            df["INCOME_AMOUNT"].fillna(0,inplace=True)
            df["EXPEND_AMOUNT"].fillna(0,inplace=True)
            df = df.loc[~((df.INCOME_AMOUNT == 0) & (df.EXPEND_AMOUNT == 0))]
            return df
        else:
            sheet = sheet_list[0]
            print(sheet)
            df = pd.read_excel(filename, sheet_name=sheet, dtype=str)
            df["sheet"] = sheet
            return df
    else:
        dict = {"TID": "nan", "SHOPNAME": "", "PLATFORM": "", "SHOPCODE": "", "BILLPLATFORM": "", "CREATED": "",
                "TITLE": "", "TRADE_TYPE": "", "BUSINESS_NO": "", "INCOME_AMOUNT": "", "EXPEND_AMOUNT": "",
                "TRADING_CHANNELS": "", "BUSINESS_DESCRIPTION": "", "remark": "", "BUSINESS_BILL_SOURCE": "",
                "IS_REFUNDAMOUNT": "", "IS_AMOUNT": "", "OID": "", "SOURCEDATA": "", "RECIPROCAL_ACCOUNT": "",
                "BATCHNO": ""}
        df = pd.DataFrame(dict,index=[0])
        return df

def get_baidushop(id, filename, type):
    df = pd.read_excel("data/百度-商品ID与店铺名称关系0(3)(2).xlsx", dtype=str)
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

    # df1 = read_all_excel(filedir, filekey)
    df1 = read_all_bill(filedir, filekey)
    df1["SHOPNAME"] = df1["SHOPNAME"].astype(str)
    df1["SHOPNAME"] = df1.apply(lambda x: x["SHOPNAME"] if len(x["SHOPNAME"]) > 2 else x["filename"], axis=1)
    del df1["filename"]

    # if df1.shape[0] < 800000:
    #     df1.to_excel(default_dir + "/处理后的账单.xlsx", index=False)
    # else:
    #     df1.to_csv(default_dir + "/处理后的账单.csv", index=False)
    index = 0
    # if "TID" in df1.columns:
    if "PDD" in df1["PLATFORM"].values.tolist():
        df1.replace("nan", np.nan, inplace=True)
    # if "KAOLA" in df1["PLATFORM"].values.tolist():
    #     df1 = df1.loc[~((df1.INCOME_AMOUNT == 0) & (df1.EXPEND_AMOUNT == 0))]
    # df1 = pd.DataFrame(df1).reset_index()
    # df1["index"] = df1.index + 1
    # df1["TID"] = "OF" + df1["SHOPCODE"] + df1["CREATED"].str.replace("-","") + df1["index"].map(lambda x: "{:0>6d}".format(x))
    # del df1["index"]
    else:
        df1["TID"] = df1["TID"].astype(str)
        # df1["TID"] = df1["TID"].apply(lambda x:x.str.replace(" ",np.nan).str.replace("\n",np.nan))
        df1.replace("nan", np.nan, inplace=True)
        df1.dropna(subset=["TID"], inplace=True)
        # df1.drop_duplicates(inplace=True)
    df1 = df1.sort_values(by=["PLATFORM", "SHOPNAME", "TID", "CREATED"])
    df1 = df1.loc[~((df1.INCOME_AMOUNT == 0) & (df1.EXPEND_AMOUNT == 0))]
    # print(df1.head().to_markdown())
    # df1["SHOPNAME"] = df1.apply(lambda x: x["SHOPNAME"] if pd.notnull(x["SHOPNAME"]) else x["filename"], axis=1)
    df1["CREATED"] = df1["CREATED"].astype("datetime64[ns]")
    plat = os.sep.join(default_dir.split(os.sep)[-1:])
    print("第{}个表格,记录数:{}".format(index, df1.shape[0]))
    print(df1.head(10).to_markdown())
    # df.to_excel(r"work/合并表格_test.xlsx")
    print("账单总行数：")
    print(df1.shape[0])
    for i in range(0, int(df1.shape[0] / 200000) + 1):
        print("存储分页：{}  from:{} to:{}".format(i, i * 200000, (i + 1) * 200000))
        # df.iloc[i*500000:(i+1)*500000].to_csv(default_dir + "\{}_合并表格_{}.csv".format(index,i))
        df1.iloc[i * 200000:(i + 1) * 200000].to_excel(default_dir + "\{}-处理后的账单{}.xlsx".format(plat, i), index=False)

    return df1


if __name__ == "__main__":
    print("开始:", time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()))
    # print("abc".find("x"))

    # combine_excel()

    combine_bill()

    # groupby_amt()
    # math_file()
    # get_shopcode("JD","dentylactive旗舰店")

    print("ok")