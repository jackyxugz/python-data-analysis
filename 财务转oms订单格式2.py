# __coding=utf8__
# /** 作者：zengyanghui **/

import pandas as pd
import numpy as np
from collections import Counter
import re
import os
import warnings
import time
import sys
import math
import datetime as dt
import uuid
import tabulate
import os.path
import xlrd
import xlwt
import xlsxwriter

warnings.filterwarnings("ignore")


# import Tkinter
# import win32api
# import win32ui
# import win32con
# import win32com
# from win32com.shell import shell
# import json


# 检查目录是否存在
def mkdir(default_path):
    path1 = default_path + "/数据转换"

    isExists1 = os.path.exists(path1)
    if not isExists1:
        os.makedirs(path1)
        print(path1 + ' 创建成功')
    else:
        print(path1 + ' 目录已存在')


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


def read_order(filename):
    print(filename)
    df = pd.read_excel(filename, dtype=str)
    # 给文件增加平台字段
    df["平台"] = ""
    plat = os.path.basename(filename)
    if ((filename.find("淘宝") > 0) and (filename.count("淘宝") > filename.count("天猫"))):
        plat = "淘宝"
    elif ((filename.find("淘宝") > 0) and (filename.count("淘宝") < filename.count("天猫"))):
        plat = "天猫"
    elif filename.find("京东") > 0:
        plat = "京东"
    elif filename.find("拼多多") > 0:
        plat = "拼多多"
    elif filename.find("抖音") > 0:
        plat = "抖音"
    elif filename.find("快手") > 0:
        plat = "快手"
    elif filename.find("小红书") > 0:
        plat = "小红书"
    elif filename.find("考拉") > 0:
        plat = "网易考拉"
    elif filename.find("有赞") > 0:
        plat = "有赞"
    df["平台"] = plat

    # 淘宝/天猫的平台数据预处理
    if df[df["平台"].isin(["淘宝", "天猫"])].shape[0] > 0:
        # print(df["平台"].value_counts())
        if "订单编号" not in df.columns:  # 如果没有订单编号，则说明表格不是订单数据，跳过处理
            pass
        else:
            df1 = pd.read_excel(filename, sheet_name="属性原表", dtype=str)  # 读取属性原表的数据，并与订单原表合并
            dfs = pd.merge(df, df1[["订单编号", "标题", "商家编码"]], how="left", on="订单编号")
            df = pd.DataFrame(dfs)
            if "宝贝标题" in df.columns and "标题" in df.columns:
                df["宝贝标题"] = df.apply(lambda x: x["标题"] if pd.isnull(x["宝贝标题"]) else x["宝贝标题"], axis=1)
                del df["标题"]

    # 京东的平台数据预处理
    elif df[df["平台"].isin(["京东"])].shape[0] > 0:
        # print(df["平台"].value_counts())
        express_file = filename[:-5] + "快递信息.xlsx"  # 对文件名进行修改
        df1 = pd.read_excel(express_file, dtype=str)
        dfs = pd.merge(df, df1[["订单号", "快递公司", "快递单号"]], how="left", on="订单号")
        df = pd.DataFrame(dfs)

    # 有赞的平台数据预处理
    elif df[df["平台"].isin(["有赞"])].shape[0] > 0:
        # print(df["平台"].value_counts())
        express_file = filename.replace("订单20", "商品20")  # 对文件名进行修改
        print(express_file)
        df1 = pd.read_excel(express_file, dtype=str)
        dfs = pd.merge(df, df1[["订单号", "商品名称", "规格编码", "商品单价", "商品数量", "商品发货物流公司", "商品发货物流单号"]], how="left", on="订单号")
        df = pd.DataFrame(dfs)

    # 去掉表头空格与换行符，以及通用字段重命名
    for column_name in df.columns:
        df.rename(columns={column_name: column_name.replace(" ", "").replace("\n", "").strip()}, inplace=True)
        # 订单编号重命名
        if column_name == "子订单编号":
            df.rename(columns={"子订单编号": "订单编号"}, inplace=True)
        elif column_name == "订单号":
            df.rename(columns={"订单号": "订单编号"}, inplace=True)
        # 店铺重命名
        elif column_name == "店铺名称":
            df.rename(columns={"店铺名称": "店铺"}, inplace=True)
        elif column_name == "店铺号":
            df.rename(columns={"店铺号": "店铺"}, inplace=True)
        elif column_name == "卖家名称":
            df.rename(columns={"卖家名称": "店铺"}, inplace=True)
        elif column_name == "归属店铺":
            df.rename(columns={"归属店铺": "店铺"}, inplace=True)
        # 商品名称重命名
        elif column_name == "标题":
            df.rename(columns={"标题": "商品名称"}, inplace=True)
        elif column_name == "宝贝标题":
            df.rename(columns={"宝贝标题": "商品名称"}, inplace=True)
        elif column_name == "商品":
            df.rename(columns={"商品": "商品名称"}, inplace=True)
        elif column_name == "商品款号":
            df.rename(columns={"商品款号": "商品名称"}, inplace=True)
        elif column_name == "选购商品":
            df.rename(columns={"选购商品": "商品名称"}, inplace=True)
        elif column_name == "选购产品":
            df.rename(columns={"选购产品": "商品名称"}, inplace=True)
        elif column_name == "商品名":
            df.rename(columns={"商品名": "商品名称"}, inplace=True)
        elif column_name == "货品标题":
            df.rename(columns={"货品标题": "商品名称"}, inplace=True)
        # 购买数量重命名
        elif column_name == "商品数量":
            df.rename(columns={"商品数量": "购买数量"}, inplace=True)
        elif column_name == "订购数量":
            df.rename(columns={"订购数量": "购买数量"}, inplace=True)
        elif column_name == "数量":
            df.rename(columns={"数量": "购买数量"}, inplace=True)
        elif column_name == "商品数量(件)":
            df.rename(columns={"商品数量(件)": "购买数量"}, inplace=True)
        elif column_name == "宝贝总数量":
            df.rename(columns={"宝贝总数量": "购买数量"}, inplace=True)
        elif column_name == "成交数量":
            df.rename(columns={"成交数量": "购买数量"}, inplace=True)
        # 订单状态重命名
        elif column_name == "状态":
            df.rename(columns={"状态": "订单状态"}, inplace=True)
        # 订单时间重命名
        elif column_name == "订单创建时间":
            df.rename(columns={"订单创建时间": "订单时间"}, inplace=True)
        elif column_name == "下单时间":
            df.rename(columns={"下单时间": "订单时间"}, inplace=True)
        elif column_name == "订单确认时间":
            df.rename(columns={"订单确认时间": "订单时间"}, inplace=True)
        elif column_name == "用户下单时间":
            df.rename(columns={"用户下单时间": "订单时间"}, inplace=True)
        elif column_name == "订单提交时间":
            df.rename(columns={"订单提交时间": "订单时间"}, inplace=True)
        elif column_name == "支付时间":
            df.rename(columns={"支付时间": "订单时间"}, inplace=True)
        # 支付方式重命名
        elif column_name == "支付详情":
            df.rename(columns={"支付详情": "支付方式"}, inplace=True)
        elif column_name == "订单类型":
            df.rename(columns={"订单类型": "支付方式"}, inplace=True)
        # 商家编码重命名
        elif column_name == "商家SKUID":
            df.rename(columns={"商家SKUID": "商家编码"}, inplace=True)
        elif column_name == "商品SKU":
            df.rename(columns={"商品SKU": "商家编码"}, inplace=True)
        elif column_name == "商家编码-SKU维度":
            df.rename(columns={"商家编码-SKU维度": "商家编码"}, inplace=True)
        elif column_name == "小红书编码":
            df.rename(columns={"小红书编码": "商家编码"}, inplace=True)
        elif column_name == "商品条形码":
            df.rename(columns={"商品条形码": "商家编码"}, inplace=True)
        elif column_name == "商品编码":
            df.rename(columns={"商品编码": "商家编码"}, inplace=True)
        elif column_name == "sku编码":
            df.rename(columns={"sku编码": "商家编码"}, inplace=True)
        elif column_name == "商家编号":
            df.rename(columns={"商家编号": "商家编码"}, inplace=True)
        elif column_name == "货号":
            df.rename(columns={"货号": "商家编码"}, inplace=True)
        elif column_name == "规格编码":
            df.rename(columns={"规格编码": "商家编码"}, inplace=True)
        # 销售金额重命名
        elif column_name == "商家实收金额(元)":
            df.rename(columns={"商家实收金额(元)": "销售金额"}, inplace=True)
        elif column_name == "售价单价":
            df.rename(columns={"售价单价": "销售金额"}, inplace=True)
        elif column_name == "商品单价":
            df.rename(columns={"商品单价": "销售金额"}, inplace=True)
        elif column_name == "结算金额":
            df.rename(columns={"结算金额": "销售金额"}, inplace=True)
        elif column_name == "结算单价":
            df.rename(columns={"结算单价": "销售金额"}, inplace=True)
        elif column_name == "结算总额":
            df.rename(columns={"结算总额": "销售金额"}, inplace=True)
        elif column_name == "订单实付金额":
            df.rename(columns={"订单实付金额": "销售金额"}, inplace=True)
        elif column_name == "客户实际支付金额（商品金额-商家优惠":
            df.rename(columns={"客户实际支付金额（商品金额-商家优惠": "销售金额"}, inplace=True)
        elif column_name == "买家实际支付金额":
            df.rename(columns={"买家实际支付金额": "销售金额"}, inplace=True)
        elif column_name == "实收款（到付按此收费）":
            df.rename(columns={"实收款（到付按此收费）": "销售金额"}, inplace=True)
        elif column_name == "总价":
            df.rename(columns={"总价": "销售金额"}, inplace=True)
        elif column_name == "实付款":
            df.rename(columns={"实付款": "销售金额"}, inplace=True)
        elif column_name == "实付款(元)":
            df.rename(columns={"实付款(元)": "销售金额"}, inplace=True)
        # 物流单号重命名
        elif column_name == "快递单号":
            df.rename(columns={"快递单号": "物流单号"}, inplace=True)
        elif column_name == "单号":
            df.rename(columns={"单号": "物流单号"}, inplace=True)
        elif column_name == "配送单号":
            df.rename(columns={"配送单号": "物流单号"}, inplace=True)
        elif column_name == "商品发货物流单号":
            df.rename(columns={"商品发货物流单号": "物流单号"}, inplace=True)
        # 物流公司重命名
        elif column_name == "快递公司":
            df.rename(columns={"快递公司": "物流公司"}, inplace=True)
        elif column_name == "配送公司":
            df.rename(columns={"配送公司": "物流公司"}, inplace=True)
        elif column_name == "商品发货物流公司":
            df.rename(columns={"商品发货物流公司": "物流公司"}, inplace=True)
        # 收货人姓名重命名
        elif column_name == "客户姓名":
            df.rename(columns={"客户姓名": "收货人姓名"}, inplace=True)
        elif column_name == "收件人":
            df.rename(columns={"收件人": "收货人姓名"}, inplace=True)
        elif column_name == "收货人":
            df.rename(columns={"收货人": "收货人姓名"}, inplace=True)
        elif column_name == "收货人/提货人":
            df.rename(columns={"收货人/提货人": "收货人姓名"}, inplace=True)
        # 收货地址重命名
        elif column_name == "客户地址":
            df.rename(columns={"客户地址": "收货地址"}, inplace=True)
        elif column_name == "收货人具体地址":
            df.rename(columns={"收货人具体地址": "收货地址"}, inplace=True)
        elif column_name == "收件地址":
            df.rename(columns={"收件地址": "收货地址"}, inplace=True)
        elif column_name == "完整地址":
            df.rename(columns={"完整地址": "收货地址"}, inplace=True)
        elif column_name == "收货地址":
            df.rename(columns={"收货地址": "收货地址"}, inplace=True)
        elif column_name == "详细地址":
            df.rename(columns={"详细地址": "收货地址"}, inplace=True)
        elif column_name == "详细收货地址/提货地址":
            df.rename(columns={"详细收货地址/提货地址": "收货地址"}, inplace=True)
        # 联系手机重命名
        elif column_name == "收货人电话":
            df.rename(columns={"收货人电话": "联系手机"}, inplace=True)
        elif column_name == "联系电话":
            df.rename(columns={"联系电话": "联系手机"}, inplace=True)
        elif column_name == "收货人电话":
            df.rename(columns={"收货人电话": "联系手机"}, inplace=True)
        elif column_name == "收件人手机号":
            df.rename(columns={"收件人手机号": "联系手机"}, inplace=True)
        elif column_name == "手机号":
            df.rename(columns={"手机号": "联系手机"}, inplace=True)
        elif column_name == "手机":
            df.rename(columns={"手机": "联系手机"}, inplace=True)
        elif column_name == "收货人手机号/提货人手机号":
            df.rename(columns={"收货人手机号/提货人手机号": "联系手机"}, inplace=True)
        # 买家会员名重命名
        elif column_name == "下单帐号":
            df.rename(columns={"下单帐号": "买家会员名"}, inplace=True)
        elif column_name == "用户id":
            df.rename(columns={"用户id": "买家会员名"}, inplace=True)
        elif column_name == "下单人id":
            df.rename(columns={"下单人id": "买家会员名"}, inplace=True)
        elif column_name == "客户昵称":
            df.rename(columns={"客户昵称": "买家会员名"}, inplace=True)
        # 销售单价
        elif column_name == "商品单价":
            df.rename(columns={"商品单价": "销售单价"}, inplace=True)
    # print(df.head(5).to_markdown())

    # 给文件增加通用字段
    if "主订单编号" in df.columns:
        pass
    else:
        df["主订单编号"] = ""
    if "商品名称" in df.columns:
        pass
    else:
        df["商品名称"] = ""
    if "商家编码" in df.columns:
        pass
    else:
        df["商家编码"] = ""
    if "收货人姓名" in df.columns:
        pass
    else:
        df["收货人姓名"] = ""
    if "收货地址" in df.columns:
        pass
    else:
        df["收货地址"] = ""
    if "联系手机" in df.columns:
        pass
    else:
        df["联系手机"] = ""
    if "买家会员名" in df.columns:
        pass
    else:
        df["买家会员名"] = ""
    if "物流公司" in df.columns:
        pass
    else:
        df["物流公司"] = ""
    if "物流单号" in df.columns:
        pass
    else:
        df["物流单号"] = ""
    if "支付方式" in df.columns:
        pass
    else:
        df["支付方式"] = ""
    if "销售单价" in df.columns:
        pass
    else:
        df["销售单价"] = ""
    # 店铺如果原表没有字段，则取文件夹最下层名称
    if "店铺" in df.columns:
        pass
    else:
        df["店铺"] = "".join(filename.split("/")[-2:-1])

    # 小红书销售金额处理
    if df[df["平台"].isin(["小红书"])].shape[0] > 0:
        df["SKU实付单价"] = df["SKU实付单价"].astype(float)
        df["SKU件数"] = df["SKU件数"].astype(float)
        df["销售金额"] = df.apply(lambda x: x["SKU实付单价"] * x["SKU件数"], axis=1)
        df["销售单价"] = df["SKU实付单价"]
        df["购买数量"] = df["SKU件数"]

    # 异常表格处理，新增模板字段
    if "订单编号" not in df.columns:
        # pass
        dict = {"数据来源": "", "平台": "", "订单编号": "", "主订单编号": "", "店铺": "", "出现序号": "", "总序号": "", "商品名称": "", "购买数量": "",
                "订单状态": "", "订单时间": "", "支付方式": "", "商家编码": "", "销售单价": "",
                "销售金额": "", "成本占比": "", "成本单价": "", "成本金额": "", "支付单号": "", "物流单号": "", "物流公司": "", "收货人姓名": "",
                "收货地址": "", "联系手机": "", "买家会员名": "", "退款金额": "",
                "退款金额（外币）": "", "是否回款": "", "结算方式": "", "回款金额": "", "回款日期": "", "回款差额": "", "回款名称": "", "回款币种": "",
                "回款汇率": "", "回款金额（外币）": "", "税金": "", "财务费用": "",
                "快递运费": "", "是否商务单": "", "商务单金额": "", "是否预售单": "", "是否菜鸟发货": "", "是否前海仓库发货": "", "仓库名称": ""}
        df = pd.DataFrame(dict, index=[0])
        return df
    else:
        df["iid"] = df.index
        df["数据来源"] = filename
        df["出现序号"] = df.groupby(df["订单编号"])["iid"].rank(method='dense')
        df_count = pd.DataFrame(df["订单编号"].value_counts()).reset_index()
        # print(df_count.to_markdown())
        df_count.columns = ["订单编号", "总序号"]
        df = df.merge(df_count, how="left", on="订单编号")
        df["成本占比"] = ""
        df["成本单价"] = ""
        df["成本金额"] = ""
        df["支付单号"] = ""
        df["退款金额"] = ""
        df["退款金额（外币）"] = ""
        df["是否回款"] = ""
        df["结算方式"] = ""
        df["回款金额"] = ""
        df["回款日期"] = ""
        df["回款差额"] = ""
        df["回款名称"] = ""
        df["回款币种"] = ""
        df["回款汇率"] = ""
        df["回款金额（外币）"] = ""
        df["税金"] = ""
        df["财务费用"] = ""
        df["快递运费"] = ""
        df["是否商务单"] = ""
        df["商务单金额"] = ""
        df["是否预售单"] = ""
        df["是否菜鸟发货"] = ""
        df["是否前海仓库发货"] = ""
        df["仓库名称"] = ""

        # 删除重复列名
        df = df.loc[:, ~df.columns.duplicated()]

        # 对列表重新排序，只保留目标字段
        df = df[
            ["数据来源", "平台", "订单编号", "主订单编号", "店铺", "出现序号", "总序号", "商品名称", "购买数量", "订单状态", "订单时间", "支付方式", "商家编码", "销售单价",
             "销售金额", "成本占比", "成本单价", "成本金额", "支付单号", "物流单号", "物流公司", "收货人姓名", "收货地址", "联系手机", "买家会员名", "退款金额",
             "退款金额（外币）", "是否回款", "结算方式", "回款金额", "回款日期", "回款差额", "回款名称", "回款币种", "回款汇率", "回款金额（外币）", "税金", "财务费用",
             "快递运费", "是否商务单", "商务单金额", "是否预售单", "是否菜鸟发货", "是否前海仓库发货", "仓库名称"]]

        df["购买数量"] = df["购买数量"].astype(float)
        df["销售金额"] = df["销售金额"].astype(float)
        print(df.head(5).to_markdown())
        # df.to_excel("data/temp.xlsx")

        return df


def read_bill(filename):
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
    if "xlsx" in filename:
        if filename.find("考拉") > 0:
            df = pd.read_excel(filename, sheet_name="销售明细", keep_default_na=False, dtype=str)
            df = df.apply(lambda x: x.astype(str).str.replace("=", "").str.replace("'", "").str.replace('"', ''))
        elif ((filename.find("淘宝") > 0) or (filename.find("天猫") > 0)):
            if filename.find("支付宝") > 0:
                df = pd.read_excel(filename, keep_default_na=False, dtype=str)
                df = df.apply(lambda x: x.astype(str).str.replace("=", "").str.replace("'", "").str.replace('"', ''))
            else:
                dict = {"订单编号": "", "退款金额": "", "退款金额（外币）": "", "是否回款": "", "结算方式": "", "回款金额": "", "回款日期": "",
                        "回款差额": "", "回款名称": "", "回款币种": "", "回款汇率": "", "回款金额（外币）": "", "税金": "", "财务费用": ""}
                df = pd.DataFrame(dict, index=[0])
                return df
        elif filename.find("小红书") > 0:
            dict = {"订单编号": "", "退款金额": "", "退款金额（外币）": "", "是否回款": "", "结算方式": "", "回款金额": "", "回款日期": "",
                    "回款差额": "", "回款名称": "", "回款币种": "", "回款汇率": "", "回款金额（外币）": "", "税金": "", "财务费用": ""}
            df = pd.DataFrame(dict, index=[0])
            return df
        else:
            df = pd.read_excel(filename, keep_default_na=False, dtype=str)
            df = df.apply(lambda x: x.astype(str).str.replace("=", "").str.replace("'", "").str.replace('"', ''))
    elif "xls" in filename:
        if filename.find("考拉") > 0:
            df = pd.read_excel(filename, sheet_name="销售明细", keep_default_na=False, dtype=str)
            df = df.apply(lambda x: x.astype(str).str.replace("=", "").str.replace("'", "").str.replace('"', ''))
        elif ((filename.find("淘宝") > 0) or (filename.find("天猫") > 0)):
            try:
                df = pd.read_excel(filename, keep_default_na=False, dtype=str)
                df = df.apply(lambda x: x.astype(str).str.replace("=", "").str.replace("'", "").str.replace('"', ''))
            except Exception as e:
                dict = {"订单编号": "", "退款金额": "", "退款金额（外币）": "", "是否回款": "", "结算方式": "", "回款金额": "", "回款日期": "",
                        "回款差额": "", "回款名称": "", "回款币种": "", "回款汇率": "", "回款金额（外币）": "", "税金": "", "财务费用": ""}
                df = pd.DataFrame(dict, index=[0])
                print("Excel文件读取出错:", e)
                return df
        else:
            df = pd.read_excel(filename, keep_default_na=False, dtype=str)
            df = df.apply(lambda x: x.astype(str).str.replace("=", "").str.replace("'", "").str.replace('"', ''))
    else:
        if (filename.find("淘宝") > 0) or (filename.find("天猫") > 0):
            df = pd.read_csv(filename, skiprows=4, keep_default_na=False, dtype=str, encoding="gb18030")
            df = df.apply(lambda x: x.astype(str).str.replace("=", "").str.replace("'", "").str.replace('"', ''))
        elif filename.find("拼多多") > 0:
            df = pd.read_csv(filename, skiprows=4, keep_default_na=False, dtype=str, encoding="gb18030")
            df = df.apply(lambda x: x.astype(str).str.replace("=", "").str.replace("'", "").str.replace('"', ''))
        elif filename.find("考拉") > 0:
            df = pd.read_csv(filename, sheet_name="销售明细", keep_default_na=False, dtype=str, encoding="gb18030")
            df = df.apply(lambda x: x.astype(str).str.replace("=", "").str.replace("'", "").str.replace('"', ''))
        else:
            df = pd.read_csv(filename, keep_default_na=False, dtype=str, encoding="gb18030")
            df = df.apply(lambda x: x.astype(str).str.replace("=", "").str.replace("'", "").str.replace('"', ''))
    print(df.head(1).to_markdown())
    # plat = os.path.basename(filename)
    if (filename.find("淘宝") > 0) and (filename.count("淘宝") < filename.count("天猫")):
        plat = "淘宝"
        if (filename.find("账单") > 0) or (filename.find("支付宝") > 0):
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
                    df = df[
                        ["订单编号", "退款金额", "退款金额（外币）", "是否回款", "结算方式", "回款金额", "回款日期", "回款差额", "回款名称", "回款币种", "回款汇率",
                         "回款金额（外币）", "税金", "财务费用"]]
                    df = df[df["退款金额（外币）"] != df["回款金额（外币）"]]
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
                    if (df["商户订单号"].shape[0] > 0) and (df[df["商户订单号"].str.contains("T200P")].shape[0] > 0):
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
                    df = df[
                        ["订单编号", "退款金额", "退款金额（外币）", "是否回款", "结算方式", "回款金额", "回款日期", "回款差额", "回款名称", "回款币种", "回款汇率",
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
            dict = {"订单编号": "", "退款金额": "", "退款金额（外币）": "", "是否回款": "", "结算方式": "", "回款金额": "", "回款日期": "",
                    "回款差额": "", "回款名称": "", "回款币种": "", "回款汇率": "", "回款金额（外币）": "", "税金": "", "财务费用": ""}
            df = pd.DataFrame(dict, index=[0])

    elif (filename.find("天猫") > 0) and (filename.count("淘宝") < filename.count("天猫")):
        plat = "天猫"
        if (filename.find("账单") > 0) or (filename.find("支付宝") > 0):
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
                    df = df[
                        ["订单编号", "退款金额", "退款金额（外币）", "是否回款", "结算方式", "回款金额", "回款日期", "回款差额", "回款名称", "回款币种", "回款汇率",
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
                    if (df["商户订单号"].shape[0] > 0) and (df[df["商户订单号"].str.contains("T200P")].shape[0] > 0):
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
                    df = df[
                        ["订单编号", "退款金额", "退款金额（外币）", "是否回款", "结算方式", "回款金额", "回款日期", "回款差额", "回款名称", "回款币种", "回款汇率",
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
            dict = {"订单编号": "", "退款金额": "", "退款金额（外币）": "", "是否回款": "", "结算方式": "", "回款金额": "", "回款日期": "",
                    "回款差额": "", "回款名称": "", "回款币种": "", "回款汇率": "", "回款金额（外币）": "", "税金": "", "财务费用": ""}
            df = pd.DataFrame(dict, index=[0])

    elif filename.find("京东") > 0:
        plat = "京东"
        if filename.find("账单") > 0:
            df["平台"] = plat
            if filename.find("海外") > 0:
                return_file = filename[:-12] + "退货结算数据.csv"  # 对文件名进行修改
                df1 = pd.read_csv(return_file, skiprows=1, keep_default_na=False, dtype=str, encoding="gb18030")
                df = df[:-1]
                df1 = df1[:-1]
                df = pd.merge(df, df1[["订单编号", "退款金额"]], how="outer", on="订单编号")
                # print(f"df:\n{df.head(5).to_markdown()}")
                # print(f"df1:\n{df1.head(5).to_markdown()}")
                if (df.shape[0] > 0 or df1.shape[0] > 0):
                    df["订单编号"] = df["订单编号"]
                    df["退款金额"] = 0
                    df["退款金额（外币）"] = df["退款金额"]
                    df["是否回款"] = "是"
                    df["结算方式"] = ""
                    df["回款金额"] = 0
                    df["回款日期"] = ""
                    df["回款差额"] = 0
                    df["回款名称"] = "".join(filename.split("/")[-1:])
                    df["回款币种"] = "USD"
                    df["回款汇率"] = df["汇率(美元/人民币)"]
                    df["回款金额（外币）"] = df["结算金额"]
                    df["税金"] = 0.019
                    df["财务费用"] = 2.2
                    df = df[
                        ["订单编号", "退款金额", "退款金额（外币）", "是否回款", "结算方式", "回款金额", "回款日期", "回款差额", "回款名称", "回款币种", "回款汇率",
                         "回款金额（外币）", "税金", "财务费用"]]
                    # df = df[~df["退款金额（外币）"].str.contains(0) & ~df["回款金额（外币）"].str.contains(0)]
                    print("京东-海外-有回款")
                    print(df.head(5).to_markdown())
                else:
                    dict = {"订单编号": "", "退款金额": "", "退款金额（外币）": "", "是否回款": "", "结算方式": "", "回款金额": "", "回款日期": "",
                            "回款差额": "", "回款名称": "", "回款币种": "", "回款汇率": "", "回款金额（外币）": "", "税金": "", "财务费用": ""}
                    df = pd.DataFrame(dict, index=[0])
                    # df = df[["订单编号","退款金额","退款金额（外币）","是否回款","结算方式","回款金额","回款日期","回款差额","回款名称","回款币种","回款汇率","回款金额（外币）","税金","财务费用"]]
                    print("京东-海外-无回款")
                    print(df.head(5).to_markdown())
            else:
                if df.shape[0] > 0:
                    df["订单编号"] = df["订单编号"]
                    df["退款金额"] = 0
                    df["退款金额（外币）"] = 0
                    df["是否回款"] = "是"
                    df["结算方式"] = ""
                    df["回款金额"] = df["金额"]
                    df["回款日期"] = ""
                    df["回款差额"] = 0
                    df["回款名称"] = "".join(filename.split("/")[-1:])
                    df["回款币种"] = ""
                    df["回款汇率"] = 0
                    df["回款金额（外币）"] = 0
                    df["税金"] = 0
                    df["财务费用"] = 0
                    df["回款金额"] = df["回款金额"].astype(float)
                    df2 = df.loc[df["钱包结算备注"].str.contains("退货金额")]
                    df2["退款金额"] = df2["金额"]
                    df2["回款金额"] = 0
                    df1 = df.loc[((df["费用项"].str.contains("货款")) & (df["收支方向"].str.contains("收入")) & (df["回款金额"] > 0))]

                    # del df1["退款金额"]
                    print(df1.head(5).to_markdown())
                    print(df2.head(5).to_markdown())
                    dfs = [df1, df2]
                    # df = pd.merge(df1,df2[["订单编号","退款金额"]],how="outer",on="订单编号")
                    df = pd.concat(dfs)
                    df["退款金额"] = df["退款金额"].astype(float)
                    df = df[
                        ["订单编号", "退款金额", "退款金额（外币）", "是否回款", "结算方式", "回款金额", "回款日期", "回款差额", "回款名称", "回款币种", "回款汇率",
                         "回款金额（外币）", "税金", "财务费用"]]
                    df = df.sort_values(by=["订单编号"])
                    # df = df[~df["退款金额（外币）"].str.contains(0) & ~df["回款金额（外币）"].str.contains(0)]
                    print("京东-国内-有回款")
                    print(df.head(5).to_markdown())
                else:
                    dict = {"订单编号": "", "退款金额": "", "退款金额（外币）": "", "是否回款": "", "结算方式": "", "回款金额": "", "回款日期": "",
                            "回款差额": "", "回款名称": "", "回款币种": "", "回款汇率": "", "回款金额（外币）": "", "税金": "", "财务费用": ""}
                    df = pd.DataFrame(dict, index=[0])
                    print("京东-国内-无回款")
                    print(df.head(5).to_markdown())
        else:
            print("没有账单文件")
            dict = {"订单编号": "", "退款金额": "", "退款金额（外币）": "", "是否回款": "", "结算方式": "", "回款金额": "", "回款日期": "",
                    "回款差额": "", "回款名称": "", "回款币种": "", "回款汇率": "", "回款金额（外币）": "", "税金": "", "财务费用": ""}
            df = pd.DataFrame(dict, index=[0])

    elif filename.find("拼多多") > 0:
        plat = "拼多多"
        if filename.find("账单") > 0:
            df["平台"] = plat
            df = df[(~df["商户订单号"].str.contains("#|结算汇总|提现汇总"))]
            df.dropna(subset=["商户订单号"], inplace=True)
            print(df.tail(5).to_markdown())
            if filename.find("海外") > 0:
                if df.shape[0] > 0:
                    df["订单编号"] = df["商户订单号"]
                    df["退款金额"] = 0
                    if "账务类型" in df.columns:
                        print("有账务类型的列")
                        df.loc[df["账务类型"].str.contains("退款"), "退款金额"] = df["支出金额（-元）"]
                    else:
                        print("没有账务类型的列")
                        df["退款金额"] = 0
                    df["退款金额（外币）"] = 0
                    if "账务类型" in df.columns:
                        df.loc[df["账务类型"].str.contains("退款"), "退款金额（外币）"] = df["支出金额（-$）"]
                    else:
                        df["退款金额（外币）"] = 0
                    df["是否回款"] = "是"
                    df["结算方式"] = ""
                    df["回款金额"] = 0
                    if "账务类型" in df.columns:
                        df.loc[df["账务类型"].str.contains("交易收入|优惠券结算"), "回款金额"] = df["收入金额（+元）"]
                        df.loc[~df["账务类型"].str.contains("交易收入|优惠券结算"), "回款金额"] = 0
                    else:
                        df["回款金额"] = 0
                    df["回款日期"] = df["发生时间"]
                    df["回款差额"] = 0
                    df["回款名称"] = "".join(filename.split("/")[-1:])
                    if "汇率(美元/人民币)" in df.columns:
                        df["回款币种"] = "USD"
                    elif "汇率（美元兑人民币）" in df.columns:
                        df["回款币种"] = "USD"
                    else:
                        df["回款币种"] = "RMB"
                    if "汇率(美元/人民币)" in df.columns:
                        df["回款汇率"] = df["汇率(美元/人民币)"]
                    elif "汇率（美元兑人民币）" in df.columns:
                        df["回款汇率"] = df["汇率（美元兑人民币）"]
                    else:
                        df["回款汇率"] = 0
                    df["回款金额（外币）"] = 0
                    if "账务类型" in df.columns:
                        df.loc[df["账务类型"].str.contains("交易收入|优惠券结算"), "回款金额（外币）"] = df["收入金额（+$）"]
                        df.loc[~df["账务类型"].str.contains("交易收入|优惠券结算"), "回款金额（外币）"] = 0
                    else:
                        df["回款金额（外币）"] = 0
                    df["税金"] = 0.019
                    df["财务费用"] = 2.2
                    df = df[
                        ["订单编号", "退款金额", "退款金额（外币）", "是否回款", "结算方式", "回款金额", "回款日期", "回款差额", "回款名称", "回款币种", "回款汇率",
                         "回款金额（外币）", "税金", "财务费用"]]
                    df = df.loc[(df["退款金额（外币）"] != df["回款金额（外币）"])]
                    # df = df[~df["退款金额（外币）"].str.contains(0) & ~df["回款金额（外币）"].str.contains(0)]
                    print("拼多多-海外-有回款")
                    print(df.head(5).to_markdown())
                else:
                    dict = {"订单编号": "", "退款金额": "", "退款金额（外币）": "", "是否回款": "", "结算方式": "", "回款金额": "", "回款日期": "",
                            "回款差额": "", "回款名称": "", "回款币种": "", "回款汇率": "", "回款金额（外币）": "", "税金": "", "财务费用": ""}
                    df = pd.DataFrame(dict, index=[0])
                    # df = df[["订单编号","退款金额","退款金额（外币）","是否回款","结算方式","回款金额","回款日期","回款差额","回款名称","回款币种","回款汇率","回款金额（外币）","税金","财务费用"]]
                    print("拼多多-海外-无回款")
                    print(df.head(5).to_markdown())
            else:
                if df.shape[0] > 0:
                    df["订单编号"] = df["商户订单号"]
                    df["退款金额"] = 0
                    if "账务类型" in df.columns:
                        print("有账务类型的列")
                        df.loc[df["账务类型"].str.contains("退款"), "退款金额"] = df["支出金额（-元）"]
                    else:
                        print("没有账务类型的列")
                        df["退款金额"] = 0
                    df["退款金额（外币）"] = 0
                    df["是否回款"] = "是"
                    df["结算方式"] = ""
                    df["回款金额"] = 0
                    if "账务类型" in df.columns:
                        df.loc[df["账务类型"].str.contains("交易收入|优惠券结算"), "回款金额"] = df["收入金额（+元）"]
                        df.loc[~df["账务类型"].str.contains("交易收入|优惠券结算"), "回款金额"] = 0
                    else:
                        df["回款金额"] = 0
                    df["回款日期"] = df["发生时间"]
                    df["回款差额"] = 0
                    df["回款名称"] = "".join(filename.split("/")[-1:])
                    df["回款币种"] = "RMB"
                    df["回款汇率"] = 0
                    df["回款金额（外币）"] = 0
                    df["税金"] = 0
                    df["财务费用"] = 0
                    df = df[
                        ["订单编号", "退款金额", "退款金额（外币）", "是否回款", "结算方式", "回款金额", "回款日期", "回款差额", "回款名称", "回款币种", "回款汇率",
                         "回款金额（外币）", "税金", "财务费用"]]
                    df = df.loc[(df["退款金额"] != df["回款金额"])]
                    # df = df[~df["退款金额（外币）"].str.contains(0) & ~df["回款金额（外币）"].str.contains(0)]
                    print("拼多多-国内-有回款")
                    print(df.head(5).to_markdown())
                else:
                    dict = {"订单编号": "", "退款金额": "", "退款金额（外币）": "", "是否回款": "", "结算方式": "", "回款金额": "", "回款日期": "",
                            "回款差额": "", "回款名称": "", "回款币种": "", "回款汇率": "", "回款金额（外币）": "", "税金": "", "财务费用": ""}
                    df = pd.DataFrame(dict, index=[0])
                    print("拼多多-国内-无回款")
                    print(df.head(5).to_markdown())
        else:
            print("没有账单文件")

    elif filename.find("考拉") > 0:
        plat = "网易考拉"
        if filename.find("账单") > 0:
            if "xls" in filename:
                df1 = pd.read_excel(filename, sheet_name="退款明细", keep_default_na=False, dtype=str)
            else:
                df1 = pd.read_csv(filename, sheet_name="退款明细", keep_default_na=False, dtype=str, encoding="gb18030")
            if df1.shape[0] > 0:
                df = pd.merge(df, df1[["销售订单号", "商品实退金额（含税含发货运费）"]], how="outer", on="销售订单号")
            else:
                pass
            df["平台"] = plat
            df.dropna(subset=["销售订单号"], inplace=True)
            print(df.tail(5).to_markdown())
            if df.shape[0] > 0:
                df["订单编号"] = df["销售订单号"]
                df["退款金额"] = 0
                if "商品实退金额（含税含发货运费）" in df.columns:
                    print("有退款数据")
                    df["退款金额"] = df["商品实退金额（含税含发货运费）"]
                    df["退款金额"] = df["退款金额"].fillna(0)
                else:
                    print("无退款数据")
                    df["退款金额"] = 0
                df["退款金额（外币）"] = 0
                df["是否回款"] = "是"
                df["结算方式"] = ""
                df["回款金额"] = df["商品实付"]
                df["回款日期"] = df["结算日期"]
                df["回款差额"] = 0
                df["回款名称"] = "".join(filename.split("/")[-1:])
                df["回款币种"] = ""
                df["回款汇率"] = 0
                df["回款金额（外币）"] = 0
                df["税金"] = 0
                df["财务费用"] = 0
                df = df[
                    ["订单编号", "退款金额", "退款金额（外币）", "是否回款", "结算方式", "回款金额", "回款日期", "回款差额", "回款名称", "回款币种", "回款汇率",
                     "回款金额（外币）", "税金", "财务费用"]]
                df = df.loc[(df["退款金额"] != df["回款金额"])]
                # df = df[~df["退款金额（外币）"].str.contains(0) & ~df["回款金额（外币）"].str.contains(0)]
                print("网易考拉-有回款")
                print(df.head(5).to_markdown())
            else:
                dict = {"订单编号": "", "退款金额": "", "退款金额（外币）": "", "是否回款": "", "结算方式": "", "回款金额": "", "回款日期": "",
                        "回款差额": "", "回款名称": "", "回款币种": "", "回款汇率": "", "回款金额（外币）": "", "税金": "", "财务费用": ""}
                df = pd.DataFrame(dict, index=[0])
                print("网易考拉-无回款")
                print(df.head(5).to_markdown())
        else:
            print("没有账单文件")

    elif filename.find("有赞") > 0:
        plat = "有赞"
        if filename.find("账单") > 0:
            df["平台"] = plat
            df.dropna(subset=["业务单号"], inplace=True)
            print(df.tail(5).to_markdown())
            if df.shape[0] > 0:
                df["订单编号"] = df["业务单号"]
                df["退款金额"] = df["支出(元)"]
                df["退款金额（外币）"] = 0
                df["是否回款"] = "是"
                df["结算方式"] = ""
                df["回款金额"] = df["收入(元)"]
                df["回款日期"] = df["入账时间"]
                df["回款差额"] = 0
                df["回款名称"] = "".join(filename.split("/")[-1:])
                df["回款币种"] = ""
                df["回款汇率"] = 0
                df["回款金额（外币）"] = 0
                df["税金"] = 0
                df["财务费用"] = 0
                # df[["退款金额", "回款金额"]] = df[["退款金额", "回款金额"]].astype(float)
                # df1 = pd.DataFrame(df.groupby("订单编号")[["退款金额", "回款金额"]].sum()).reset_index()
                # df1[["退款金额", "回款金额"]] = df1[["退款金额", "回款金额"]].astype(float)
                # df = pd.merge(df1,df[
                #     ["订单编号", "退款金额（外币）", "是否回款", "结算方式", "回款日期", "回款差额", "回款名称", "回款币种", "回款汇率",
                #      "回款金额（外币）", "税金", "财务费用"]],how="left",on="订单编号")
                df = df[
                    ["订单编号", "退款金额", "退款金额（外币）", "是否回款", "结算方式", "回款金额", "回款日期", "回款差额", "回款名称", "回款币种", "回款汇率",
                     "回款金额（外币）", "税金", "财务费用"]]
                # df.drop_duplicates(subset=["订单编号"],keep="last",inplace=True)
                # df_grouped = pd.DataFrame(df.groupby("订单编号")[["退款金额","回款金额"]].sum()).reset_index()
                # df = df.loc[(df["退款金额"] != df["回款金额"])]
                # df = df[~df["退款金额（外币）"].str.contains(0) & ~df["回款金额（外币）"].str.contains(0)]
                print("有赞-有回款")
                print(df.to_markdown())
            else:
                dict = {"订单编号": "", "退款金额": "", "退款金额（外币）": "", "是否回款": "", "结算方式": "", "回款金额": "", "回款日期": "",
                        "回款差额": "", "回款名称": "", "回款币种": "", "回款汇率": "", "回款金额（外币）": "", "税金": "", "财务费用": ""}
                df = pd.DataFrame(dict, index=[0])
                print("有赞-无回款")
                print(df.head(5).to_markdown())
        else:
            print("没有账单文件")

    return df


def get_order_files(rootdir, filekey):
    filelist = list_all_files(rootdir, filekey)
    # print(filelist)
    mySeries = pd.Series(filelist)
    df = pd.DataFrame(mySeries)
    df.columns = ["filename"]
    # df["filename"] = df["filename"].apply(lambda x: x.strip().replace(" ", ""))

    df = df[~df["filename"].str.contains("快递")]
    df = df[~df["filename"].str.contains("账单")]
    df = df[~df["filename"].str.contains("无")]
    df = df[~df["filename"].str.contains("推广费")]
    df = df[~df["filename"].str.contains("商品")]
    df = df[~df["filename"].str.contains("数据转换")]
    df = df[~df["filename"].str.contains("支付宝")]

    # print(df.to_markdown())
    # print("抽查是否还有快递！")
    # print(df[df.filename.str.contains("快递")].to_markdown())
    return df


def get_bill_files(rootdir, filekey):
    filelist = list_all_files(rootdir, filekey)
    # print(filelist)
    mySeries = pd.Series(filelist)
    df = pd.DataFrame(mySeries)
    df.columns = ["filename"]
    # df["filename"] = df["filename"].apply(lambda x: x.strip().replace(" ", ""))

    df = df[~df["filename"].str.contains("快递")]
    df = df[~df["filename"].str.contains("订单")]
    df = df[~df["filename"].str.contains("无")]
    df = df[~df["filename"].str.contains("推广费")]
    df = df[~df["filename"].str.contains("商品")]
    df = df[~df["filename"].str.contains("数据转换")]
    df = df[~df["filename"].str.contains("汇总")]
    df = df[~df["filename"].str.contains("业务")]
    df = df[~df["filename"].str.contains(".zip")]
    df = df[~df["filename"].str.contains("结算数据")]
    df = df[~df["filename"].str.contains("其他费用")]

    # print(df.to_markdown())
    # print("抽查是否还有快递！")
    # print(df[df.filename.str.contains("快递")].to_markdown())
    return df


def read_order_excel(rootdir, filekey):
    df_files = get_order_files(rootdir, filekey)
    for index, file in df_files.iterrows():
        if 'df' in locals().keys():  # 如果变量已经存在
            dd = read_order(file["filename"])
            dd["filename"] = file["filename"]
            print("进度表：{}/{}  文件{}，行数{}".format(index + 1, df_files.shape[0], file["filename"], dd.shape[0]))
            # print(dd.head(1).to_markdown())
            df = df.append(dd)

        else:
            df = read_order(file["filename"])
            df["filename"] = file["filename"]
            print("进度表：{}/{}  文件{}，行数{}".format(index + 1, df_files.shape[0], file["filename"], df.shape[0]))

    return df


def read_bill_excel(rootdir, filekey):
    df_files = get_bill_files(rootdir, filekey)
    for index, file in df_files.iterrows():
        if 'df' in locals().keys():  # 如果变量已经存在
            dd = read_bill(file["filename"])
            dd["filename"] = file["filename"]
            print("进度表：{}/{}  文件{}，行数{}".format(index + 1, df_files.shape[0], file["filename"], dd.shape[0]))
            # print(dd.head(1).to_markdown())
            df = df.append(dd)
            print("append数据")
            # print(df.head(1).to_markdown())
        else:
            df = read_bill(file["filename"])
            df["filename"] = file["filename"]
            print("进度表：{}/{}  文件{}，行数{}".format(index + 1, df_files.shape[0], file["filename"], df.shape[0]))

    return df


def all_order():
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

    mkdir(filedir)

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

    table = read_order_excel(filedir, filekey)
    del table["filename"]
    table["订单编号"].replace("", np.nan, inplace=True)
    table.dropna(axis=0, subset=["订单编号"], inplace=True)
    table = table.drop_duplicates()

    # table.to_excel(default_dir + "/数据转换/合并表格.xlsx",index=False)
    # table.to_excel("data/合并表格.xlsx",index=False)
    # 转换后的文件合并导出
    print(len(table))
    if len(table) > 500000:
        table.to_csv(default_dir + "/数据转换/全平台_数据转换结果.csv", index=False)
    else:
        table.to_excel(default_dir + "/数据转换/全平台_数据转换后订单.xlsx", index=False)

    # 转换后的文件按平台拆分导出
    plat_list = ["淘宝", "天猫", "京东", "拼多多", "抖音", "快手", "小红书", "网易考拉", "有赞", "阿里巴巴"]
    for plat in plat_list:
        df_plat = table[table["平台"].str.contains(plat)]
        if df_plat.shape[0] > 0:
            pagecount = math.ceil(df_plat.shape[0] / 300000.00)
            pagecount = "{:d}".format(pagecount)
            print("总共需要拆分{}个文件".format(pagecount))
            writer = pd.ExcelWriter(default_dir + "/数据转换/{}_数据转换结果.xlsx".format(plat))
            # writer = pd.ExcelWriter("data/{}_数据转换结果.xlsx".format(plat))
            for x in range(int(pagecount)):
                from_line = x * 300000
                to_line = (x + 1) * 300000
                plat_table = df_plat[from_line:to_line]
                print(plat_table.head(5).to_markdown())
                print("输出文件总行数：{}".format(plat_table.shape[0]))
                sheetname = "Sheet{}".format(x + 1)
                plat_table.to_excel(writer, sheetname, engine='xlsxwriter', index=False)
                format1 = writer.book.add_format({'num_format': '0.00'})
                writer.book.sheetnames[sheetname].set_column('O:O', cell_format=format1)
            writer.save()
        else:
            pass

    return table


def all_bill():
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

    mkdir(filedir)

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

    table = read_bill_excel(filedir, filekey)
    del table["filename"]
    # table["订单编号"].replace("", np.nan, inplace=True)
    print("删除空白订单编号前行数：")
    print(table.shape[0])
    table.dropna(axis=0, subset=["订单编号"], inplace=True)
    print("删除重复数据前行数：")
    print(table.shape[0])
    table.drop_duplicates(inplace=True)
    print("最终行数：")
    print(table.shape[0])
    # table["退款金额"] = table["退款金额"].astype(float)
    # table["退款金额（外币）"] = table["退款金额（外币）"].astype(float)
    # table["回款金额"] = table["回款金额"].astype(float)
    # table["回款金额（外币）"] = table["回款金额（外币）"].astype(float)
    # table.to_excel(default_dir + "/数据转换/合并表格.xlsx",index=False)
    # table.to_excel("data/合并表格.xlsx",index=False)
    # 转换后的文件合并导出
    # print(len(table))
    if len(table) > 500000:
        table.to_csv("data/全平台_账单转换结果.csv", index=False)
    else:
        table.to_excel("/data/全平台_账单转换结果.xlsx", index=False)

    # 转换后的文件按平台拆分导出
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


def combine_excel():
    all_order()
    all_bill()


def test():
    df = pd.read_csv(r"/Users/maclove/Downloads/2019各平台数据原表/全平台_数据转换结果.csv", dtype=str, keep_default_na=False)
    table = pd.DataFrame(df)
    table.订单编号.fillna("nan")
    # print(table[table["订单编号"].str.contains("")])
    # print(table[table["订单编号"].str.contains("nan")].to_markdown())
    print(len(table))
    table["订单编号"].replace("", np.nan, inplace=True)
    table.dropna(axis=0, subset=["订单编号"], inplace=True)
    print(len(table))
    # print(table[table["订单编号"].str.contains("nan")])
    table["出现序号"] = table["出现序号"].astype(float)
    table["总序号"] = table["总序号"].astype(float)
    table["购买数量"] = table["购买数量"].astype(float)
    table["销售金额"] = table["销售金额"].astype(float)
    plat_list = ["淘宝", "天猫", "京东", "拼多多", "抖音", "快手", "小红书", "网易考拉", "有赞", "阿里巴巴"]
    for plat in plat_list:
        df_plat = table[table["平台"].str.contains(plat)]
        if df_plat.shape[0] > 0:
            pagecount = math.ceil(df_plat.shape[0] / 300000.00)
            pagecount = "{:d}".format(pagecount)
            print("总共需要拆分{}个文件".format(pagecount))
            # writer = pd.ExcelWriter(default_dir + "/数据转换/{}_数据转换结果.xlsx".format(plat))
            writer = pd.ExcelWriter("data/{}_数据转换结果.xlsx".format(plat))
            for x in range(int(pagecount)):
                from_line = x * 300000
                to_line = (x + 1) * 300000
                plat_table = df_plat[from_line:to_line]
                print(plat_table.head(5).to_markdown())
                print("输出文件总行数：{}".format(plat_table.shape[0]))
                sheetname = "Sheet{}".format(x + 1)
                plat_table.to_excel(writer, sheetname, engine='xlsxwriter', index=False)
                format1 = writer.book.add_format({'num_format': '0.00'})
                writer.book.sheetnames[sheetname].set_column('O:O', cell_format=format1)
            writer.save()
        else:
            pass


if __name__ == "__main__":
    print("开始:", time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()))
    # filename = r"/Users/maclove/Downloads/2019各平台数据原表/京东/订单/Dicora UrbanFit海外旗舰店/Dicora UrbanFit海外旗舰店201904-订单.xlsx"

    # all_order()
    all_bill()
    # test()

    print("ok")
