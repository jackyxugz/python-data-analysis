#!/usr/bin/env python
# coding: utf-8
import sys
import numpy as np
import os
import codecs
sys.stdout = codecs.getwriter("utf-8")(sys.stdout.detach())
import pandas as pd
import datetime
# from IPython.display import display, HTML
import os.path
import time
import warnings
import tabulate
import openpyxl
import math
import difflib
import re


# 读取商品销售sheet
def read_settle(filename):
    data_xls = pd.read_excel(filename, sheet_name=None, dtype=str)
    for key in data_xls:
        if key == "商品销售":
            df=pd.read_excel(filename,sheet_name="商品销售")
            df["id"]=df.index
            left_df=df[["id","订单号","商品名称","商品id","发货时间","卖家名称"]]

            for col in df.columns:
                if col in ["收入总额"]:
                    t_df=df[["id","订单号","收入总额"]]
                    t_df["类型"]="收入"
                    t_df.columns=["id","订单号","收入金额","类型"]
                    t_df["支出金额"]="0"
                    t_df["收支项目"]="商品销售"

                    if ~("df_value" in vars()):
                        df_value=t_df[["id","订单号","类型","收入金额","支出金额","收支项目"]].copy()
                    else:
                        df_value=df_value.append(t_df[["id","订单号","类型","收入金额","支出金额","收支项目"]])
                elif col in ["商品退货","人工退款","包材费"]:
                    t_df=df[["id",col]]
                    t_df["类型"]="支出"
                    t_df.columns=["id","支出金额","类型"]
                    t_df["收入金额"]="0"
                    t_df["收支项目"]=col

                    if ~("df_value" in vars()):
                        df_value=t_df[["id","订单号","类型","收入金额","支出金额","收支项目"]].copy()
                    else:
                        df_value=df_value.append(t_df[["id","订单号","类型","收入金额","支出金额","收支项目"]])

            # df_h= left_df.merge(df_Value,how="left",on="id")
            return left_df,df_value

    print("{} 没有发现 商品销售 sheet".format(filename))
    return "",""
    # return   pd.DataFrame(columns=["id","订单号","商品名称","商品id","发货时间","卖家名称"],index=[0]), pd.DataFrame(columns=["id","订单号","类型","收入金额","支出金额","收支项目"],index=[0])

# 获取商品销售sheet表——佣金总额
def read_commission(filename):
    df = pd.read_excel(filename, sheet_name="商品销售")
    df["id"] = df.index

    for col in df.columns:
        if col in ["佣金总额"]:
            t_df = df[["id", "订单号", "佣金总额"]]
            t_df["类型"] = "支出"
            t_df.columns = ["id", "订单号", "佣金总额", "类型"]
            t_df["收入金额"] = df["佣金总额"]
            t_df["收支项目"] = "销售佣金"

            # t_df = df[["id", "订单号", "支出总额", "类型"]]

            # 如果是支出
            t_df.loc[t_df.类型.str.contains("支出"), "支出金额"] = t_df["佣金总额"]
            # 如果是收入
            t_df.loc[t_df.类型.str.contains("收入"), "收入金额"] = t_df["佣金总额"]

            t_df["收支项目"] = "销售佣金"

            if ~("df_value" in vars()):
                df_value = t_df[["id", "订单号", "类型", "收入金额", "支出金额", "收支项目"]].copy()
            else:
                df_value = df_value.append(t_df[["id", "订单号", "类型", "收入金额", "支出金额", "收支项目"]])

    return df_value

# 获取商品销售sheet表——支付渠道费
def read_Payment_channel_fee(filename):
    data_xls = pd.read_excel(filename, sheet_name=None, dtype=str)
    for key in data_xls:
        if key == "商品销售":
            df = pd.read_excel(filename, sheet_name="商品销售")
            df["id"] = df.index
            # left_df = df[["id", "订单号", "支付渠道费(商品)"]]
            # df = pd.read_excel(filename, sheet_name="人工退款")
            # df["id"] = df.index

            for col in df.columns:
                if col in ["支付渠道费(商品)"]:
                    t_df = df[["id", "订单号", "支付渠道费(商品)"]]
                    t_df["类型"] = "支出"
                    t_df.columns = ["id", "订单号", "支付渠道费(商品)", "类型"]
                    t_df["收入金额"] = "0"
                    t_df["收支项目"] = "支付渠道费（商品销售）"

                    # t_df = df[["id", "订单号", "支出金额", "类型"]]

                    # 如果是支出
                    t_df.loc[t_df.类型.str.contains("支出"), "支出金额"] = t_df["支付渠道费(商品)"]
                    # 如果是收入
                    t_df.loc[t_df.类型.str.contains("收入"), "收入金额"] = t_df["支付渠道费(商品)"]

                    # t_df["收支项目"] = "支付渠道费"

                    if ~("df_value" in vars()):
                        df_value = t_df[["id", "订单号", "类型", "收入金额", "支出金额", "收支项目"]].copy()
                    else:
                        df_value = df_value.append(t_df[["id", "订单号", "类型", "收入金额", "支出金额", "收支项目"]])

                    return  df_value
        else:
            pass


# 读取 订单运费
def read_express(filename):
    df=pd.read_excel(filename,sheet_name="订单运费")
    df["id"]=df.index
    
    for col in df.columns:
        if col in ["运费","支付渠道费"]:
            t_df=df[["订单号",col]]
            t_df["类型"]="支出"
            t_df.columns=["订单号","支出金额","类型"]
            t_df["收入金额"]="0"
            t_df["收支项目"]="订单运费"
            
            if ~("df_value" in vars()):
                df_value=t_df[["订单号","类型","收入金额","支出金额","收支项目"]].copy()
            else:
                df_value=df_value.append(t_df[["订单号","类型","收入金额","支出金额","收支项目"]])
      

    
    return df_value

# 获取订单运费sheet表——支付渠道费
def read_express_Payment_channel_fee(filename):
    data_xls = pd.read_excel(filename, sheet_name=None, dtype=str)
    for key in data_xls:
        if key == "订单运费":
            df = pd.read_excel(filename, sheet_name="订单运费")
            df["id"] = df.index
            # left_df = df[["id", "订单号", "支付渠道费(商品)"]]
            # df = pd.read_excel(filename, sheet_name="人工退款")
            # df["id"] = df.index

            for col in df.columns:
                if col in ["支付渠道费"]:
                    t_df = df[["id", "订单号", "支付渠道费"]]
                    t_df["类型"] = "支出"
                    t_df.columns = ["id", "订单号", "支付渠道费", "类型"]
                    t_df["收入金额"] = "0"
                    t_df["收支项目"] = "支付渠道费（订单运费）"

                    # t_df = df[["id", "订单号", "支出金额", "类型"]]

                    # 如果是支出
                    t_df.loc[t_df.类型.str.contains("支出"), "支出金额"] = t_df["支付渠道费"]
                    # 如果是收入
                    t_df.loc[t_df.类型.str.contains("收入"), "收入金额"] = t_df["支付渠道费"]

                    # t_df["收支项目"] = "支付渠道费"

                    if ~("df_value" in vars()):
                        df_value = t_df[["id", "订单号", "类型", "收入金额", "支出金额", "收支项目"]].copy()
                    else:
                        df_value = df_value.append(t_df[["id", "订单号", "类型", "收入金额", "支出金额", "收支项目"]])

                    return  df_value
        else:
            pass
          
        
# 读取 赔偿用户薯券
def read_compensate(filename):
    df=pd.read_excel(filename,sheet_name="赔偿用户薯券")
    df["id"]=df.index
    
    for col in df.columns:
        if col in ["支出金额"]:
            t_df=df[["id","订单号","支出金额"]]
            t_df["类型"]="支出"
            t_df.columns=["id","订单号","支出金额","类型"]
            t_df["收入金额"]="0"
            t_df["收支项目"]="赔偿用户薯券"
            
            if ~("df_value" in vars()):
                df_value=t_df[["id","订单号","类型","收入金额","支出金额","收支项目"]].copy()
            else:
                df_value=df_value.append(t_df[["id","订单号","类型","收入金额","支出金额","收支项目"]])

    return df_value

# 读取 人工调账
def read_adjust(filename):
    df=pd.read_excel(filename,sheet_name="人工调账")
    df["id"]=df.index
    
    for col in df.columns:
        if col in ["结算总额"]:
            t_df=df[["id","订单号","结算总额","类型"]]
            
            # 如果是支出
            t_df.loc[t_df.类型.str.contains("支出"),"支出金额"]=t_df["结算总额"]
            #如果是收入
            t_df.loc[t_df.类型.str.contains("收入"),"收入金额"]=t_df["结算总额"]
            
            t_df["收支项目"]="人工调账"
            
            if ~("df_value" in vars()):
                df_value=t_df[["id","订单号","类型","收入金额","支出金额","收支项目"]].copy()
            else:
                df_value=df_value.append(t_df[["id","订单号","类型","收入金额","支出金额","收支项目"]])

    return df_value


# 读取 退货
def read_ReturnGoods(filename):
    df = pd.read_excel(filename, sheet_name="退货")
    df["id"] = df.index

    for col in df.columns:
        if col in ["支出总额"]:
            t_df = df[["id", "订单号", "支出总额"]]
            t_df["类型"] = "支出"
            t_df.columns = ["id", "订单号", "支出总额", "类型"]
            t_df["收入金额"] = "0"
            t_df["收支项目"] = "退货"

            # t_df = df[["id", "订单号", "支出总额", "类型"]]

            # 如果是支出
            t_df.loc[t_df.类型.str.contains("支出"), "支出金额"] = t_df["支出总额"]
            # 如果是收入
            t_df.loc[t_df.类型.str.contains("收入"), "收入金额"] = t_df["支出总额"]

            t_df["收支项目"] = "退货"

            if ~("df_value" in vars()):
                df_value = t_df[["id", "订单号", "类型", "收入金额", "支出金额", "收支项目"]].copy()
            else:
                df_value = df_value.append(t_df[["id", "订单号", "类型", "收入金额", "支出金额", "收支项目"]])

    return df_value


# 获取退货sheet表——支付渠道费
def read_ReturnGoods_Payment_channel_fee(filename):
    data_xls = pd.read_excel(filename, sheet_name=None, dtype=str)
    for key in data_xls:
        if key == "退货":
            df = pd.read_excel(filename, sheet_name="退货")
            df["id"] = df.index
            # left_df = df[["id", "订单号", "支付渠道费(商品)"]]
            # df = pd.read_excel(filename, sheet_name="人工退款")
            # df["id"] = df.index

            for col in df.columns:
                if col in ["支付渠道费"]:
                    t_df = df[["id", "订单号", "支付渠道费"]]
                    t_df["类型"] = "收入"
                    t_df.columns = ["id", "订单号", "支付渠道费", "类型"]
                    t_df["支出金额"] = "0"
                    t_df["收支项目"] = "支付渠道费（退货）"

                    # t_df = df[["id", "订单号", "支出金额", "类型"]]

                    # 如果是支出
                    t_df.loc[t_df.类型.str.contains("支出"), "支出金额"] = t_df["支付渠道费"]
                    # 如果是收入
                    t_df.loc[t_df.类型.str.contains("收入"), "收入金额"] = t_df["支付渠道费"]

                    # t_df["收支项目"] = "支付渠道费"

                    if ~("df_value" in vars()):
                        df_value = t_df[["id", "订单号", "类型", "收入金额", "支出金额", "收支项目"]].copy()
                    else:
                        df_value = df_value.append(t_df[["id", "订单号", "类型", "收入金额", "支出金额", "收支项目"]])

                    return df_value
        else:
            pass


# # 读取 退货-佣金
def read_ReturnGoods_commission(filename):
    data_xls = pd.read_excel(filename, sheet_name=None, dtype=str)
    for key in data_xls:
        if key == "退货":
            df = pd.read_excel(filename, sheet_name="退货")
            df["id"] = df.index
            # left_df = df[["id", "订单号", "佣金总额"]]
    # df = pd.read_excel(filename, sheet_name="人工退款")
    # df["id"] = df.index

            for col in df.columns:
                if col in ["佣金总额"]:
                    t_df = df[["id", "订单号", "佣金总额"]]
                    t_df["类型"] = "支出"
                    t_df.columns = ["id", "订单号", "佣金总额", "类型"]
                    t_df["收入金额"] = df["佣金总额"]
                    t_df["收支项目"] = "退货佣金"

                    # t_df = df[["id", "订单号", "支出金额", "类型"]]

                    # 如果是支出
                    t_df.loc[t_df.类型.str.contains("支出"), "支出金额"] = t_df["佣金总额"]
                    # 如果是收入
                    t_df.loc[t_df.类型.str.contains("收入"), "收入金额"] = t_df["佣金总额"]

                    t_df["收支项目"] = "退货佣金"

                    if ~("df_value" in vars()):
                        df_value = t_df[["id", "订单号", "类型", "收入金额", "支出金额", "收支项目"]].copy()
                    else:
                        df_value = df_value.append(t_df[["id", "订单号", "类型", "收入金额", "支出金额", "收支项目"]])

                    return df_value
        else:
            pass

# 读取 人工退款
def read_ManualRefund(filename):
    df = pd.read_excel(filename, sheet_name="人工退款")
    df["id"] = df.index

    for col in df.columns:
        if col in ["支出金额"]:
            t_df = df[["id", "订单号", "支出金额"]]
            t_df["类型"] = "支出"
            t_df.columns = ["id", "订单号", "支出金额", "类型"]
            t_df["收入金额"] = "0"
            t_df["收支项目"] = "人工退款"
            # t_df = df[["id", "订单号", "支出金额", "类型"]]
            t_df["收支项目"] = "人工退款"
            t_df.loc[t_df.类型.str.contains("支出"), "支出金额"] = t_df["支出金额"]
            # 如果是收入
            t_df.loc[t_df.类型.str.contains("收入"), "收入金额"] = t_df["支出金额"]
        # elif col in ["佣金金额"]:
        #     t_df = df[["id", "订单号", "佣金金额"]]
        #     t_df["类型"] = "支出"
        #     t_df.columns = ["id", "订单号", "佣金金额", "类型"]
        #     t_df["收入金额"] = df["佣金金额"]
        #     t_df["收支项目"] = "人工退款佣金"
        #     # t_df = df[["id", "订单号", "支出金额", "类型"]]
        #     # 如果是支出
        #     t_df.loc[t_df.类型.str.contains("支出"), "支出金额"] = t_df["佣金金额"]
        #     # 如果是收入
        #     t_df.loc[t_df.类型.str.contains("收入"), "收入金额"] = t_df["佣金金额"]
        #     t_df["收支项目"] = "人工退款佣金"

            if ~("df_value" in vars()):
                df_value = t_df[["id", "订单号", "类型", "收入金额", "支出金额", "收支项目"]].copy()
            else:
                df_value = df_value.append(t_df[["id", "订单号", "类型", "收入金额", "支出金额", "收支项目"]])

    return df_value

# # 读取 人工退款-佣金
def read_ManualRefund_commission(filename):
    data_xls = pd.read_excel(filename, sheet_name=None, dtype=str)
    for key in data_xls:
        if key == "人工退款":
            df = pd.read_excel(filename, sheet_name="人工退款")
            df["id"] = df.index
            # left_df = df[["id", "订单号", "佣金金额"]]
    # df = pd.read_excel(filename, sheet_name="人工退款")
    # df["id"] = df.index

            for col in df.columns:
                if col in ["佣金金额"]:
                    t_df = df[["id", "订单号", "佣金金额"]]
                    t_df["类型"] = "支出"
                    t_df.columns = ["id", "订单号", "佣金金额", "类型"]
                    t_df["收入金额"] = df["佣金金额"]
                    t_df["收支项目"] = "人工退款佣金"

                    # t_df = df[["id", "订单号", "支出金额", "类型"]]

                    # 如果是支出
                    t_df.loc[t_df.类型.str.contains("支出"), "支出金额"] = t_df["佣金金额"]
                    # 如果是收入
                    t_df.loc[t_df.类型.str.contains("收入"), "收入金额"] = t_df["佣金金额"]

                    t_df["收支项目"] = "人工退款佣金"

                    if ~("df_value" in vars()):
                        df_value = t_df[["id", "订单号", "类型", "收入金额", "支出金额", "收支项目"]].copy()
                    else:
                        df_value = df_value.append(t_df[["id", "订单号", "类型", "收入金额", "支出金额", "收支项目"]])

                    return df_value
        else:
            pass

def judge_sheet(filename):
    data_xls = pd.read_excel(filename, sheet_name=None, dtype=str)
    for key in data_xls:
        if key == "订单运费":
            if 'df_h2' in locals().keys():  # 如果变量已经存在
                df_h2 = df_h2.append(read_express(filename)).append(read_express_Payment_channel_fee(filename))
            else:
                df_h2 = read_express(filename).append(read_express_Payment_channel_fee(filename))
        elif key == "赔偿用户薯券":
            if 'df_h2' in locals().keys():  # 如果变量已经存在
                df_h2 = df_h2.append(read_compensate(filename))
            else:
                df_h2 = read_compensate(filename)
        elif key == "人工调账":
            if 'df_h2' in locals().keys():  # 如果变量已经存在
                df_h2 = df_h2.append(read_adjust(filename))
            else:
                df_h2 = read_adjust(filename)
        elif key == "退货":
            if 'df_h2' in locals().keys():  # 如果变量已经存在
                df_h2 = df_h2.append(read_ReturnGoods(filename)).append(read_ReturnGoods_commission(filename)).append(read_ReturnGoods_Payment_channel_fee(filename))
            else:
                df_h2 = read_ReturnGoods(filename).append(read_ReturnGoods_commission(filename)).append(read_ReturnGoods_Payment_channel_fee(filename))
        elif key == "人工退款":
            if 'df_h2' in locals().keys():  # 如果变量已经存在
                df_h2 = df_h2.append(read_ManualRefund(filename)).append(read_ManualRefund_commission(filename))
            else:
                df_h2 = read_ManualRefund(filename).append(read_ManualRefund_commission(filename))
        elif key == "商品销售":
            if 'df_h2' in locals().keys():  # 如果变量已经存在
                df_h2 = df_h2.append(read_commission(filename)).append(read_Payment_channel_fee(filename))
            else:
                df_h2 = read_commission(filename).append(read_Payment_channel_fee(filename))
        # elif key == "人工退款":
        #     if 'df_h2' in locals().keys():  # 如果变量已经存在
        #         df_h2 = df_h2.append(read_ManualRefund_commission(filename))
        #     else:
        #         df_h2 = read_ManualRefund_commission(filename)
    return df_h2

# 表格拼接
def merge_settle(filename):

    # df_h2=read_express(filename).append(read_compensate(filename)).append(read_adjust(filename))

    left_df,df_value=read_settle(filename)
    df_h1= left_df.merge(df_value[["id","类型","收入金额","支出金额","收支项目"]],how="left",on="id")
    # df_h2=read_express(filename).append(read_compensate(filename)).append(read_adjust(filename)).append(read_ReturnGoods(filename)).append(read_ManualRefund(filename))
    df_h2=judge_sheet(filename)
    # "id","订单号","商品名称","商品id","发货时间","卖家名称"]]

    # 通过订单关联，获得发货时间，卖家名称信息
    df_h2=df_h2.merge(left_df[["订单号","发货时间","卖家名称"]],how="left",on="订单号")
    df_h2["商品名称"]=""
    df_h2["商品id"]=""
    
    # print("df_h1")
    # print(df_h1.to_markdown())
    
    # print("df_h2")
    # print(df_h2.to_markdown())
    
    
    df_sum=df_h1[["订单号","商品名称","商品id","发货时间","卖家名称","类型","收入金额","支出金额","收支项目"]].append(df_h2[["订单号","商品名称","商品id","发货时间","卖家名称","类型","收入金额","支出金额","收支项目"]])
    df_sum["卖家名称"].fillna(method="ffill",inplace=True)
    
    df_sum["订单号"].fillna("",inplace=True)
    df_sum["发货时间"].fillna("",inplace=True)
    df_sum["收入金额"].fillna("0",inplace=True)
    df_sum["支出金额"].fillna("0",inplace=True)
    
    df_sum.rename(columns={"发货时间":"结算时间"},inplace=True)
    
    
    return df_sum[["订单号","商品名称","商品id","结算时间","卖家名称","收支项目","收入金额","支出金额"]] 


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


def read_excel(filename):

    # if filename[-4:].find("xls")>0:
    #     temp_df = pd.read_excel(filename,dtype=str)
    # else:
    # print("wenjian:",filename)
    if filename.find("小红书")>0:
    #     if filename.find("张佳丽") > 0:
    #         temp_df = pd.read_csv(filename, dtype=str,encoding="gb18030", skiprows=4,skipfooter=4, sep = None)
    #         temp_df=temp_df[["业务订单号", "业务基础订单号","账务流水号", "发生时间", "业务描述","收入金额（+元）","支出金额（-元）"]]
    #
    #         temp_df.columns = ["业务订单号", "业务基础订单号", "账务流水号", "发生时间", "业务描述", "收入金额（+元）", "支出金额（-元）"]
    #         # print(temp_df.dtypes)
    #         print(temp_df.head(1).to_markdown())
    #         return temp_df
    #     else:
        temp_df = pd.read_excel(filename, dtype=str)       #去表头
    #     temp_df = pd.read_csv(filename, dtype=str, encoding="gb18030", sep=None)

    #     temp_df = pd.read_excel(filename, sheet_name = None )
    #     for i in temp_df.keys():
    #         if i == ["商品销售"]:
    #             temp_df = temp_df[["业务订单号", "业务基础订单号", "账务流水号", "发生时间", "业务描述", "收入金额（+元）", "支出金额（-元）"]]
    #         else:
    #             temp_df = temp_df[["业务订单号", "业务基础订单号", "账务流水号", "发生时间", "业务描述", "收入金额（+元）", "支出金额（-元）"]]


        print("—————————————————————————————打印数据————————————————————————")
        print(temp_df.head(10).to_markdown())
        # temp_df = temp_df[["业务订单号", "业务基础订单号", "账务流水号", "发生时间", "业务描述", "收入金额（+元）", "支出金额（-元）"]]
        # temp_df.columns = ["业务订单号", "业务基础订单号", "账务流水号", "发生时间", "业务描述", "收入金额（+元）", "支出金额（-元）"]
        # print(temp_df.dtypes)

        return temp_df

    # blank table
    df = pd.DataFrame({"id":"1"},index=[0])
    df = df.head(0)
    return df


def get_all_files(rootdir, filekey):
    filelist = list_all_files(rootdir, filekey)
    # print(filelist)
    mySeries = pd.Series(filelist)
    df = pd.DataFrame(mySeries)
    df.columns = ["filename"]
    # df["filename"] = df["filename"].apply(lambda x: x.strip().replace(" ", ""))

    # print(df.to_markdown())
    print(df)
    return df


def read_all_excel(rootdir, filekey):
    df_files = get_all_files(rootdir, filekey)
    for index, file in df_files.iterrows():
        if 'df' in locals().keys():  # 如果变量已经存在
            # dd = read_excel(file["filename"])
            dd = merge_settle(file["filename"])


            dd["filename"] = file["filename"]
            print("进度表：{}/{}  文件{}，行数{}".format(index + 1, df_files.shape[0], file["filename"], dd.shape[0]))
            df = df.append(dd)
        else:
            df = merge_settle(file["filename"])
            print("————————————————————测试——————————————————")
            print(type(df))
            # if isinstance(df, NoneType):
            # if df is None:
            if df.shape[0]:
                pass
            print(df.to_markdown())
            print(file["filename"])
            df["filename"] = file["filename"]
            print("进度表：{}/{}  文件{}，行数{}".format(index + 1, df_files.shape[0], file["filename"], df.shape[0]))

    return df


def combine_excel():
    print('请选择要合并的文件目录（请确定所有excel的表头都是相同的哦！）:')
    filedir=""
    # filedir = r"C:\Users\mega\Desktop\2019各平台数据原表\天猫、淘宝"
    filedir = r"C:\Users\mega\Desktop\2019线上店铺订单账单全\小红书"
    # input()
    # myTuple = shell.SHBrowseForFolder(0, None, "", 64)
    # try:
    #     path = shell.SHGetPathFromIDList(myTuple[0])
    # except:
    #     print("你没有输入任何目录 :(")
    #     sys.exit()
    #     return
    #
    # filedir=path.decode('ansi')
    print("你选择的路径是：", filedir)


    global default_dir
    default_dir = filedir

    if len(filedir) == 0:
        print("你没有输入任何目录 :(")
        sys.exit()
        return

    print('请输入要合并的文件名称或者后缀匹配符:（比如都必须包含 序时账 三个字，那么就请输入  "序时账" ， 不输入就表示所有excel都要合并！）')
    # filekey = "fee_"
    # filekey = "settle_"
    # filekey = "账务明细_1"
    filekey = "-账单.xlsx"
    # input()

    if len(filedir) == 0:
        print("你没有输入任何关键词 :(")
        filekey = ''
        # sys.exit()
        # return

    print("你希望在'{}'目录下找到所有的包含“{}”文件，然后合并。".format(filedir, filekey))

    table = read_all_excel(filedir, filekey)

    return table


def caiwu_xushizhang():
    df = combine_excel()
    if 'df' in locals().keys():  # 如果变量已经存在
        # print(df.head(10).to_markdown())
        print(df.head(3))
        # df.to_clipboard(index=False)
        print("ok1")
        if len(df)>500000:
            for i in range(0,int(len(df)/500000)+1):
                df[i*500000:(i+1)*500000].to_csv(default_dir + r"\小红书账单_{}.csv".format(i+1))
        else:
            df.to_excel(default_dir + r"\合并表格.xlsx")
        print("生成完毕，现在关闭吗？yes/no")
        byebye = input()
        print('bybye:', byebye)
    else:
        print("不好意思，什么也没有做哦 :(")
        # pyinstaller -p D:\Anaconda3\envs\duizhang -F .\xushizhang.py


if __name__ == "__main__":
    print("开始:", time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()))

    # filename=r"C:\Users\mega\Desktop\2019各平台数据原表\小红书\小红书Dentyl Active旗舰店\小红书Dentyl Active旗舰店201902-账单.xlsx"
    # df_sum=merge_settle(filename)
    # print(df_sum.to_markdown())

    caiwu_xushizhang()
    print("ok")

