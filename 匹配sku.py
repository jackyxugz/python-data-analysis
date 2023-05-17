# __coding=utf8__

import pandas as pd
import numpy as np
from collections import Counter
import re
import os
import time
import warnings

warnings.filterwarnings("ignore")


def get_sku(fn_file, oms_file):
    # 读取文件
    df1 = pd.read_pickle(fn_file)
    df2 = pd.read_pickle(oms_file)
    del df1["Unnamed: 0"]
    print(len(df1))
    print(df1.head().to_markdown())
    print(len(df2))
    print(df2.head().to_markdown())
    df1["实际卖出数量"] = df1["实际卖出数量"].astype(int)
    df1["商家编码"] = df1["商家编码"].astype(str)
    df1["支付日期"] = df1["支付日期"].astype("datetime64[ns]")
    df1["支付日期"] = df1["支付日期"].dt.date
    df2["实际卖出数量"] = df2["实际卖出数量"].astype(int)
    df2["商家编码"] = df2["商家编码"].astype(str)
    df2["支付日期"] = df2["支付日期"].astype("datetime64[ns]")
    df2["支付日期"] = df2["支付日期"].dt.date

    # df1 = df1[["账单主体","平台","订单店铺","支付日期","商家编码","总回款","总退款","实际收款","回款产品数量","退款产品数量","实际卖出数量","平均价格"]]
    # df2 = df2[["订单主体","平台","订单店铺","支付日期","商家编码","总回款","总退款","实际收款","回款产品数量","退款产品数量","实际卖出数量","平均价格"]]
    # print(len(df1))
    # print(df1.head().to_markdown())
    # print(len(df2))
    # print(df2.head().to_markdown())
    df1 = df1.groupby(["账单主体", "商家编码"]).agg({"实际卖出数量": "sum"})
    df1 = pd.DataFrame(df1).reset_index()
    df2 = df2.groupby(["订单主体", "商家编码"]).agg({"实际卖出数量": "sum"})
    df2 = pd.DataFrame(df2).reset_index()

    print(len(df1))
    print(df1.head().to_markdown())
    print(len(df2))
    print(df2.head().to_markdown())

    # df = pd.merge(df1,df2,how="outer",left_on=["账单主体","平台","订单店铺","支付日期","商家编码"],right_on=["订单主体","平台","订单店铺","支付日期","商家编码"])
    df = pd.merge(df1, df2, how="outer", left_on=["账单主体", "商家编码"], right_on=["订单主体", "商家编码"])
    df["商家编码"] = df["商家编码"].astype(str)
    print(len(df))
    print(df.head().to_markdown())

    df.to_excel(r"D:\沙井\财务需求\2020导出数据(账单匹配订单2022-1-14)16pm\匹配sku结果.xlsx", index=False)


def get_cover(fn_file):
    df = pd.read_excel(fn_file)
    df.to_pickle("data/2020年线上账单匹配订单-汇总.pkl")


if __name__ == "__main__":
    print("开始:", time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()))
    print("请按规定的路径存放成本文件：D:\沙井\匹配SKU\最新成本19号(1).xlsx")
    # fn_file = r"D:\沙井\财务需求\2020年线上账单匹配订单-汇总(3).xlsx"
    fn_file = r"data/2020年线上账单匹配订单-汇总.pkl"
    oms_file = "data/订单主体_支付日期合并表格.pkl"

    # get_cover(fn_file)
    get_sku(fn_file, oms_file)
