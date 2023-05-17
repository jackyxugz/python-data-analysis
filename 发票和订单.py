# __coding=utf8__
# /** 作者：zengyanghui **/

import pandas as pd
import numpy as np
import xlrd
import xlwt
import xlsxwriter


def read_inv(filename1):
    df = pd.read_excel(filename1,dtype=str)
    df["金额不含税"] = df["金额不含税"].astype(float)
    df["税额"] = df["税额"].astype(float)
    df["含税金额"] = df["金额不含税"]+df["税额"]
    # group_df = df.groupby(["发票号码"]).agg({"含税金额": "sum"})
    # group_df = pd.DataFrame(group_df).reset_index()
    # group_df.columns = ["发票号码", "含税金额-汇总"]
    # df1 = pd.merge(df,group_df,how="left",on="发票号码")
    return df


def read_order(filename2):
    df = pd.read_excel(filename2,dtype=str)
    df["销售已交货含税金额"].fillna(0,inplace=True)
    df["销售已交货含税金额"] = df["销售已交货含税金额"].astype(float)
    df = df[df["销售已交货含税金额"]>0]
    # group_df = df.groupby(["销售订单"]).agg({"销售已交货含税金额":"sum"})
    # group_df = pd.DataFrame(group_df).reset_index()
    # group_df.columns = ["销售订单","销售已交货含税金额-汇总"]
    # df1 = pd.merge(df,group_df,how="left",on="销售订单")
    return df


def merge_io(filename1,filename2):
    df1 = read_inv(filename1)
    df2 = read_order(filename2)
    df = pd.merge(df1,df2,how="outer",left_on="含税金额",right_on="销售已交货含税金额")
    print(df.head().to_markdown())
    df.to_excel(r"D:\沙井\匹配SKU\发票-订单2.xlsx",index=False)



if __name__ == "__main__":
    filename1 = r"D:\沙井\匹配SKU\2019年开票汇总-.xlsx"
    filename2 = r"D:\沙井\匹配SKU\19年麦凯莱库存移动明细.xlsx"

    merge_io(filename1,filename2)