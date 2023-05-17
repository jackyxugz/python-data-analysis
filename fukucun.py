#!/usr/bin/env python
# coding: utf-8

import pandas as pd
import numpy as np
import time
import xlrd
import tabulate
import openpyxl



def read_jxc(filename):
    df=pd.read_excel(filename,sheet_name="总表")
    return df


def modify_jxc(df):
    df["产品条码"]=df["产品条码"].astype(str)
    # df=df[df["产品条码"].str.contains("4580444280108",na=True)]

    print(df.head(10).to_markdown())

    df["id"]=df.index
    df["期末数量"].fillna(0, inplace=True)

    df["采购入库数量"].fillna(0, inplace=True)
    df["销售出库数量"].fillna(0, inplace=True)
    df["销售退货入库数量"].fillna(0, inplace=True)
    df["采购退货出库数量"].fillna(0, inplace=True)
    df["其他出库数量"].fillna(0, inplace=True)

    error_item=df[df["期末数量"] < 0][["产品条码"]]

    print("test")
    print(error_item.to_markdown())


    error_item=error_item["产品条码"].unique()
    # error_item = error_item["产品条码"].unique()
    error_item = pd.DataFrame(error_item).reset_index()
    error_item.columns = ["zero_index", "产品条码"]

    error_list=df.merge(error_item,how="left",on="产品条码")
    error_list["zero_index"].fillna(-1, inplace=True)
    error_list=error_list[error_list["zero_index"]>=0]

    print("过滤数据")
    print(error_list.to_markdown())

    error_list = error_list.sort_values(["产品条码", "仓库名称", "本期开始时间"])

    xiugai = []
    old_sku = ""
    old_balance = 0
    sum_balance=0

    for index, row in error_list.iterrows():
        sku = row["产品条码"]
        # print(old_sku , sku)
        if old_sku != sku:
            if int( row["期末数量"]) < 0:  # -54
                # print("发现负库存1",old_balance)
                row_add=- row["期末数量"]  #  54
                # sum_add = sum_add + row_add
                # xiugai.append([row["id"],row["本期开始时间"], row["仓库名称"], row["产品条码"], row_add])

                sum_balance = row["期末数量"] + row_add

                print("抽查", row["本期开始时间"], row["期末数量"],row_add, old_balance, sum_balance)

                data = [{"id": row["id"], "本期开始时间": row["本期开始时间"],"仓库名称": row["仓库名称"],"产品条码": row["产品条码"],"调增":row_add,"结存":sum_balance}]
                df_temp = pd.DataFrame.from_dict(data, orient='columns')

                if "df_sum" in vars():
                    df_sum=df_sum.append(df_temp)
                else:
                    df_sum =  df_temp
            else:
                sum_balance = row["期末数量"]
                #  row["期初数量"]+ row["采购入库数量"] - row["销售出库数量"] + row["销售退货入库数量"] - row["采购退货出库数量"] - row[
                #                     "其他出库数量"]

                data = [
                    {"id": row["id"], "本期开始时间": row["本期开始时间"], "仓库名称": row["仓库名称"], "产品条码": row["产品条码"], "调增": 0,
                     "结存": sum_balance}]
                df_temp = pd.DataFrame.from_dict(data, orient='columns')

                if "df_sum" in vars():
                    df_sum = df_sum.append(df_temp)
                else:
                    df_sum = df_temp

        else:
            print(row["期末数量"],old_balance)
            # 上期结存不够本期出库
            # 出库数量
            qty_out = old_balance - row["期末数量"]
            if ((   qty_out > sum_balance ) & (row["期末数量"]<0)):
                # print("发现负库存2", row["期末数量"])


                row_add = qty_out-sum_balance  # 54
                # old_balance = row["期末数量"]

                # 结存
                sum_balance = sum_balance + row["采购入库数量"] - row["销售出库数量"] + row["销售退货入库数量"] - row["采购退货出库数量"] - row[
                    "其他出库数量"] + row_add

                # xiugai.append([row["id"],row["本期开始时间"], row["仓库名称"], row["产品条码"], row_add])
                data = [{"id": row["id"], "本期开始时间": row["本期开始时间"], "仓库名称": row["仓库名称"], "产品条码": row["产品条码"],"调增":row_add,"结存":sum_balance}]
                df_temp = pd.DataFrame.from_dict(data, orient='columns')

                print("抽查", row["本期开始时间"], row["期末数量"], row_add,old_balance, sum_balance)


                if "df_sum" in vars():
                    df_sum = df_sum.append(df_temp)
                else:
                    df_sum = df_temp
            else:   # 当期不是负库存
                sum_balance = sum_balance + row["采购入库数量"] - row["销售出库数量"] + row["销售退货入库数量"] - row["采购退货出库数量"] - row[
                    "其他出库数量"] + 0
                data = [
                    {"id": row["id"], "本期开始时间": row["本期开始时间"], "仓库名称": row["仓库名称"], "产品条码": row["产品条码"], "调增": 0,
                     "结存": sum_balance}]
                df_temp = pd.DataFrame.from_dict(data, orient='columns')

                if "df_sum" in vars():
                    df_sum = df_sum.append(df_temp)
                else:
                    df_sum = df_temp


        old_sku=sku
        old_balance = row["期末数量"]

    print("查看问题:")
    # print(df_temp.head(10).to_markdown())
    print(df_sum.head(10).to_markdown())
    # xiugai=pd.DataFrame(xiugai,columns=["id","本期开始时间","仓库名称","产品条码","调增"],index=["id"])

    df=df.merge(df_sum[["id","调增","结存"]],how="left",on="id")
    df.to_excel(r"D:\沙井\匹配SKU\19年进销存明细_卖家联合_0103_01_cc_修改1111.xlsx")

if __name__ == "__main__":
    print("开始:", time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()))
    filename=r"D:\沙井\匹配SKU\19年进销存明细_卖家联合_0103_01_需补采购明细.xlsx"
    df=read_jxc(filename)
    modify_jxc(df)
    print("ok")
