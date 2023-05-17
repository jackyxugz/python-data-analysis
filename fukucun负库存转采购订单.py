#!/usr/bin/env python
# coding: utf-8

import pandas as pd
import numpy as np
import time
import xlrd
import tabulate
import openpyxl
import os
import sys
from tkinter import filedialog



# 设置文件对话框会显示的文件类型
my_filetypes = [ ('text excel files', '.xlsx'),('all excel files', '.xls')]


def sel_file(title_message):
    # 请求选择文件
    filename = filedialog.askopenfilename(initialdir=os.getcwd(),
                                          title=title_message,
                                          filetypes=my_filetypes)

    if len(filename) == 0:
        sys.exit()

    print("你选择的文件名是：", filename)
    return filename



def read_jxc(filename):
    if filename.find("xls")>0:
        # df=pd.read_excel(filename,sheet_name="Result 1")
        df=pd.read_excel(filename,dtype={"产品条码":object})
    else:
        df = pd.read_excel(filename, dtype={"产品条码":object})

    return df


def modify_jxc(df):
    df["产品条码"]=df["产品条码"].astype(str)
    # df=df[df["产品条码"].str.contains("4580444280108",na=True)]
    # df=df[df["产品条码"].str.contains("4580444280108",na=True)]

    df["产品条码"]=df.apply(lambda x: x["产品名称"] if len(str(x["产品条码"]).replace("nan",""))==0 else  x["产品条码"] ,axis=1 )

    print(df.head(10).to_markdown())

    df["id"]=df.index
    df["期末数量"].fillna(0, inplace=True)

    if "采购入库数量" in df.columns:
        df["采购入库数量"].fillna(0, inplace=True)

    if "销售出库数量" in df.columns:
        df["销售出库数量"].fillna(0, inplace=True)

    if "销售退货入库数量" in df.columns:
        df["销售退货入库数量"].fillna(0, inplace=True)

    if "采购退货出库数量" in df.columns:
        df["采购退货出库数量"].fillna(0, inplace=True)

    if "其他出库数量" in df.columns:
        df["其他出库数量"].fillna(0, inplace=True)

    error_item=df[df["期末数量"] < 0][["产品条码"]]
    # error_item=df[["产品条码"]]

    # print("test")
    # print(error_item.to_markdown())

    error_item=error_item["产品条码"].unique()
    # error_item = error_item["产品条码"].unique()
    error_item = pd.DataFrame(error_item).reset_index()
    error_item.columns = ["zero_index", "产品条码"]

    error_list=df.merge(error_item,how="left",on="产品条码")
    error_list["zero_index"].fillna(-1,inplace=True)
    error_list=error_list[error_list["zero_index"]>=0]

    # print("过滤数据")
    # print(error_list.to_markdown())

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
                sum_balance = row["期末数量"] + row_add

                print("抽查", row["本期开始时间"], row["期末数量"],row_add, old_balance, sum_balance)

                data = [{"id": row["id"], "本期开始时间": row["本期开始时间"],"仓库名称": row["仓库名称"],"产品条码": row["产品条码"],"调增":row_add,"结存":sum_balance}]
                df_temp = pd.DataFrame.from_dict(data, orient='columns')

            else:
                sum_balance = row["期末数量"]
                #  row["期初数量"]+ row["采购入库数量"] - row["销售出库数量"] + row["销售退货入库数量"] - row["采购退货出库数量"] - row[
                #                     "其他出库数量"]

                # print("抽查", row["本期开始时间"], row["期末数量"], 0, old_balance, sum_balance)


                data = [
                    {"id": row["id"], "本期开始时间": row["本期开始时间"], "仓库名称": row["仓库名称"], "产品条码": row["产品条码"], "调增": 0,
                     "结存": sum_balance}]
                df_temp = pd.DataFrame.from_dict(data, orient='columns')


        else:
            # print(row["期末数量"],old_balance)
            # 上期结存不够本期出库
            # 出库数量
            qty_out = old_balance - row["期末数量"]
            if ((   qty_out > sum_balance ) & (row["期末数量"]<0)):
                # print("发现负库存2", row["期末数量"])
                row_add = qty_out-sum_balance  # 54
                # old_balance = row["期末数量"]
                # 结存
                sum_balance = sum_balance + row["采购入库数量"] - row["销售出库数量"]

                if "销售退货入库数量" in error_list.columns:
                    sum_balance = sum_balance + row["销售退货入库数量"]

                if "采购退货出库数量" in error_list.columns:
                    sum_balance = sum_balance - row["采购退货出库数量"]

                if "其他出库数量" in error_list.columns:
                    sum_balance = sum_balance - row["其他出库数量"]

                sum_balance = sum_balance + row_add

                # xiugai.append([row["id"],row["本期开始时间"], row["仓库名称"], row["产品条码"], row_add])
                data = [{"id": row["id"], "本期开始时间": row["本期开始时间"], "仓库名称": row["仓库名称"], "产品条码": row["产品条码"],"调增":row_add,"结存":sum_balance}]
                df_temp = pd.DataFrame.from_dict(data, orient='columns')

                # print("抽查", row["本期开始时间"], row["期末数量"], row_add,old_balance, sum_balance)

            else:   # 当期不是负库存
                sum_balance = sum_balance + row["采购入库数量"] - row["销售出库数量"]
                    #           + row["销售退货入库数量"] - row["采购退货出库数量"] - row[
                    # "其他出库数量"] + 0

                if "销售退货入库数量" in error_list.columns:
                    sum_balance = sum_balance + row["销售退货入库数量"]

                if "采购退货出库数量" in error_list.columns:
                    sum_balance = sum_balance - row["采购退货出库数量"]

                if "其他出库数量" in error_list.columns:
                    sum_balance = sum_balance - row["其他出库数量"]


                # print("抽查", row["本期开始时间"], row["期末数量"], 0, old_balance, sum_balance)


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




    # 请求选择一个用以保存的文件
    # save_file = filedialog.asksaveasfilename(initialdir=os.getcwd(),
    #                                          title="请输入要保存的结果文件名:",
    #                                          filetypes=my_filetypes)
    #
    # if len(save_file)>0:
    #     if save_file.find("xls") < 0:
    #         save_file = save_file + ".xlsx"
    #     print(save_file)
    #     df.to_excel(save_file)
    #     print("文件 {} 已经生成".format(save_file))
    # else:
    #     print("你没有选择文件!")
    # input("请按2次回车退出！")

    df.to_excel(r"D:\沙井\匹配SKU\19年进销存明细_卖家联合_0103_01_20220104_result.xlsx",index=False)

if __name__ == "__main__":
    print("开始:", time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()))
    filename=r"D:\沙井\匹配SKU\19年进销存移动库存明细__卖家联合20220104(1).xls"
    # filename = sel_file("请选择要转的excel文件")
    df=read_jxc(filename)
    modify_jxc(df)
    print("ok")
