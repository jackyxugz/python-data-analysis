# __coding=utf8__
# /** 作者：zengyanghui **/

import pandas as pd
import numpy as np
import xlrd
import xlwt
import xlsxwriter


def read_excel(filename,sheet):
    df = pd.read_excel(filename,sheet_name=sheet)
    df["期末数量"] = df["期末数量"].astype(float)
    df.fillna(0,inplace=True)
    # df = df[df["期末数量"]<0]
    df = df.sort_values(["产品条码","本期开始时间","仓库名称"])
    return df


def modify_excel(filename,sheet):
    df = read_excel(filename,sheet)
    print(df.head(10).to_markdown())
    df["iid"] = df.index
    df["出现序号"] = df.groupby(df["产品条码"])["iid"].rank(method='dense')
    df_count = pd.DataFrame(df["产品条码"].value_counts()).reset_index()
    df_count.columns = ["产品条码", "总序号"]
    df1 = df.merge(df_count, how="left", on="产品条码")
    del df["iid"]
    print((df1.head(10).to_markdown()))
    # df1["期末库存1"] = df1.apply(lambda x:x["期初数量"]+x["采购入库数量"]-x["销售出库数量"]+x["销售退货入库数量"]-x["采购退货出库数量"]-x["其他出库数量"] if x["出现序号"]==1 else x["采购入库数量"]-x["销售出库数量"]+x["销售退货入库数量"]-x["采购退货出库数量"]-x["其他出库数量"],axis=1)
    # df1["期末库存1"] = df1["期初数量"]+df1["采购入库数量"]-df1["销售出库数量"]+df1["销售退货入库数量"]-df1["采购退货出库数量"]-df1["其他出库数量"]
    df1["期末库存1"] = df1.apply(lambda x:x["期初数量"]+x["采购入库数量"]-x["销售出库数量"]+x["销售退货入库数量"]-x["采购退货出库数量"]-x["其他出库数量"] if x["期初数量"]>0 else x["采购入库数量"]-x["销售出库数量"]+x["销售退货入库数量"]-x["采购退货出库数量"]-x["其他出库数量"],axis=1)
    df_1 = df1.shift(1)
    print((df1.head(10).to_markdown()))
    df1["需补采购数量1"] = df1.apply(lambda x:-x["期末库存1"] if x["期末库存1"]<0 else 0,axis=1)
    df1["产品条码"] = df1["产品条码"].astype(str)
    print((df1.head(10).to_markdown()))
    df1.to_excel(r"D:\沙井\匹配SKU\19年进销存明细_卖家联合_0103_01_需补采购明细_已补录.xlsx",index=False)


if __name__ == "__main__":

    filename = r"D:\沙井\匹配SKU\19年进销存明细_卖家联合_0103_01_需补采购明细.xlsx"
    sheet = "总表"
    # df = read_excel(filename,sheet)
    modify_excel(filename,sheet)