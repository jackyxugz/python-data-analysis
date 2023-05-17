# __coding=utf8__
# /** 作者：zengyanghui **/
import datetime

import pandas as pd
import numpy as np
import xlrd
import xlwt
import xlsxwriter


def read_jxc(filename1):
    df = pd.read_excel(filename1,dtype=str)
    df = df[["供应商名称","委外工厂名称","出入库日期","BOM物料编码","产品名称","入库数量"]]
    df["出入库日期"] = df["出入库日期"].astype("datetime64[D]")
    df["出入库日期"] = df["出入库日期"].dt.date
    # df["供应商-日期-编码"] = df["供应商名称"]+"-"+df["出入库日期"].astype(str)+"-"+df["BOM物料编码"]
    # df["加工厂-编码"] = df["委外工厂名称"]+"-"+df["BOM物料编码"]
    df = df.sort_values(["供应商名称","委外工厂名称","出入库日期","BOM物料编码","入库数量"])
    df["数据来源"] = "进销存文件"
    print(df.head().to_markdown())
    return df


def read_weiwai(filename2):
    df1 = pd.read_excel(filename2,sheet_name="对账单_20年1月-21年5月",dtype=str)
    df1 = df1[["供应商名称","加工厂", "送货日期", "BOM零件编码", "产品名称", "送货数量"]]
    df2 = pd.read_excel(filename2,sheet_name="对账单_21年6-10月",dtype=str)
    df2 = df2[["供应商名称","加工厂", "送货日期", "条码", "产品名称", "送货数量"]]
    df2.columns = ["供应商名称", "加工厂", "送货日期", "BOM零件编码", "产品名称", "送货数量"]
    df = pd.concat([df1,df2])
    # df = df[["加工厂","送货日期", "条码", "产品名称", "送货数量"]]
    df["送货日期"] = df["送货日期"].astype("datetime64[D]")
    df["送货日期"] = df["送货日期"].dt.date
    # df["供应商-日期-编码"] = df["供应商名称"] + "-" + df["送货日期"].astype(str) + "-" + df["BOM零件编码"]
    # df["加工厂-编码"] = df["加工厂"] + "-" + df["条码"]
    df = df.sort_values(["供应商名称","加工厂", "送货日期", "BOM零件编码", "送货数量"])
    df.columns = ["供应商名称","委外工厂名称","出入库日期","BOM物料编码","产品名称","入库数量"]
    df["数据来源"] = "委外文件"
    print(df.head().to_markdown())
    return df


def cover_file(filename1,filename2):
    df1 = read_jxc(filename1)
    df2 = read_weiwai(filename2)
    # df = pd.merge(df1,df2,how="right",on="供应商-日期-编码")
    df = pd.concat([df1,df2])

    print(df.head().to_markdown())

    df.to_excel(r"D:\沙井\委外\委外数据_detailResult.xlsx",index=False)


if __name__ == "__main__":
    filename1=r"D:\沙井\委外\供应商进销存汇总8月份_样本.xls"
    filename2=r"D:\沙井\委外\委外数据_all_V1.xlsx"
    cover_file(filename1,filename2)