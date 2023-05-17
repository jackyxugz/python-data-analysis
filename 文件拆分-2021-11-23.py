# __coding=utf8__
# /** 作者：zengyanghui **/

import pandas as pd
import numpy as np
from collections import Counter
import re
import os
import warnings

warnings.filterwarnings("ignore")
# import Tkinter
# import win32api
# import win32ui
# import win32con
import tabulate
import math
import datetime as dt
import uuid

def get_order(default_path,file):
    # 1.打开文件
    # filename = open_file()
    filename = default_path + os.sep + file
    df = pd.read_excel(filename,sheet_name="对账单", dtype=str, na_values=False)
    print("读取order文件成功，打印前5行数据：")
    print(f"\n原文件行数：{len(df)}")
    print(df.head(5).to_markdown())

    # 一、格式清理
    # 1.删除表头所有空格、换行符
    for column_name in df.columns:
        df.rename(columns={column_name: column_name.replace(" ", "").replace("\n", "").strip()}, inplace=True)

    # 2.列数据格式化
    df["送货数量"] = df["送货数量"].astype(float)
    df["含税单价"] = df["含税单价"].astype(float)
    df["委外类别"] = df["委外类别"].astype(str)

    df["送货日期"] = df["送货日期"].astype("datetime64[ns]")
    df["送货日期"] = df["送货日期"].dt.date

    # 二、数据分类
    # 1、采购订单：数量大于0的全部数据，如果"类别（OEM加工厂/包材）"值为"OEM加工厂"，需要拼接"委外类别"的值
    df1 = df
    df1 = df1.loc[df["送货数量"] > 0]
    df1["类别（OEM加工厂-供应商自采/包材）"] = df1.apply(lambda x:x["类别（OEM加工厂/包材）"] + "-" + x["委外类别"] if x["类别（OEM加工厂/包材）"]=="OEM加工厂" else x["类别（OEM加工厂/包材）"],axis=1)
    df1["单位"] = df1["计量单位"]
    df1 = df1[["序号","单号","供应商名称","公司主体","订单日期","采购员","产品编码","BOM零件编码","单位","送货数量","含税单价","类别（OEM加工厂-供应商自采/包材）"]]
    print(f"\ndf1:采购订单行数{len(df1)}")
    print(df1.head(5).to_markdown())
    df1.to_excel(default_path + "/采购订单.xlsx",index=False)

    # 2、入库单确认：数量大于0并且"委外类别"值不等于"发原料"，按送货日期排序
    df2 = df
    df2 = df2.loc[df["送货数量"] > 0]
    df2 = df2[~df2["委外类别"].str.contains("发原料")]
    df2 = df2[["序号","单号","产品编码","BOM零件编码","送货数量","送货日期"]]
    df2 = df2.sort_values(by=["送货日期"])
    print(f"\ndf2:入库单确认行数{len(df2)}")
    print(df2.head(5).to_markdown())
    df2.to_excel(default_path + "/入库单确认.xlsx", index=False)

    # 3、调拨单入委外：数量大于0并且"类别（OEM加工厂/包材）"="包材｜原料"的数据，按送货日期排序
    df3 = df
    df3 = df3.loc[df["送货数量"] > 0]
    df3 = df3[df3["类别（OEM加工厂/包材）"].str.contains("包材|原料")]
    df3 = df3[["序号","公司主体","送货日期","产品编码","BOM零件编码","送货数量","产品单位","加工厂"]]
    df3 = df3.sort_values(by=["送货日期"])
    print(f"\ndf3:调拨单入委外行数{len(df3)}")
    print(df3.head(5).to_markdown())
    df3.to_excel(default_path + "/调拨单入委外.xlsx", index=False)

    # 4、委外生产单：数量大于0并且"类别（OEM加工厂/包材）"="OEM加工厂"的数据，按送货日期排序
    df4 = df
    df4 = df4.loc[df["送货数量"] > 0]
    df4 = df4[df4["类别（OEM加工厂/包材）"].str.contains("OEM加工厂")]
    df4 = df4[["序号","公司主体","送货日期","采购员","产品编码","产品单位","送货数量","单号","加工厂"]]
    df4 = df4.sort_values(by=["送货日期"])
    print(f"\ndf4:委外生产单行数{len(df4)}")
    print(df4.head(5).to_markdown())
    df4.to_excel(default_path + "/委外生产单.xlsx", index=False)

    # 5、委外加工费："BOM零件编码"="PROCESSCOST-001"的数据
    df5 = df
    df5 = df5.loc[df["BOM零件编码"] == "PROCESSCOST-001"]
    # df5 = df5[df5["BOM零件编码"].str.contains("PROCESSCOST-001")]
    df5 = df5[["序号","BOM零件编码","含税单价"]]
    print(f"\ndf5:委外加工费行数{len(df5)}")
    print(df5.head(5).to_markdown())
    df5.to_excel(default_path + "/委外加工费.xlsx", index=False)


if __name__ == "__main__":
    # print("请选择所要拆分的文件所在路径：")
    # default_path = input()
    # print("请选择所要拆分的文件名称（包含文件后缀）：")
    # file = input()
    file = "委外数据20211020(2)(1).xlsx"
    default_path = r"/Users/maclove/Downloads/拆分委外数据"
    get_order(default_path,file)
    print("ok")


