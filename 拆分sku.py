import sys
import os
import pandas as pd
import numpy as np
import tabulate
import openpyxl
import win32api
import win32ui
import win32con
import win32com


def open_file():
    dlg = win32ui.CreateFileDialog(1)  # 1表示打开文件对话框
    dlg.SetOFNInitialDir('D:/')  # 设置打开文件对话框中的初始显示目录
    dlg.DoModal()
    filename = dlg.GetPathName()  # 获取选择的文件名称
    print("filename=", filename)
    print("read ok")
    return filename


def split_sku(file,sheet):
    df = pd.read_excel(file,sheet_name=sheet)
    print(df.head(1).to_markdown())
    print("原表总行数：{}".format(len(df)))
    if "商品条码" in df.columns:
        # df["商品条码"] = df["商品条码"].replace("*","X")
        df["商品条码"] = df["商品条码"].astype(str)
        # df0是原表不带+|X和特殊编码的单条码数据
        df0 = df.loc[(~df["商品条码"].str.contains("\+"))&(~df["商品条码"].str.contains("X"))]
        print("原表不带+|X的总行数：{}".format(len(df0)))

        # df1是原表带-|+|X和特殊编码的单条码或者组合条码数据
        df1 = df.loc[(df["商品条码"].str.contains("\+"))|(df["商品条码"].str.contains("X"))]
        df1["商品条码1"] = df1["商品条码"].apply(lambda x:"nan" if ((x.find("X")>0)&(x.find("+")<0)) else x)
        print("原表带-的总行数：{}".format(len(df1)))

        # df2是原表带X的单条码数据
        df2 = df1.loc[df1["商品条码1"].str.contains("nan")]
        del df2["商品条码1"]
        print("原表带X的单条码总行数：{}".format(len(df2)))

        # df3是原表带+的组合条码数据
        df3 = df1.loc[~df1["商品条码1"].str.contains("nan")]
        if len(df3)>0:
            del df3["商品条码1"]
            print("原表带+的组合条码总行数：{}".format(len(df3)))
            # 把商品条码根据+进行拆分
            df3_split = df3["商品条码"].str.split("+", expand=True)
            # 旋转行和列
            df3_split = df3_split.stack()
            # 重置index
            df3_split = df3_split.reset_index()
            # 统计组合商品条码拆分后的数量(行数)
            df3_count = pd.DataFrame(df3_split["level_0"].value_counts()).reset_index()
            # 列重命名，df3_split的索引=level_0，拆分的数量=level_2
            df3_count.columns = ["level_0", "level_2"]
            # 重置索引index = level_0
            df3_split = df3_split.set_index("level_0")
            # 列重命名，商品条码的拆分序号=level_1，拆分后的单个商品条码=商品条码
            df3_split.columns = ["level_1", "商品条码"]
            # 合并拆分的商品条码表和统计商品条码数量表
            df3_split = df3_split.merge(df3_count, how="left", on="level_0")
            # 重置索引index = level_0
            df3_split = df3_split.set_index("level_0")

            # print(df3.head().to_markdown())
            # print(df3_split.head().to_markdown())

            # df3表删掉商品条码列后，左联df3_split拆分后的商品条码表
            df3_new = df3.drop(["商品条码"], axis=1).join(df3_split)

            print(df3_new.head().to_markdown())
            print("原表带+|X的组合条码拆分后总行数：{}".format(len(df3_new)))

            # 重新计算单价和含税金额
            df3_new["单价（含税）"] = df3_new["单价（含税）"].astype(float) / df3_new["level_2"].astype(int)
            df3_new["含税金额"] = df3_new["单价（含税）"].astype(float) * df3_new["数量"].astype(int)

            df3 = pd.concat([df0, df2, df3_new])

            # df4表为多单条码X多个数量
            df4 = df3[df3["商品条码"].str.contains("X")]
            print(df4.head().to_markdown())
            df4["cnt"] = df4["商品条码"].apply(lambda x:x[x.find("X") +1:])
            df4["商品条码"] = df4["商品条码"].apply(lambda x:x[:x.find("X")])
            df4_new = df4[["商品条码","cnt"]]
            df4_new["index"] = df4_new.index
            print(df4_new.head().to_markdown())
            print("拆分后仍带X的单条码总行数：{}".format(len(df4_new)))

            df4_count = pd.read_excel("data/sku_qty.xlsx")
            df4_new["cnt"] = df4_new["cnt"].astype(str)
            df4_count["cnt"] = df4_count["cnt"].astype(str)
            df4_new = pd.merge(df4_new,df4_count,how="left",on="cnt")
            df4_new = df4_new.set_index("index")
            print(df4_new.head().to_markdown())
            print(df4.head().to_markdown())

            del df4["cnt"]
            df4_new = df4.drop(["商品条码"], axis=1).join(df4_new)
            print(df4_new.head().to_markdown())

            # 重新计算单价和含税金额
            df4_new["单价（含税）"] = df4_new["单价（含税）"].astype(float) / df4_new["cnt"].astype(int)
            df4_new["含税金额"] = df4_new["单价（含税）"].astype(float) * df4_new["数量"].astype(int)

            # 合并所有表
            df3 = df3[~df3["商品条码"].str.contains("X")]
            df5 = pd.concat([df3,df4_new])

            print(df5.tail().to_markdown())
            print("原表总行数：{}".format(len(df)))
            print("处理后总行数：{}".format(len(df5)))

            return df5

        else:
            ms = win32api.MessageBox(0, "此文件没有需要拆分的sku！", "提醒", win32con.MB_OK)
            sys.exit()

    else:
        ms = win32api.MessageBox(0, "此文件没有商品条码字段！请检查选择的文件或者输入的sheet名是否有误！", "提醒", win32con.MB_OK)

    print(df)

def read_file():
    print("请选择需要拆分sku的文件")
    file = open_file()
    print(file)
    df = pd.read_excel(file, sheet_name=None, dtype=str)
    sheetlist = list(df)
    print(sheetlist)

    if len(sheetlist) > 1:
        print(len(sheetlist))
        print("请输入需要拆分sku的sheet页")
        sheet = input()
        df = split_sku(file,sheet)
    else:
        df = split_sku(file,df)

    out_file = os.sep.join(file.split(".")[:-1])
    df.to_excel(out_file + "-" + sheet + ".xlsx",index=False)
    print("处理完毕！")

    byby = input()

    print("再见！")


if __name__ == "__main__":

    read_file()
