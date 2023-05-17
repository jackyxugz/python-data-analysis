# encoding:utf-8

import sys
import os
import pandas as pd
import numpy as np
import time
import tabulate
import xlrd
import openpyxl
from tkinter import filedialog


# 设置文件对话框会显示的文件类型
my_filetypes = [ ('text excel files', '.xlsx'),('all excel files', '.xls')]

# filename=r"C:\Users\ns2033\Downloads\代销与出货.xlsx"
# filename=r"work/代销与出货.xlsx"

month_list=["202009","202010","202011","202012","202101","202102","202103","202104","202105","202106","202107","202108","202109","202110","202111","202112"]
def read_out(filename):
    df=pd.read_excel(filename,sheet_name="出货")
    df.rename(columns={"出货月份":"yearmonth","商品码": "sku", "求和项:入库数量": "qty_in"}, inplace=True)
    return df

def read_sales(filename):
    df=pd.read_excel(filename,sheet_name="已销")
    df.rename(columns={"代销月份":"yearmonth","商品号":"sku","销售数量":"qty_out"},inplace=True)
    df=df.groupby(["yearmonth","sku"]).agg({"qty_out":np.sum})
    df=pd.DataFrame(df).reset_index()
    df.columns=["yearmonth","sku","qty_out"]
    return df


def next_month(tmonth):
    month_list=["202009","202010","202011","202012","202101","202102","202103","202104","202105","202106","202107","202108","202109","202110","202111","202112"]
    i=0
    j=0
    # print("查找:",tmonth)
    for m in month_list:
        if m==tmonth:
            j=i
            # print("j:", i)

        i=i+1

    # print("index:", j+1)
    # print("yearmonth:", month_list[j+1])
    return month_list[j+1]



def sel_file(title_message):
    # 请求选择文件
    filename = filedialog.askopenfilename(initialdir=os.getcwd(),
                                          title=title_message,
                                          filetypes=my_filetypes)

    if len(filename) == 0:
        sys.exit()

    print("你选择的文件名是：", filename)
    return filename

def trans_sc(filename):


    df1=read_out(filename)
    df2=read_sales(filename)

    left_sku=df1["sku"].unique()
    left_sku=pd.DataFrame(left_sku).reset_index()
    left_sku.columns=["index","sku"]
    left_sku["id"]=1
    del left_sku["index"]

    left_month = df1["yearmonth"].unique()
    left_month = pd.DataFrame(left_month).reset_index()
    left_month.columns = ["index", "yearmonth"]
    left_month["id"] = 1
    del left_month["index"]

    # left_month=pd.DataFrame(month_list[:7])
    # left_month.columns=["yearmonth"]
    # left_month["id"] = 1

    left_df=left_sku.merge(left_month,how="left",on=["id"])
    print(left_df)


    del left_df["id"]

    # print(left_sku)
    # print(left_month)
    # print(left_df)

    # sys.exit()
    left_df["yearmonth"]=left_df["yearmonth"].astype(str)
    left_df["sku"]=left_df["sku"].astype(str)

    df1["yearmonth"] = df1["yearmonth"].astype("str")
    df1["sku"] = df1["sku"].astype(str)

    df1=left_df.merge(df1,how="left",on=["yearmonth","sku"])

    print(df1.head(5).to_markdown())

    df1["nextmonth1"]=df1.apply(lambda x:  next_month(x["yearmonth"]) ,axis=1)
    df1["nextmonth2"]=df1.apply(lambda x:  next_month(x["nextmonth1"]) ,axis=1)
    df1["nextmonth3"]=df1.apply(lambda x:  next_month(x["nextmonth2"]) ,axis=1)
    df1["nextmonth4"]=df1.apply(lambda x:  next_month(x["nextmonth3"]) ,axis=1)
    df1["nextmonth5"]=df1.apply(lambda x:  next_month(x["nextmonth4"]) ,axis=1)
    df1["nextmonth6"]=df1.apply(lambda x:  next_month(x["nextmonth5"]) ,axis=1)

    df2["yearmonth"] = df2["yearmonth"].astype("str")

    # 先计算入库
    df3 = df1.merge(df1[["yearmonth", "sku", "qty_in"]], how="left", left_on=["nextmonth1", "sku"],
                    right_on=["yearmonth", "sku"])
    df3.rename(columns={"qty_in_x": "qty_in0", "qty_in_y": "qty_in1", "yearmonth_x": "yearmonth"}, inplace=True)
    del df3["yearmonth_y"]

    df3 = df3.merge(df1[["yearmonth", "sku", "qty_in"]], how="left", left_on=["nextmonth2", "sku"],
                    right_on=["yearmonth", "sku"])
    df3.rename(columns={"qty_in": "qty_in2" , "yearmonth_x": "yearmonth"}, inplace=True)
    del df3["yearmonth_y"]

    df3 = df3.merge(df1[["yearmonth", "sku", "qty_in"]], how="left", left_on=["nextmonth3", "sku"],
                    right_on=["yearmonth", "sku"])
    df3.rename(columns={"qty_in": "qty_in3", "yearmonth_x": "yearmonth"}, inplace=True)
    del df3["yearmonth_y"]

    df3 = df3.merge(df1[["yearmonth", "sku", "qty_in"]], how="left", left_on=["nextmonth4", "sku"],
                    right_on=["yearmonth", "sku"])
    df3.rename(columns={"qty_in": "qty_in4", "yearmonth_x": "yearmonth"}, inplace=True)
    del df3["yearmonth_y"]

    df3 = df3.merge(df1[["yearmonth", "sku", "qty_in"]], how="left", left_on=["nextmonth5", "sku"],
                    right_on=["yearmonth", "sku"])
    df3.rename(columns={"qty_in": "qty_in5", "yearmonth_x": "yearmonth"}, inplace=True)
    del df3["yearmonth_y"]

    df3 = df3.merge(df1[["yearmonth", "sku", "qty_in"]], how="left", left_on=["nextmonth6", "sku"],
                    right_on=["yearmonth", "sku"])
    df3.rename(columns={"qty_in": "qty_in6", "yearmonth_x": "yearmonth"}, inplace=True)
    del df3["yearmonth_y"]

    # print("debug_0")
    # print(df3.head(10).to_markdown())

    df3["qty_in0"].fillna(0,inplace=True)
    df3["qty_in1"].fillna(0,inplace=True)
    df3["qty_in2"].fillna(0,inplace=True)
    df3["qty_in3"].fillna(0,inplace=True)
    df3["qty_in4"].fillna(0,inplace=True)
    df3["qty_in5"].fillna(0,inplace=True)
    df3["qty_in6"].fillna(0,inplace=True)

    # sys.exit()
    # 在计算出库

    df2["yearmonth"] = df2["yearmonth"].astype("str")
    df2["sku"] = df2["sku"].astype(str)

    df3=df3.merge(df2[["yearmonth","sku","qty_out"]],how="left",left_on=["yearmonth","sku"],right_on=["yearmonth","sku"])
    # df3.rename(columns={ "yearmonth_x": "yearmonth"}, inplace=True)
    print(df3.head(10).to_markdown())

    # del df3["yearmonth_y"]
    df3=df3.merge(df2[["yearmonth","sku","qty_out"]],how="left",left_on=["nextmonth1","sku"],right_on=["yearmonth","sku"])
    df3.rename(columns={"qty_out_x":"qty_out0","qty_out_y":"qty_out1","yearmonth_x":"yearmonth"},inplace=True)
    df3.rename(columns={"yearmonth_x": "yearmonth"}, inplace=True)
    del df3["yearmonth_y"]


    # print("debug_1")
    # print(df3.head(10).to_markdown())

    df3 = df3.merge(df2[["yearmonth", "sku", "qty_out"]], how="left", left_on=["nextmonth2", "sku"], right_on=["yearmonth", "sku"])
    df3.rename(columns={"yearmonth_x": "yearmonth"}, inplace=True)
    del df3["yearmonth_y"]
    df3 = df3.merge(df2[["yearmonth", "sku", "qty_out"]], how="left", left_on=["nextmonth3", "sku"], right_on=["yearmonth", "sku"])
    df3.rename(columns={"qty_out_x": "qty_out2", "qty_out_y": "qty_out3"}, inplace=True)
    df3.rename(columns={"yearmonth_x": "yearmonth"}, inplace=True)
    del df3["yearmonth_y"]


    # print("debug_2")
    # print(df3.head(10).to_markdown())

    df3 = df3.merge(df2[["yearmonth", "sku", "qty_out"]], how="left", left_on=["nextmonth4", "sku"], right_on=["yearmonth", "sku"])
    df3.rename(columns={"yearmonth_x": "yearmonth"}, inplace=True)
    del df3["yearmonth_y"]
    df3 = df3.merge(df2[["yearmonth", "sku", "qty_out"]], how="left", left_on=["nextmonth5", "sku"], right_on=["yearmonth", "sku"])
    df3.rename(columns={"yearmonth_x": "yearmonth"}, inplace=True)
    del df3["yearmonth_y"]
    df3.rename(columns={"qty_out_x": "qty_out4", "qty_out_y": "qty_out5"}, inplace=True)

    # print("debug_3")
    # print(df3.head(10).to_markdown())

    df3 = df3.merge(df2[["yearmonth", "sku", "qty_out"]], how="left", left_on=["nextmonth6", "sku"], right_on=["yearmonth", "sku"])
    df3.rename(columns={"yearmonth_x": "yearmonth"}, inplace=True)
    del df3["yearmonth_y"]
    df3.rename(columns={"qty_out": "qty_out6"}, inplace=True)

    df3["qty_out0"].fillna(0, inplace=True)
    df3["qty_out1"].fillna(0, inplace=True)
    df3["qty_out2"].fillna(0, inplace=True)
    df3["qty_out3"].fillna(0, inplace=True)
    df3["qty_out4"].fillna(0, inplace=True)
    df3["qty_out5"].fillna(0, inplace=True)
    df3["qty_out6"].fillna(0, inplace=True)

    # df3["in_sum"] = 0
    # df3["out_sum"] = 0
    in_sum_list=[]
    out_sum_list=[]
    # print("testtest")
    # print(df3.head(10).to_markdown())

    df_temp=df3.copy()

    for index,rrow in df3.iterrows():
        # print(rrow)
        sku=rrow["sku"]
        yearmonth=int(rrow["yearmonth"])
        # print("sku=",sku)
        # print("yearmonth=",yearmonth)
        # 小于当月的所有入库和出库汇总
        df_sum=df_temp[  (df_temp["sku"].astype(str)==sku) & (df_temp["yearmonth"].astype(int)<=yearmonth) ].groupby("sku").agg({"qty_in0":np.sum,"qty_out0":np.sum})
        df_sum=pd.DataFrame(df_sum).reset_index()
        # print("ceshi:")
        # print(df_sum)
        df_sum.columns=["sku","in_sum","out_sum"]
        # 截至当月，累计入库数量
        in_sum=df_sum["in_sum"].iloc[0]
        # 截至当月，累计出库数量
        out_sum=df_sum["out_sum"].iloc[0]

        in_sum_list.append(in_sum)
        out_sum_list.append(out_sum)

        rrow["in_sum"]=in_sum
        rrow["out_sum"]=out_sum

    # print("入库列表")
    # print(in_sum_list)
    # print("出库列表")
    # print(out_sum_list)

    in_sum_list=pd.DataFrame(in_sum_list).reset_index()
    out_sum_list=pd.DataFrame(out_sum_list).reset_index()

    in_sum_list.columns=["index","qty_in_sum"]
    out_sum_list.columns=["index","qty_out_sum"]

    # print(in_sum_list)
    # print(out_sum_list)

    # df3 = pd.concat([df3, in_sum_list], axis=1, ignore_index=True)
    # df3 = pd.concat([df3, out_sum_list], axis=1, ignore_index=True)

    df3["index"]=df3.index
    df3=df3.merge(in_sum_list,how="left",on=["index"])
    df3=df3.merge(out_sum_list,how="left",on=["index"])

    # 当月库存余额
    df3["qty_balance"]=df3["qty_in_sum"]-df3["qty_out_sum"]

    del df3["nextmonth1"]
    del df3["nextmonth2"]
    del df3["nextmonth3"]
    del df3["nextmonth4"]
    del df3["nextmonth5"]
    del df3["nextmonth6"]
    del df3["index"]

    print("更新后结果:")
    print(df3.head(10).to_markdown())


    # df3["202009"]=0
    # df3["202010"]=0
    # df3["202011"]=0
    # df3["202012"]=0
    # df3["202101"]=0
    # df3["202102"]=0
    # df3["202103"]=0
    #
    # df3["stock0"] = 0
    # df3["stock1"] = 0
    # df3["stock2"] = 0
    # df3["stock3"] = 0
    # df3["stock4"] = 0
    # df3["stock5"] = 0
    # df3["stock6"] = 0

    # df3.loc[df3["出货月份"].str.contains("202009")]["202009"]=df3["销售数量_当月"]
    # df3.loc[df3["出货月份"].str.contains("202009")]["202010"]=df3["销售数量_次月"]
    # df3.loc[df3["出货月份"].str.contains("202009")]["202011"]=df3["销售数量_下下月"]
    # df3.loc[df3["出货月份"].str.contains("202009")]["202012"]=df3["销售数量_下下下月"]
    # df3.loc[df3["出货月份"].str.contains("202009")]["202101"]=df3["销售数量_下下下下月"]
    # df3.loc[df3["出货月份"].str.contains("202009")]["202102"]=df3["销售数量_下下下下下月"]
    # df3.loc[df3["出货月份"].str.contains("202009")]["202103"]=df3["销售数量_下下下下下下月"]
    #
    # df3.loc[df3["出货月份"].str.contains("202010")]["202010"] = df3["销售数量_当月"]
    # df3.loc[df3["出货月份"].str.contains("202010")]["202011"] = df3["销售数量_次月"]
    # df3.loc[df3["出货月份"].str.contains("202010")]["202012"] = df3["销售数量_下下月"]
    # df3.loc[df3["出货月份"].str.contains("202010")]["202101"] = df3["销售数量_下下下月"]
    # df3.loc[df3["出货月份"].str.contains("202010")]["202102"] = df3["销售数量_下下下下月"]
    # df3.loc[df3["出货月份"].str.contains("202010")]["202103"] = df3["销售数量_下下下下下月"]
    #
    # df3.loc[df3["出货月份"].str.contains("202011")]["202011"] = df3["销售数量_当月"]
    # df3.loc[df3["出货月份"].str.contains("202011")]["202012"] = df3["销售数量_次月"]
    # df3.loc[df3["出货月份"].str.contains("202011")]["202101"] = df3["销售数量_下下月"]
    # df3.loc[df3["出货月份"].str.contains("202011")]["202102"] = df3["销售数量_下下下月"]
    # df3.loc[df3["出货月份"].str.contains("202011")]["202103"] = df3["销售数量_下下下下月"]
    #
    # df3.loc[df3["出货月份"].str.contains("202012")]["202012"] = df3["销售数量_当月"]
    # df3.loc[df3["出货月份"].str.contains("202012")]["202101"] = df3["销售数量_次月"]
    # df3.loc[df3["出货月份"].str.contains("202012")]["202102"] = df3["销售数量_下下月"]
    # df3.loc[df3["出货月份"].str.contains("202012")]["202103"] = df3["销售数量_下下下月"]
    #
    # df3.loc[df3["出货月份"].str.contains("202101")]["202101"] = df3["销售数量_当月"]
    # df3.loc[df3["出货月份"].str.contains("202101")]["202102"] = df3["销售数量_次月"]
    # df3.loc[df3["出货月份"].str.contains("202101")]["202103"] = df3["销售数量_下下月"]
    #
    # df3.loc[df3["出货月份"].str.contains("202102")]["202102"] = df3["销售数量_当月"]
    # df3.loc[df3["出货月份"].str.contains("202102")]["202103"] = df3["销售数量_次月"]
    #
    # df3.loc[df3["出货月份"].str.contains("202103")]["202103"] = df3["销售数量_当月"]

    df3.rename(columns={"qty_in0":"当月入库","qty_out0":"当月出库","qty_in_sum":"累计入库","qty_out_sum":"累计出库","qty_balance":"期末库存"},inplace=True)

    # print(df3)
    print(df3.head(10).to_markdown())
    return  df3[["sku","yearmonth","求和项:供货总额（不含税）","求和项:含税总额","当月入库","qty_in1","qty_in2","qty_in3","qty_in4","qty_in5","qty_in6","当月出库","qty_out1","qty_out2","qty_out3","qty_out4","qty_out5","qty_out6","累计入库","累计出库","期末库存"]]


if __name__ == "__main__":
    print("开始:", time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()))

    filename = sel_file("请选择要转的excel文件")
    df=trans_sc(filename)


    # 请求选择一个用以保存的文件
    save_file = filedialog.asksaveasfilename(initialdir=os.getcwd(),
                                             title="请输入要保存的结果文件名:",
                                             filetypes=my_filetypes)

    if len(save_file)>0:
        if save_file.find("xls") < 0:
            save_file = save_file + ".xlsx"
        print(save_file)
        df.to_excel(save_file)
        print("文件 {} 已经生成".format(save_file))
    else:
        print("你没有选择文件!")
    input("请按2次回车退出！")

    # pyinstaller -p D:\Anaconda3\envs\duizhang -F .\ziranxianjinxianchu.py
    # x=next_month("202012")
    # print(x)

