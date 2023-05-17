#!/usr/bin/env python
# coding: utf-8

import pandas as pd
import numpy as np
import time
import xlrd
import tabulate
import openpyxl
import os
from tkinter import filedialog
import sys


# 设置文件对话框会显示的文件类型
my_filetypes = [ ('text excel files', '.xlsx'),('all excel files', '.xls')]

debug_sku="8809647230014"
download_path=r"/Users/lichunlei/Downloads/"

def sel_file(title_message):
    # 请求选择文件
    filename = filedialog.askopenfilename(initialdir=os.getcwd(),
                                          title=title_message,
                                          filetypes=my_filetypes)

    if len(filename) == 0:
        sys.exit()

    print("你选择的文件名是：", filename)
    return filename


def pipei_12(fn_jxc,fn_tiaoma):
    # 统计发货数量/金额，未匹配数量/金额
    df_jxc,df_title=read_jxc(fn_jxc)
    df_chuhuo=read_chuhuo(fn_jxc)
    # 重复合并,sku为空
    #  ,"title"
    # df=df_jxc.merge(df_chuhuo,how="left",on=["sono","refno","sku"])
    # df.columns = ["sono", "refno", "sku", "quty", "amnt", "cnt"]

    df_jxc.rename(columns={"quty":"quty_jxc","amnt":"amnt_jxc"},inplace=True)

    df = df_jxc.merge(df_chuhuo[["sono", "sku", "quty", "amnt"]], how="left", on=["sono" , "sku"])
    print("跟踪匹配0")
    print(df.head(10).to_markdown())

    df = df.merge(df_chuhuo[[ "refno","sku", "quty", "amnt"]], how="left", on=["refno" , "sku"])

    df["quty_x"].fillna(0,inplace=True)
    df["quty_y"].fillna(0,inplace=True)
    df["amnt_x"].fillna(0, inplace=True)
    df["amnt_y"].fillna(0, inplace=True)

    print("跟踪匹配1")
    print(df.head(10).to_markdown())

    df["quty_ch"]=df.apply(lambda x:  x["quty_x"]  if x["quty_x"]>0 else x["quty_y"] ,axis=1 )
    df["amnt_ch"]=df.apply(lambda x:  x["amnt_x"]  if x["amnt_x"]>0 else x["amnt_y"] ,axis=1 )


    df["quty_error"]=df["quty_jxc"]-df["quty_ch"]
    df["amnt_error"]=df["amnt_jxc"]-df["amnt_ch"]

    print("跟踪匹配2")
    print(df.head(10).to_markdown())

    df=df[(df["quty_jxc"]>0) & (df["amnt_jxc"]>0)]
    # df=df[(df["quty_x"]>0) & (df["amnt_x"]>0)]

    df = df[~ ((df["quty_error"] > 0) & (df["amnt_error"] == 0))]
    df = df[~((df["quty_error"] == 0) & (df["amnt_error"] > 0))]

    print("跟踪匹配3")
    print(df.head(10).to_markdown())
    df.to_excel(download_path+"进销存匹配出货的结果.xlsx")

    # ,"title"
    df1 = df.groupby(["sku"]).agg({"quty_jxc": np.sum, "amnt_jxc": np.sum})
    df1=pd.DataFrame(df1).reset_index()
    #  "title",
    df1.columns=["sku","quty","amnt"]
    df1["类型"] = "进销存出货"

    # ,"title"
    df2 = df.groupby(["sku"]).agg({"quty_ch": np.sum, "amnt_ch": np.sum})
    df2 = pd.DataFrame(df2).reset_index()
    # "title",
    df2.columns = ["sku", "quty", "amnt"]
    df2["类型"] = "进销存未匹配"

    return df1,df2,df_title

def pipei_3(fn_jxc, fn_tiaoma):

    # 生成发票减折扣
    df_fapiao=read_fapiao(fn_jxc)
    # df_fapiao.rename(columns={"商品名称":"title"},inplace=True)

    # 给发票补充sku编号
    # df_sku=read_sku(fn_tiaoma)
    # df_fapiao=df_fapiao.merge(df_sku,how="left",on=["title"])

    df_fapiao_nosku=df_fapiao[df_fapiao["sku"].str.len()==0]
    print("发票缺少sku信息的有:")
    print(df_fapiao_nosku.head(10).to_markdown())
    # df_fapiao_nosku.to_markdown("异常.xlsx")

    df_fapiao_ok= df_fapiao[df_fapiao["sku"].str.len() > 0]
    df_fapiao_ok=df_fapiao_ok.groupby(["sku"]).agg({"quty":np.sum,"amnt":np.sum})
    df_fapiao_ok=pd.DataFrame(df_fapiao_ok).reset_index()
    df_fapiao_ok.columns=["sku","quty","amnt"]

    df_fapiao_ok["类型"] = "发票"

    if len(debug_sku)>0:
        print("抽查发票")
        print(df_fapiao_ok[df_fapiao_ok["sku"].str.contains(debug_sku)].to_markdown())

    return df_fapiao_ok

def pipei_sum(fn_jxc, fn_tiaoma):

    df1,df2,df_title=pipei_12(fn_jxc, fn_tiaoma)
    df_fapiao=pipei_3(fn_jxc, fn_tiaoma)

    print("出货，记录数")
    print(df1.shape[0])

    print("未匹配，记录数")
    print(df2.shape[0])


    df_2019=read_fachu_2019(fn_jxc)
    df_2019["类型"] = "2019发出商品"

    print("2019发出商品，记录数")
    print(df_2019.shape[0])

    df_2020=read_fachu_2020(fn_jxc)
    df_2020["类型"]="2020发出商品"

    print("2020发出商品，记录数")
    print(df_2020.shape[0])

    # 5项拼接
    df_sum=df1.append(df2).append(df_fapiao).append(df_2019).append(df_2020)
    del df_sum["cnt"]

    print("合并拼接结果:")
    print(df_sum.head(10).to_markdown())

    # 只保留长度最长的名称
    # df_left=df_sum[["sku","title"]].copy()
    # df_left["title"]=df_left["title"].astype(str)
    # df_left["title_length"]=df_left["title"].apply(lambda x: len(x))
    # df_left=df_left.sort_values(by=["sku","title_length"])
    # df_left=df_left.drop_duplicates(subset=["sku"],keep="last")
    # del df_left["title_length"]
    df_left=df_title[["sku","title"]].copy()

    # 统计出所有涉及的SKU，作为左连接的依据
    # df_left=df_sum.groupby(["sku","title"]).agg({"sku":["count"]})
    # df_left=pd.DataFrame(df_left).reset_index()
    # df_left.columns=["sku","title","count"]
    # del df_left["count"]

    print("左侧的产品信息:")
    print(df_left.head(10).to_markdown())

    # 表格横向拼接
    df_result=df_left.merge(df_sum[df_sum["类型"].str.contains("发票")][["sku","quty","amnt"]],how="left",on="sku")
    df_result=df_result.merge(df_sum[df_sum["类型"].str.contains("2019发出商品")][["sku","quty","amnt"]],how="left",on="sku")

    df_result.rename(columns={"quty_x":"发票_数量","quty_y":"2019发出商品_数量","amnt_x":"发票_价税合计","amnt_y":"2019发出商品_价税合计"},inplace=True)

    if len(debug_sku) > 0:
        print("抽查发票2")
        print(df_result[df_result["sku"].str.contains(debug_sku, na=False)].to_markdown())

    print("抽查1:")
    print(df_result.head(10).to_markdown())

    df_result=df_result.merge(df_sum[df_sum["类型"].str.contains("2020发出商品")][["sku","quty","amnt"]],how="left",on="sku")
    df_result = df_result.merge(df_sum[df_sum["类型"].str.contains("进销存出货")][["sku", "quty", "amnt"]],how="left",on="sku")

    df_result.rename(columns={"quty_x": "2020发出商品_数量", "quty_y": "进销存出货_数量", "amnt_x": "2020发出商品_价税合计", "amnt_y": "进销存出货_价税合计"},
                     inplace=True)



    print("抽查2:")
    print(df_result.head(10).to_markdown())
    if len(debug_sku) > 0:
        print("抽查发票3")
        print(df_result[df_result["sku"].str.contains(debug_sku, na=False)].to_markdown())


    df_result = df_result.merge(df_sum[df_sum["类型"].str.contains("进销存未匹配")][["sku", "quty", "amnt"]],how="left",on="sku")
    df_result.rename(
        columns={"quty": "进销存未匹配_数量",  "amnt": "进销存未匹配_价税合计"},
        inplace=True)

    df_result["sku"]=df_result["sku"].astype(str)

    print("抽查3:")
    print(df_result.head(10).to_markdown())

    df_result["发票_单价"]=df_result["发票_价税合计"]/df_result["发票_数量"]
    df_result["2019发出商品_单价"]=df_result["2019发出商品_价税合计"]/df_result["2019发出商品_数量"]
    df_result["2020发出商品_单价"]=df_result["2020发出商品_价税合计"]/df_result["2020发出商品_数量"]
    df_result["进销存出货_单价"]=df_result["进销存出货_价税合计"]/df_result["进销存出货_数量"]
    df_result["进销存未匹配_单价"]=df_result["进销存未匹配_价税合计"]/df_result["进销存未匹配_数量"]

    df_result["发票_单价"].fillna(0, inplace=True)
    df_result["2019发出商品_单价"].fillna(0, inplace=True)
    df_result["2020发出商品_单价"].fillna(0, inplace=True)
    df_result["进销存出货_单价"].fillna(0, inplace=True)
    df_result["进销存未匹配_单价"].fillna(0, inplace=True)

    df_result["发票_数量"].fillna(0, inplace=True)
    df_result["发票_价税合计"].fillna(0, inplace=True)
    df_result["2019发出商品_数量"].fillna(0, inplace=True)
    df_result["2019发出商品_价税合计"].fillna(0, inplace=True)
    df_result["2020发出商品_数量"].fillna(0, inplace=True)
    df_result["2020发出商品_价税合计"].fillna(0, inplace=True)
    df_result["进销存出货_数量"].fillna(0, inplace=True)
    df_result["进销存出货_价税合计"].fillna(0, inplace=True)
    df_result["进销存未匹配_数量"].fillna(0, inplace=True)
    df_result["进销存未匹配_价税合计"].fillna(0, inplace=True)

    print("抽查4:")
    print(df_result.head(10).to_markdown())

    df_result["平均单价"]=df_result.apply(lambda x: cal_price_avg(x),axis=1 )
    df_result["最大单价"]=df_result.apply(lambda x: cal_price_max(x) ,axis=1 )
    df_result["最小单价"]=df_result.apply(lambda x: cal_price_min(x) ,axis=1 )
    df_result["单价标准差"]=df_result.apply(lambda x: cal_price_var(x),axis=1 )

    df_result=df_result[["sku","title","发票_数量","发票_价税合计","发票_单价","2019发出商品_数量","2019发出商品_价税合计","2019发出商品_单价","2020发出商品_数量","2020发出商品_价税合计","2020发出商品_单价","进销存出货_数量","进销存出货_价税合计","进销存出货_单价","进销存未匹配_数量","进销存未匹配_价税合计","进销存未匹配_单价" ,"最小单价","平均单价","最大单价","单价标准差"]]

    print(df_result.head(10).to_markdown())


    return df_result

def cal_price_avg(x):
    price_list = [x["发票_单价"], x["2019发出商品_单价"], x["2020发出商品_单价"], x["进销存出货_单价"], x["进销存未匹配_单价"]]
    # 剔除价格为0的元素
    # price_list.remove(0)
    price_list = [x for x in price_list if x != 0]
    # return   sum(price_list)/len(price_list)
    return   np.mean(price_list)

def cal_price_max(x):
    price_list = [x["发票_单价"], x["2019发出商品_单价"], x["2020发出商品_单价"], x["进销存出货_单价"], x["进销存未匹配_单价"]]
    # 剔除价格为0的元素
    # price_list.remove(0)
    price_list = [x for x in price_list if x != 0]
    if len(price_list)>0:
        return max(price_list)
    else:
        return  0

def cal_price_min(x):
    price_list = [x["发票_单价"], x["2019发出商品_单价"], x["2020发出商品_单价"], x["进销存出货_单价"], x["进销存未匹配_单价"]]
    # 剔除价格为0的元素
    # price_list.remove(0)
    price_list = [x for x in price_list if x != 0]
    if len(price_list) > 0:
        return min(price_list)
    else:
        return  0

def cal_price_var(x):
    price_list = [x["发票_单价"], x["2019发出商品_单价"], x["2020发出商品_单价"], x["进销存出货_单价"], x["进销存未匹配_单价"]]
    # 剔除价格为0的元素
    # price_list.remove(0)
    price_list = [ int(x) for x in price_list if x != 0]
    if len(price_list) > 0:
        return np.var(price_list)
    else:
        return  0

def read_jxc(filename):
    df=pd.read_excel(filename,sheet_name="进销存")
    print(df.head(5).to_markdown())
    df=df[["源单据","销售明细行/订单关联/客户参考","产品/条码","产品","完成数量","销售明细行/小计","销售明细行/税额总计" ]]
    df.columns=["sono","refno","sku","title","quty","amnt_notax","tax"]
    # 价税合计
    df["amnt"]=df["amnt_notax"]+df["tax"]

    df["sono"].fillna("", inplace=True)
    df["refno"].fillna("", inplace=True)

    # 取最长的名称当做产品名
    df["title"] = df["title"].astype(str)
    df["title_length"] = df["title"].apply(lambda x: len(x))
    df_title = df.sort_values(by=["sku", "title_length"])
    df_title = df_title.drop_duplicates(subset=["sku"], keep="last")
    df_title=df_title[["sku","title"]]

    df=df.groupby(["sono","refno","sku"]).agg({"quty":np.sum,"amnt":np.sum,"sku":["count"]})
    df=pd.DataFrame(df).reset_index()
    df.columns=["sono","refno","sku","quty","amnt","cnt"]


    # df=df.merge(df_title[["sku","title"]],how="left",on="sku")

    df["sono"]=df["sono"].astype(str)
    df["refno"]=df["refno"].astype(str)
    df["sku"]=df["sku"].astype(str)
    # df["title"]=df["title"].astype(str)

    print("进销存有重复记录的有：")
    print(df[df["cnt"]>1].to_markdown())
    df[df["cnt"]>1].to_excel(download_path+"进销存有重复记录.xlsx")

    # return df[df["cnt"]==1]
    return df,df_title

def read_chuhuo(filename):
    df=pd.read_excel(filename,sheet_name="出货")
    # df=df[["订单号","客户参考","条码","出货数量","出货金额"]]
    df=df[["订单号（审批单号）" ,"条形码","商品名称","实发数量" ,"含税总额"]]
    df.columns = ["no", "sku","title","quty","amnt"]
    df["no"]=df["no"].astype(str)
    df["sono"]=df["no"].apply(lambda x: x  if x.find("ML/")>=0 else ""  )
    df["refno"]=df["no"].apply(lambda x: ""  if x.find("ML/")>=0 else x  )

    df["sono"].fillna("",inplace=True)
    df["refno"].fillna("",inplace=True)

    # , "title"
    df = df.groupby(["sono", "refno", "sku"]).agg({"quty": np.sum, "amnt": np.sum, "sku": ["count"]})
    df = pd.DataFrame(df).reset_index()
    df.columns = ["sono", "refno", "sku", "quty", "amnt", "cnt"]

    df["sono"] = df["sono"].astype(str)
    df["refno"] = df["refno"].astype(str)
    df["sku"] = df["sku"].astype(str)
    # df["title"] = df["title"].astype(str)

    print("出货有重复记录的有：")
    print(df[df["cnt"] > 1].to_markdown())
    df[df["cnt"] > 1].to_excel(download_path+"出货有重复记录.xlsx")


    # return df[df["cnt"]==1]
    return df

def read_fachu_2019(filename):
    df=pd.read_excel(filename,sheet_name="2019发出商品")
    # df=df[["订单号","客户参考","条码","出货数量","出货金额"]]
    df=df[["条形码","商品名称" ,"实发数量" ,"送货含税总额"]]
    df.columns = ["sku", "title","quty","amnt"]

    df = df.groupby([ "sku" ]).agg({"quty": np.sum, "amnt": np.sum, "sku": ["count"]})
    df = pd.DataFrame(df).reset_index()
    df.columns = ["sku",  "quty", "amnt", "cnt"]

    df["sku"]=df["sku"].astype(str)


    if len(debug_sku)>0:
        print("抽查2019")
        print(df[df["sku"].str.contains(debug_sku,na=False)].to_markdown())

    # print("2019发出商品重复记录的有：")
    # print(df[df["cnt"] > 1].to_markdown())

    # return df[df["cnt"]==1]
    return df

def read_fachu_2020(filename):
    df=pd.read_excel(filename,sheet_name="2020发出商品")
    # df=df[["订单号","客户参考","条码","出货数量","出货金额"]]
    df=df[["条形码","商品名称" ,"实发数量" ,"含税总额"]]
    df.columns = ["sku", "title","quty","amnt"]

    df = df.groupby(["sku" ]).agg({"quty": np.sum, "amnt": np.sum, "sku": ["count"]})
    df = pd.DataFrame(df).reset_index()
    df.columns = ["sku",  "quty", "amnt", "cnt"]

    df["sku"] = df["sku"].astype(str)

    if len(debug_sku) > 0:
        print("抽查2020")
        print(df[df["sku"].str.contains(debug_sku,na=False)].to_markdown())

    # print("2020发出商品有重复记录的有：")
    # print(df[df["cnt"] > 1].to_markdown())

    # return df[df["cnt"]==1]
    return df

def read_fapiao(filename):
    df = pd.read_excel(filename, sheet_name="发票")
    # df_fapiao=df_fapiao[~(df_fapiao["数量"]==0 & df_fapiao["金额"]<0)]
    df = df[(df["数量"] != 0)]

    df = df[["商品条码", "商品名称", "数量", "合计"]]
    df.columns = ["sku", "title", "quty", "amnt"]

    df = df.groupby(["sku", "title"]).agg({"quty": np.sum, "amnt": np.sum, "sku": ["count"]})
    df = pd.DataFrame(df).reset_index()
    df.columns = ["sku", "title", "quty", "amnt", "cnt"]
    df["sku"] = df["sku"].astype(str)

    print("发票有重复记录的有：")
    print(df[df["cnt"] > 1].to_markdown())
    df[df["cnt"] > 1].to_excel(download_path+"发票有重复记录.xlsx")


    # return df[df["cnt"] == 1]
    return df

def read_sku(filename):
    df1=pd.read_excel(filename,sheet_name="2019年")
    df1=df1[["商品名称","条码"]]

    df2 = pd.read_excel(filename, sheet_name="2020")
    df2 = df2[["商品名称", "条码"]]

    df3 = pd.read_excel(filename, sheet_name="2021")
    df3 = df3[["商品名称", "条码"]]

    df=df1.append(df2).append(df3)
    df=df.groupby(["商品名称","条码"]).agg({"条码":["count"]})
    df=pd.DataFrame(df).reset_index()
    df.columns=["title","sku","count"]

    return  df[["title","sku"]]

if __name__ == "__main__":
    print("开始:", time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()))
    fn_jxc="/Users/lichunlei/PycharmProjects/kis2cloud/data/xuchengqing/华润万家 深圳   出货开票  匹配表--2020年度 （欠1.5.12月开票出货） 2021.12.13 OK.xlsx"
    fn_tiaoma="/Users/lichunlei/PycharmProjects/kis2cloud/data/xuchengqing/2019-2021销项条码.xlsx"
    # filename = sel_file("请选择要转的excel文件")
    df=pipei_sum(fn_jxc,fn_tiaoma)
    # df.to_excel(r"/Users/mac/Downloads/合并结算结果.xlsx")
    df.to_excel(download_path+"合并结算结果.xlsx")
    print("ok")


