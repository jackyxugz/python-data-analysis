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
from string import printable
import re
import operator


# 设置文件对话框会显示的文件类型
my_filetypes = [ ('text excel files', '.xlsx'),('all excel files', '.xls')]


def del_noprintchar(the_string):
    # new_string = re.sub("[^{}]+".format(printable), "", the_string)
    new_string = ''.join(char for char in the_string if char in printable)
    # new_string=the_string[:len(the_string)]+"_lichunlei"
    return new_string

def read_kis(filename):

    # 注意去除字符串的前置单引号  '
    df_ori=pd.read_excel(filename) # ,sheet_name="kis云会计分录序时簿"
    df_ori["vouch_iid"] = df_ori.index

    df=df_ori.copy()
    df_ori.fillna("", inplace=True)

    df["日期"].fillna(method="ffill", inplace=True)
    df["会计期间"].fillna(method="ffill",inplace=True)
    df["凭证字号"].fillna(method="ffill", inplace=True)

    df["科目名称"]=df["科目名称"].astype(str)
    df["科目名称"]=df["科目名称"].apply(lambda x: "".join(x.replace("\r","").replace("\n","").replace("_x000D_","").strip())  )
    # 确定凭证序号

    df["凭证字"] = df["凭证字号"].astype(str)
    df["凭证字"] = df["凭证字号"].apply(lambda x: x.split("-")[0].strip() )
    df["凭证号"] = df["凭证字号"].apply(lambda x: x.replace("-", "").replace(" ", "").replace("记", ""))
    # 会计年度	期间
    df["会计期间"] = df["会计期间"].astype(str)

    # df["会计年度"] = df["会计期间"].apply(lambda x: x.split(".")[0])
    # df["期间"] = df["会计期间"].apply(lambda x: x.split(".")[1])

    df["日期"] = df["日期"].astype(str)
    df["日期"] = df["日期"].apply(lambda x: x.replace("-", "/"))

    df["会计年度"] = df["日期"].astype("datetime64[ns]").dt.year
    df["期间"] = df["日期"].astype("datetime64[ns]").dt.month

    df["vouchrank"] = df.apply(
        lambda x: "{}{:0>2d}{:0>6d}".format(int(x["会计年度"]), int(x["期间"]), int(x["凭证号"])), axis=1)
    # 分组内部建立序号 1,2,3
    df["iseq"] = df["vouch_iid"].groupby(df["vouchrank"]).rank(method='first', ascending=True)

    df.fillna("", inplace=True)

    # print("debug 111")
    # print(df.to_markdown())
    # df.to_excel("work/test.xlsx")

    df["借方金额"]=df["借方"].apply(lambda x:   ''  if  abs(float(x))<0.00001  else x )
    df["贷方金额"]=df["贷方"].apply(lambda x:  ''  if  abs(float(x))<0.00001  else x )

    # print("debug222")
    # print(df.head(10).to_markdown())

    df["科目代码"]=df["科目代码"].astype(str)
    df["科目代码"] = df["科目代码"].apply(lambda x: kemu_clear(x))


    # 去掉合计
    k=df.shape[0]
    df["科目代码"]=df["科目代码"].astype(str)
    df=df[df["科目代码"].str.len()>0]
    print("去掉合计，删除了{}行".format(k-df.shape[0]))

    # print("debug333")
    # print(df.head(10).to_markdown())

    # print("debug 2")
    # print(df.head(10).to_markdown())

    # df = df.fillna(method='ffill')
    # df=df[["vouch_iid","vouchno","iseq","日期","会计期间","凭证字号","摘要","科目代码","科目名称","币别","汇率","原币金额","借方","贷方"]]

#
    return  df[["审核","复核","过账","日期","会计年度","会计期间","期间","凭证字号","凭证字","凭证号","摘要","科目代码","科目名称","币别","汇率","原币金额","借方","贷方","借方金额","贷方金额" ,"vouchrank","iseq"]]
        # df_ori[["审核","复核","过账","日期","会计期间","凭证字号","摘要","科目代码","科目名称","币别","汇率","原币金额","借方","贷方"]],\



def expand_codename(df,code,name,sept_char):
    # 获得科目全称,例如  应收帐款_客户_沃尔玛

    # 支持到5级科目
    df=df[["iid",code,name]].copy()

    df["科目级次"] = df[code].apply(lambda x: x.count(".") + 1)
    df["parent_code"]=""  # 上级科目   （共2级）
    df["parent_code2"]="" # 上上级     （共3级）
    df["parent_code3"]="" # 上上上级   （共4级）
    df["parent_code4"]="" # 上上上上级 （共5级）

    df.loc[df["科目级次"] == 2, "parent_code"] = df[code].apply(lambda x: ".".join(x.split(".")[:1]))

    df.loc[df["科目级次"] == 3, "parent_code"] = df[code].apply(lambda x: ".".join(x.split(".")[:2]))
    df.loc[df["科目级次"] == 3, "parent_code2"] = df[code].apply(lambda x: ".".join(x.split(".")[:1]))

    df.loc[df["科目级次"] == 4, "parent_code"] = df[code].apply(lambda x: ".".join(x.split(".")[:3]))
    df.loc[df["科目级次"] == 4, "parent_code2"] = df[code].apply(lambda x: ".".join(x.split(".")[:2]))
    df.loc[df["科目级次"] == 4, "parent_code3"] = df[code].apply(lambda x: ".".join(x.split(".")[:1]))

    # 5级科目
    df.loc[df["科目级次"] == 5, "parent_code"] = df[code].apply(lambda x: ".".join(x.split(".")[:4]))
    df.loc[df["科目级次"] == 5, "parent_code2"] = df[code].apply(lambda x: ".".join(x.split(".")[:3]))
    df.loc[df["科目级次"] == 5, "parent_code3"] = df[code].apply(lambda x: ".".join(x.split(".")[:2]))
    df.loc[df["科目级次"] == 5, "parent_code4"] = df[code].apply(lambda x: ".".join(x.split(".")[:1]))


    df.loc[df["科目级次"] == 1, "parent_parent_code"] = ""
    df.loc[df["科目级次"] == 1, "parent_code"] = ""

    df.fillna("", inplace=True)

    # print("debug1")
    # print(df[df[code].str.contains("2221")].to_markdown())
    # print(df[df[code].str.contains("6601")].to_markdown())

    parent_df = df[[code, name]].copy()
    parent_df = parent_df.drop_duplicates(subset=[code, name], keep='first')
    parent_df = parent_df[parent_df[code].str.len() > 0]

    df = df.merge(parent_df, how="left", left_on="parent_code", right_on=code)[
        ["iid",  code+"_x", name+"_x",  "parent_code", name+"_y", "parent_code2","parent_code3","parent_code4"]]

    df.columns = ["iid",  code, name,  "parent_code", "parent_name", "parent_code2","parent_code3","parent_code4"]
    df.fillna("", inplace=True)

    # print("debug2")
    # print(df[df[code].str.contains("2221")].to_markdown())
    # print(df[df[code].str.contains("6601")] .to_markdown())

    # 3级科目
    df = df.merge(parent_df, how="left", left_on="parent_code2", right_on=code)[
        ["iid",  code+"_x", name+"_x",   "parent_code", "parent_name", "parent_code2",name+"_y", "parent_code3", "parent_code4"]]
    # 要用单引号，不要用双引号，否则rename会失败
    # df.rename(columns={code+'_x':code, code+'_y': code+'_parent_parent',name+'_x':name, name+'_y': name+'_parent_parent'}, inplace=True)
    df.columns = ["iid",  code, name,     "parent_code",  "parent_name", "parent_code2",  "parent_name2", "parent_code3", "parent_code4"]
    df.fillna("", inplace=True)

    # 4级科目
    df = df.merge(parent_df, how="left", left_on="parent_code3", right_on=code)[
        ["iid", code + "_x", name + "_x", "parent_code",  "parent_name","parent_code2", "parent_name2", "parent_code3", name + "_y", "parent_code4"]]
    df.columns = ["iid", code, name, "parent_code",  "parent_name", "parent_code2",   "parent_name2", "parent_code3", "parent_name3", "parent_code4"]
    df.fillna("", inplace=True)

    # 5级科目
    df = df.merge(parent_df, how="left", left_on="parent_code4", right_on=code)[
        ["iid", code + "_x", name + "_x", "parent_code", "parent_name", "parent_code2", "parent_name2", "parent_code3", "parent_name3", "parent_code4",
         name + "_y"]]
    df.columns = ["iid", code, name, "parent_code", "parent_name", "parent_code2", "parent_name2", "parent_code3",
                  "parent_name3", "parent_code4", "parent_name4"]
    df.fillna("", inplace=True)

    # print("debug3")
    # print(df[df[code].str.contains("2221")].to_markdown())
    # print(df[df[code].str.contains("6601")].to_markdown())

    # print("debug")
    # print(df[df["对应科目代码"].str.contains("2221")][["对应科目代码","对应科目名称","parent_code","对应科目名称_parent","parent_parent_code","对应科目名称_parent_parent"]].to_markdown())

    df[name+"_full"] = ""
    df.loc[df["parent_name"].str.replace("nan", "").str.len() == 0, name+"_full"] = df[name]
    df.loc[df["parent_name"].str.replace("nan", "").str.len() > 0, name+"_full"] = df["parent_name"] + sept_char + df[
        name]

    df.loc[df["parent_name2"].str.replace("nan", "").str.len() > 0, name+"_full"] = df["parent_name2"] + sept_char +df["parent_name"] + sept_char + df[name]
    df.loc[df["parent_name3"].str.replace("nan", "").str.len() > 0, name + "_full"] = df["parent_name3"] + sept_char +df["parent_name2"] + sept_char +df["parent_name"] + sept_char +  df[name]
    df.loc[df["parent_name4"].str.replace("nan", "").str.len() > 0, name + "_full"] = df["parent_name4"] + sept_char +df["parent_name3"] + sept_char +df["parent_name2"] + sept_char +df["parent_name"] + sept_char +  df[name]


    # print("debug4")
    # print(df[df[code].str.contains("2221")].to_markdown())
    # print(df[df[code].str.contains("6601")].to_markdown())

    # del df[name]
    del df["parent_code"]
    del df["parent_name"]
    del df["parent_code2"]
    del df["parent_name2"]
    del df["parent_code3"]
    del df["parent_name3"]
    del df["parent_code4"]
    del df["parent_name4"]

    df = df.drop_duplicates(subset=[code,name,name+"_full"], keep='first')

    # print("debug5")
    # print(df[df[code].str.contains("2221")].to_markdown())
    # print(df[df[code].str.contains("6601")].to_markdown())

    # print(df[df["科目代码"].str.contains("2202.1143")].head(20).to_markdown())

    # df.rename(columns={name+"_full": "对应科目名称"}, inplace=True)
    return df[[code,name,name+"_full"]]

def kemu_clear(kemu):
    # 去掉换行，回车

    # if kemu.find("6601")>=0:
    #     print("1-:",kemu)

    kemu= "aa_" + kemu.replace("\r","").replace("\n","").replace("_x000D_", "").strip() + "_bb"
    # 避免自动在科目后面追加 .0
    # if kemu.find("6601")>=0:
    #     print("2-:",kemu)

    kemu=kemu.replace(".0_", "")
    kemu=kemu.replace(".1_", ".10_")
    kemu=kemu.replace(".2_", ".20_")

    # kemu=kemu.replace(".3_", ".30_").replace("aa_", "").replace("_bb", "").replace("bb", "")
    kemu=kemu.replace(".3_", ".30_")
    kemu=kemu.replace(".4_", ".40_")
    kemu=kemu.replace(".5_", ".50_")
    kemu=kemu.replace(".6_", ".60_")
    kemu=kemu.replace(".7_", ".70_")
    kemu=kemu.replace(".8_", ".80_")
    kemu=kemu.replace(".9_", ".90_")
    kemu=kemu.replace("aa_", "").replace("_bb", "").replace("bb", "")
    return kemu

def read_cloudvouch(cloud_vouch_file):
    try:
        cloud_vouch = pd.read_excel(cloud_vouch_file)
    except Exception as  e:
        print("Excel文件读取出错:", e)
        sys.exit()
    # 删除第一行
    cloud_vouch["vouch_iid"] = cloud_vouch.index
    cloud_vouch["会计年度"].fillna(method="ffill", inplace=True)
    cloud_vouch["期间"].fillna(method="ffill", inplace=True)
    cloud_vouch["凭证号"].fillna(method="ffill", inplace=True)
    # cloud_vouch["FAccountBookID"].fillna(method="ffill", inplace=True)

    cloud_vouch["科目编码"] = cloud_vouch["科目编码"].astype(str)

    print("科目编码抽查1:")
    print(cloud_vouch[cloud_vouch["科目编码"].str.contains("6601.3")].to_markdown())


    cloud_vouch["科目编码"] = cloud_vouch["科目编码"].apply(lambda x: kemu_clear(x))

    cloud_vouch["vouchrank"] = cloud_vouch.apply(
        lambda x: "{}{:0>2d}{:0>6d}".format(int(x["会计年度"]), int(x["期间"]), int(x["凭证号"])), axis=1)
    # 分组内部建立序号 1,2,3
    cloud_vouch["iseq"] = cloud_vouch["vouch_iid"].groupby(cloud_vouch["vouchrank"]).rank(method='first',
                                                                                          ascending=True)

    # df["科目编码"] = df["科目编码"].astype(str)
    # df["科目编码"] = df["科目编码"].apply(lambda x: kemu_clear(x))

    print("科目编码抽查2:")
    print(cloud_vouch[cloud_vouch["科目编码"].str.contains("6601.3")].to_markdown())


    return cloud_vouch


def read_accountcode(filename,FAccountBookName):
    df=pd.read_excel(filename,sheet_name="对照表",skiprows=1, dtype={'对应科目代码':str,'科目代码':str,'核算项目代码':str,'对应核算项目编码':str})  # ,sheet_name="星空KIS科目对照表模板" dtype=str,
    # 科目代码	科目名称	核算项目代码	核算项目名称		对应科目代码	对应科目名称	项目辅助核算	对应核算项目编码	对应核算项目名称
    # *单据体(序号)	(单据体)摘要	*(单据体)科目编码#编码	(单据体)科目编码#名称	(单据体)客户分组#编码	(单据体)客户分组#名称(Null)	(单据体)仓库#编码	(单据体)仓库#名称(Null)	(单据体)组织机构#编码	(单据体)组织机构#名称(Null)	(单据体)物料分组#编码	(单据体)物料分组#名称(Null)	(单据体)存货类别#编码	(单据体)存货类别#名称(Null)	(单据体)银行账号#编码	(单据体)银行账号#名称(Null)	(单据体)其他往来单位#编码	(单据体)其他往来单位#名称(Null)	(单据体)借款单位#编码	(单据体)借款单位#名称(Null)	(单据体)部门#编码	(单据体)部门#名称(Null)	(单据体)客户#编码	(单据体)客户#名称(Null)	(单据体)供应商#编码	(单据体)供应商#名称(Null)	(单据体)费用项目#编码	(单据体)费用项目#名称(Null)	(单据体)资产类别#编码	(单据体)资产类别#名称(Null)	(单据体)员工#编码	(单据体)员工#名称(Null)	(单据体)物料#编码	(单据体)物料#名称(Null)

    # 筛选所属公司
    df=df[df["kis科目使用公司"].str.contains(FAccountBookName,na=False)]

    print("筛选的科目记录数：", df.shape[0])
    if df.shape[0] == 0:
        print("科目对照表没有找到记录，请检查公司名称是否输入正确？")
        sys.exit()


    df.fillna("", inplace=True)
    df["iid"] = df.index

    # 6601.3 -> 6601.30 如果最后一个级别的科目长度等于1，表示末尾的0被去掉了，现在补回来
    # df["对应科目代码"]=df["对应科目代码"].apply(lambda x:  x+"0" if len(x)>4 & len("".join(x.split(".")[-1:]))==1 else x )
    # df["科目代码"]=df["科目代码"].apply(lambda x:  x+"0" if  len(x)>4 & len("".join(x.split(".")[-1:]))==1 else x )
    # df["对应科目代码xx"] = df["对应科目代码"].apply(lambda x:  "".join(x.split(".")[-1:]) )

    df["科目代码"]=df["科目代码"].astype(str)
    df["科目名称"] = df["科目名称"].astype(str)
    # print(df[df.科目代码.str.contains("1012.01")].to_markdown())
    # print(df.to_markdown())

    # 科目修正
    df["科目代码"] =df["科目代码"].apply(lambda x: kemu_clear(x) )
    df["科目名称"] =df["科目名称"].apply(lambda x: kemu_clear(x) )

    # print("kemu1")
    # print(df[df["科目代码"].str.contains("2202")][["科目代码","对应科目代码", "对应科目名称"]].to_markdown())
    # print(df[df["科目代码"].str.contains("1122")][["科目代码","对应科目代码", "对应科目名称"]].to_markdown())
    # print(df.to_markdown())

    # print("debug1115")
    # print(df[df["科目代码"].str.contains("2221")].to_markdown())
    # print(df[df["科目代码"].str.contains("6601")].to_markdown())

    # print(df.head(3).to_markdown())
    # df=df.dropna(axis=1,how="all")
    # print(df.head(4).to_markdown())

    df["对应科目代码"]=df["对应科目代码"].astype(str)
    df["对应科目代码"] = df["对应科目代码"].apply(lambda x: kemu_clear(x))
    df["科目级次"] = df["对应科目代码"].apply(lambda x:x.count(".")+1)

    # print("kemu3")
    # print(df[df["科目代码"].str.contains("2202")][["科目代码", "对应科目代码", "对应科目名称","科目级次"]].to_markdown())
    # print(df[df["科目代码"].str.contains("1122")][["科目代码", "对应科目代码", "对应科目名称","科目级次"]].to_markdown())

    df_fullname=expand_codename(df, "科目代码" , "科目名称", "-")
    df = df.merge(df_fullname[["科目代码","科目名称_full"]], how="left", on="科目代码" )[
        ["iid", "科目代码", "科目名称_full","核算项目代码", "核算项目名称", "对应科目代码", "对应科目名称", "项目辅助核算", "对应核算项目编码", "对应核算项目名称" ]]

    df.rename(columns={"科目名称_full":"科目名称"})
    df.columns=["iid", "科目代码", "科目名称","核算项目代码", "核算项目名称", "对应科目代码", "对应科目名称", "项目辅助核算", "对应核算项目编码", "对应核算项目名称" ]

    df_fullname_duiying = expand_codename(df, "对应科目代码", "对应科目名称", "_")
    df = df.merge(df_fullname_duiying[["对应科目代码", "对应科目名称_full"]], how="left", on="对应科目代码")[
        ["iid", "科目代码", "科目名称", "核算项目代码", "核算项目名称", "对应科目代码", "对应科目名称_full", "项目辅助核算", "对应核算项目编码", "对应核算项目名称"]]

    df.rename(columns={"对应科目名称_full": "对应科目名称"})
    df.columns = ["iid", "科目代码", "科目名称", "核算项目代码", "核算项目名称", "对应科目代码", "对应科目名称", "项目辅助核算", "对应核算项目编码", "对应核算项目名称"]

    # 剔除不适用的科目
    # print("剔除不适用的科目")
    df["核算项目代码"] = df["核算项目代码"].astype(str)
    df = df[~df.核算项目代码.str.contains("不适用")]
    # print(df[df.科目代码.str.contains("1012.01")].to_markdown())

    #仓库
    df["warehouse_code"]=""
    df["warehouse_name"]=""
    # 存货类别
    df["productclass_code"] = ""
    df["productclass_name"] = ""
    # 客户分组
    df["custgroup_code"] = ""
    df["custgroup_name"] = ""
    # 部门
    df["dept_code"] = ""
    df["dept_name"] = ""

    # df.loc[df["项目辅助核算"].isin(["仓库"]), "warehouse_code"] = df["对应核算项目编码"]
    # df.loc[df["项目辅助核算"].isin(["仓库"]), "warehouse_name"] = df["对应核算项目名称"]

    # "仓库/物料分组","费用项目","费用项目/仓库/物料","供应商","客户","银行/借款单位","银行账号","员工","资产类别","组织机构","组织机构/其他往来单位"

    df["productgroup_code"] = ""
    df["productgroup_name"] = ""
    df.loc[df["项目辅助核算"].str.contains("仓库/物料分组"),   "productgroup_code"] = df["对应核算项目编码"]
    df.loc[df["项目辅助核算"].str.contains("仓库/物料分组"),   "productgroup_name"] = df["对应核算项目名称"]

    # print(df.head(9).to_markdown())

    df["fee_code"] = ""
    df["fee_name"] = ""
    df.loc[df["项目辅助核算"].isin(["费用项目"]), "fee_code"] = df["对应核算项目编码"]
    df.loc[df["项目辅助核算"].isin(["费用项目"]), "fee_name"] = df["对应核算项目名称"]

    df["product_code"] = ""
    df["product_name"] = ""
    df.loc[df["项目辅助核算"].isin(["费用项目/仓库/物料"]), "product_code"] = df["对应核算项目编码"]
    df.loc[df["项目辅助核算"].isin(["费用项目/仓库/物料"]), "product_name"] = df["对应核算项目名称"]

    df["vendor_code"] = ""
    df["vendor_name"] = ""
    df.loc[df["项目辅助核算"].str.contains("供应商"), "vendor_code"] = df["对应核算项目编码"]
    df.loc[df["项目辅助核算"].str.contains("供应商"), "vendor_name"] = df["对应核算项目名称"]

    df["cust_code"] = ""
    df["cust_name"] = ""
    df.loc[df["项目辅助核算"].str.contains("客户"), "cust_code"] = df["对应核算项目编码"]
    df.loc[df["项目辅助核算"].str.contains("客户"), "cust_name"] = df["对应核算项目名称"]

    df["ar_code"] = ""
    df["ar_name"] = ""
    df.loc[df["项目辅助核算"].str.contains("借款单位"), "ar_code"] = df["对应核算项目编码"]
    df.loc[df["项目辅助核算"].str.contains("借款单位"), "ar_name"] = df["对应核算项目名称"]

    df["bank_code"] = ""
    df["bank_name"] = ""
    df.loc[df["项目辅助核算"].str.contains("银行账号"), "bank_code"] = df["对应核算项目编码"]
    df.loc[df["项目辅助核算"].str.contains("银行账号"), "bank_name"] = df["对应核算项目名称"]

    df["employee_code"] = ""
    df["employee_name"] = ""
    df.loc[df["项目辅助核算"].str.contains("员工"), "employee_code"] = df["对应核算项目编码"]
    df.loc[df["项目辅助核算"].str.contains("员工"), "employee_name"] = df["对应核算项目名称"]

    df["asset_code"] = ""
    df["asset_name"] = ""
    df.loc[df["项目辅助核算"].str.contains("资产类别"), "asset_code"] = df["对应核算项目编码"]
    df.loc[df["项目辅助核算"].str.contains("资产类别"), "asset_name"] = df["对应核算项目名称"]

    df["unit_code"] = ""
    df["unit_name"] = ""
    df.loc[df["项目辅助核算"].isin(["组织机构"]), "unit_code"] = df["对应核算项目编码"]
    df.loc[df["项目辅助核算"].isin(["组织机构"]), "unit_name"] = df["对应核算项目名称"]

    df["others_code"] = ""
    df["others_name"] = ""
    df.loc[df["项目辅助核算"].str.contains("其他往来单位"), "others_code"] = df["对应核算项目编码"]
    df.loc[df["项目辅助核算"].str.contains("其他往来单位"), "others_name"] = df["对应核算项目名称"]

    # print(df[df["对应科目代码"].str.contains("4103")].to_markdown())
    # print(df[df["对应科目代码"].str.contains("6601")].to_markdown())

    return df

def translate(vouch,subject):
    # kis凭证根据科目对照关系表，转换生成星空云会计凭证

    print("凭证{}张，科目{}个：".format(vouch.shape[0],subject.shape[0]))

    subject.fillna("", inplace=True)

    vouch_count1=vouch.shape[0]

    vouch["vouch_iid"] = vouch.index

    vouch["subject_name"]=vouch["科目名称"].apply(lambda x:x.replace("-","-").replace(" ","").replace("\r","").replace("\n","").strip())
    vouch["subject_name"] = vouch["subject_name"].apply(lambda x: 'xxx_' + x + '_yyy')
    vouch["科目代码"] = vouch["科目代码"].apply(lambda x: x.strip())
    vouch["科目代码"]=vouch["科目代码"].astype(str)

    vouch["科目代码"] = vouch["科目代码"].apply(lambda x: kemu_clear(x))
    vouch["kemu"] = vouch["科目代码"].apply(lambda x: "aa_" + x + "_bb")

    # print("检查6601.02")
    # print(vouch[vouch.科目代码.str.contains("6601.02")].to_markdown())
    #
    # vouch.to_excel("work/test1.xlsx")

    # subject["subject_name"]=subject["科目名称"]
    subject["核算项目代码"]=subject["核算项目代码"].astype(str)
    # 如果有辅助核算项目
    # subject.loc[subject.核算项目代码.str.len()>0, "subject_name"]=subject.apply(lambda x: x["科目名称"]+"/"+"["+x["核算项目代码"]+"]"+x["核算项目名称"],axis=1)

    # 拼接科目名称
    subject["subject_name"]=subject.apply(lambda x: x["科目名称"]+"/"+"["+x["核算项目代码"]+"]"+x["核算项目名称"] if len(str(x["核算项目代码"]).replace("nan",""))>0 else x["科目名称"] ,axis=1)
    subject["subject_name"] = subject["subject_name"].apply(lambda x: x.replace("-", "-").replace(" ", "").replace("\r","").replace("\n","").strip())
    subject["subject_name"] = subject["subject_name"].apply(lambda x: 'xxx_'+x+'_yyy')
    subject["科目代码"] = subject["科目代码"].apply(lambda x: x.strip())
    subject["kemu"] = subject["科目代码"].apply(lambda x: "aa_" + x + "_bb")

    # print("数据检查1：")
    # print(vouch[vouch.摘要.str.contains("1/25付广东舒畅日用品有限公司")].to_markdown())
    # print(vouch.to_markdown())
    # print(subject.to_markdown())

    # print(vouch[vouch["科目代码"].str.contains("2202.1143")].head(20).to_markdown())
    # print(vouch[vouch["科目代码"].str.contains("1122.612")].head(20).to_markdown())

    # print("数据检查2：")
    # print(subject[subject["科目代码"].str.contains("2202.1143")].head(20).to_markdown())
    # print(subject[subject["科目代码"].str.contains("1122.612")].head(20).to_markdown())

    # sys.exit()

    # df = df[["vouch_iid","vouchno","iseq","日期", "会计期间", "凭证字号", "摘要", "科目代码", "科目名称", "币别", "汇率", "原币金额", "借方", "贷方"]]
    vouch2=vouch.merge(subject,how="left",on=["科目代码","subject_name"])

    # print("数据检查3：")
    # print(vouch2[vouch2["科目代码"].str.contains("2202.1143")].head(20).to_markdown())
    # print(vouch2[vouch2["科目代码"].str.contains("1122.612")].head(20).to_markdown())

    vouch2["对应科目名称"].fillna("",inplace=True)

    # print("合并后的结果：")
    # print(vouch2.to_markdown())

    blank_subject=vouch2[vouch2["对应科目名称"].str.len()==0]
    if blank_subject.shape[0]>0:
        print("数据异常，没有找到对应科目名称的凭证有{}张：".format(blank_subject.shape[0]))

        print("异常凭证：")
        print(blank_subject[["日期","凭证字号","科目代码","科目名称_x","借方","贷方","subject_name","kemu_x","kemu_y"]].to_markdown())

        exception_subject=pd.DataFrame(blank_subject.groupby(["科目代码","科目名称_x","subject_name","kemu_x"])["凭证字号"].count()).reset_index()
        exception_subject.columns=["科目代码","科目名称","subject_name","kemu","count"]
        #,"核算项目代码","核算项目名称","subject_name","kemu"
        exception_subject=exception_subject[["科目代码","科目名称"]].merge(subject[["科目代码","科目名称"]],how="left",on="科目代码")
        exception_subject.fillna("",inplace=True)
        exception_subject.drop_duplicates(subset=["科目代码","科目名称_x","科目名称_y"],keep='first')

        print("\r\n凭证中没有匹配到的会计科目（科目+辅助核算不匹配）：")
        print(exception_subject.to_markdown())
        # print("确定是由于辅助核算原因不匹配的会计科目：")
        # print(subject[subject["科目代码"].isin(blank_subject["科目代码"])][["科目代码","科目名称","核算项目代码","核算项目名称","subject_name","kemu"]].to_markdown())
        # print(subject[subject["科目代码"].str.contains("1012.01")][["科目代码","科目名称","核算项目代码","核算项目名称","subject_name","kemu"]].to_markdown())

        # print("匹配异常，按任意键退出！")
        # input("")
        # sys.exit()

    vouch_count2 = vouch2.shape[0]

    # vouch2.rename(columns={"对应科目名称":"FACCOUNTID#Name"},inplace=True)
    # vouch2.rename(columns={"对应科目代码":"FACCOUNTID"},inplace=True)

    # 科目名称只取最后一个级次就可以了
    vouch2["对应科目末级名称"]=vouch2["对应科目名称"].apply(lambda x:  "".join(x.split("_")[-1]) if  len(x.split("_"))>0 else x )

    # vouch2=vouch2.drop([""],axis=1)
    # vouch2.drop(columns=['审核','复核','过账','日期','会计期间','凭证字号','摘要','科目代码','科目名称_x','币别','汇率','原币金额','借方','贷方','制单','审核人','出纳人','核准','过账人','经办','批注','结算方式','结算号','数量',
    #                             '单价','参考信息','业务日期','往来业务','附件数','序号','系统模块','业务描述','vouch_iid','vouchno'],inplace=True)

    # 掏空重复数据
    # vouch2.loc[vouch2.iseq>1,"FBillHead(GL_VOUCHER)"]=""

    # print("test1")
    # print(vouch2.head(10).to_markdown())

    vouch2["核算维度"]=""

    vouch2.loc[vouch2["bank_code"].str.len()>0,"核算维度"]=vouch2.apply(lambda x: "{}/{}".format(x["bank_code"],x["bank_name"]) ,axis=1)
    vouch2.loc[vouch2["cust_code"].str.len()>0,"核算维度"]=vouch2.apply(lambda x: "{}/{}".format(x["cust_code"],x["cust_name"]) ,axis=1)
    vouch2.loc[vouch2["vendor_code"].str.len()>0,"核算维度"]=vouch2.apply(lambda x: "{}/{}".format(x["vendor_code"],x["vendor_name"]) ,axis=1)
    vouch2.loc[vouch2["unit_code"].str.len()>0,"核算维度"]=vouch2.apply(lambda x: "{}/{}".format(x["unit_code"],x["unit_name"]) ,axis=1)
    vouch2.loc[vouch2["employee_code"].str.len()>0,"核算维度"]=vouch2.apply(lambda x: "{}/{}".format(x["employee_code"],x["employee_name"]) ,axis=1)
    vouch2.loc[vouch2["others_code"].str.len()>0,"核算维度"]=vouch2.apply(lambda x: "{}/{}".format(x["others_code"],x["others_name"]) ,axis=1)
    vouch2.loc[vouch2["asset_code"].str.len()>0,"核算维度"]=vouch2.apply(lambda x: "{}/{}".format(x["asset_code"],x["asset_name"]) ,axis=1)

    vouch2.rename(columns={"对应科目名称":"对应科目全称"},inplace=True)
    vouch2.rename(columns={"对应科目末级名称":"对应科目名称","科目名称_x":"科目名称"},inplace=True)

    #
    # return  df_ori[["审核","复核","过账","日期","会计期间","凭证字号","摘要","科目代码","科目名称","币别","汇率","原币金额","借方","贷方"]],\
    #         df[["日期","会计年度","期间","凭证字号","凭证字","凭证号","摘要","科目代码","科目名称","币别","原币金额","借方","贷方","借方金额","贷方金额" ,"vouchrank","iseq"]]


    vouch3=vouch2[["vouch_iid","审核","复核","过账","日期","会计年度","会计期间","期间","凭证字号","凭证字","凭证号","摘要","科目代码","科目名称","对应科目代码","对应科目名称","对应科目全称","币别","汇率","原币金额","借方","贷方","借方金额","贷方金额","核算维度" ,"vouchrank","iseq"]]


    # print("test2")
    # print(vouch3.head(10).to_markdown())

    # title={"":{}}
    # vouch4=pd.DataFrame(columns=)

    temp_df=vouch3.vouchrank.value_counts()
    temp_df=pd.DataFrame(temp_df)

    # temp_df.columns=["FEntity_count"]
    # temp_df=temp_df.sort_values(["FEntity_count"],ascending=False)

    if blank_subject.shape[0]==0:
        if abs(vouch_count2-vouch_count1)==0:
            print("检查合格！")
        else:
            print("检查异常：{}".format(abs(vouch_count2)))
            print(temp_df.to_markdown())

    # 删除原来的索引
    # vouch3 = vouch3.reset_index() # drop=True
    # vouch3.id=vouch3.index

    # 表头拼接
    # vouch3=combine_multi_title(vouch3).copy()

    # # 请求选择一个用以保存的文件
    # save_file = filedialog.asksaveasfilename(initialdir=os.getcwd(),
    #                                       title="第三步：请输入要保存的文件:",
    #                                       filetypes=my_filetypes)
    #
    # # save_file="D:\数据处理\财务部\kis转k3cloud\ 盈养泉_测试2.xlsx"
    # # save_file=r"/Users/mac/Desktop/营养泉星空云2.xlsx"
    #
    # if save_file.find("xls")<0:
    #     save_file=save_file+".xlsx"
    #
    # vouch3.to_excel(save_file,index=False)
    # vouch3.to_excel(r"work/k3cloud.xlsx",sheet_name="凭证#单据头(FBillHead)")
    return vouch3



def sel_file(title_message):
    # 请求选择文件
    filename = filedialog.askopenfilename(initialdir=os.getcwd(),
                                          title=title_message,
                                          filetypes=my_filetypes)

    if len(filename) == 0:
        sys.exit()

    print("你选择的文件名是：", filename)
    return filename


if __name__ == "__main__":
    print("开始:", time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()))
    print("版本号：v2.0")

    print("请输入公司全称:")
    FAccountBookName = input("")
    # input_value ="025、盈养泉（深圳）化妆品有限公司, 124、盈养泉（深圳）化妆品有限公司"

    if len(FAccountBookName) == 0:
        print("你什么也没有输入，按任意键退出！")
        input("")
        sys.exit()

    # 请求选择文件
    filename_vouch = sel_file("第一步：请选择kis凭证序时账簿文件(vouch):")
    # filename_vouch = "D:\数据处理\财务部\kis转k3cloud\源数据_盈养泉 凭证序时簿 201912~202107(2).xls"
    # filename_vouch = "/Users/mac/Library/Containers/com.tencent.xinWeChat/Data/Library/Application Support/com.tencent.xinWeChat/2.0b4.0.9/30b19085e36d0f4274a8809b5f07ae0a/Message/MessageTemp/52fab857a1264ed949740f45e6071f64/File/源数据_盈养泉 凭证序时簿 201912~202107(2).xls"

    filename_map = sel_file("第二步：请选择科目对照表(map):")
    # filename= "D:\数据处理\财务部\kis转k3cloud\源数据_盈养泉_实际源数据对照表(1).xlsx"
    # filename_map= "/Users/mac/Library/Containers/com.tencent.xinWeChat/Data/Library/Application Support/com.tencent.xinWeChat/2.0b4.0.9/30b19085e36d0f4274a8809b5f07ae0a/Message/MessageTemp/52fab857a1264ed949740f45e6071f64/File/源数据_实际源数据对照表(1).xlsx"

    cloud_vouch_file = sel_file("第三步：请选择云星空导出的凭证文件:")

    vouch2 = read_kis(filename_vouch)
    # print(vouch1.head(10).to_markdown())
    # print(vouch2.head(10).to_markdown())

    cloud_vouch=read_cloudvouch(cloud_vouch_file)

    subject = read_accountcode(filename_map,FAccountBookName)

    if subject.shape[0]==0:
        print("科目对照表没有找到记录，请检查公司名称是否输入正确？")
        sys.exit()

    print("抽查科目对照表：")
    print(subject.head(10).to_markdown())

    kis_vouch=translate(vouch2,subject)


    kis_vouch.fillna("",inplace=True)
    cloud_vouch.fillna("",inplace=True)

    print("抽查kis导出凭证转换结果:")
    print(kis_vouch.head(20).to_markdown())
    print("抽查云星空导出凭证:")
    print(cloud_vouch.head(20).to_markdown())

    kd_compare=kis_vouch.merge(cloud_vouch,how="left",on=["vouchrank","iseq"])
    # print("比对的最后结果是：:")

    kd_compare.drop(columns=['审核_y','过账_y','日期_y'],inplace=True)
    kd_compare.rename(columns={"审核_x": "审核", "过账_x": "过账", "日期_x": "日期"}, inplace=True)

    kd_compare.rename(columns={"会计年度_x":"会计年度_1","会计年度_y":"会计年度_2"}, inplace=True)
    kd_compare.rename(columns={"期间_x":"期间_1","期间_y":"期间_2"}, inplace=True)
    kd_compare.rename(columns={"凭证号_x":"凭证号_1","凭证号_y":"凭证号_2"}, inplace=True)
    kd_compare.rename(columns={"摘要_x":"摘要_1","摘要_y":"摘要_2"}, inplace=True)
    kd_compare.rename(columns={"凭证字_x":"凭证字_1","凭证字_y":"凭证字_2"}, inplace=True)

    kd_compare.rename(columns={"核算维度_x":"核算维度_1","核算维度_y":"核算维度_2"}, inplace=True)
    kd_compare.rename(columns={"币别_x":"币别_1","币别_y":"币别_2"}, inplace=True)
    kd_compare.rename(columns={"原币金额_x":"原币金额_1","原币金额_y":"原币金额_2"}, inplace=True)
    kd_compare.rename(columns={"借方金额_x":"借方金额_1","借方金额_y":"借方金额_2"}, inplace=True)
    kd_compare.rename(columns={"贷方金额_x":"贷方金额_1","贷方金额_y":"贷方金额_2"}, inplace=True)

    kd_compare["凭证号_1"].fillna("0",inplace=True)
    kd_compare["凭证号_2"].fillna("0",inplace=True)

    kd_compare["摘要_1"]=kd_compare["摘要_1"].astype(str)
    kd_compare["摘要_2"]=kd_compare["摘要_2"].astype(str)

    # 替换换行符
    kd_compare["摘要_1"]= kd_compare["摘要_1"].apply(lambda x: x.replace("\r","").replace("\n","").replace(" ", "").replace("_x000D_","").strip())
    kd_compare["摘要_2"]= kd_compare["摘要_2"].apply(lambda x: x.replace("\r","").replace("\n","").replace(" ", "").replace("_x000D_","").strip())


    kd_compare["年度_"]=kd_compare.apply(lambda x:  "True" if    operator.eq( x["会计年度_1"] ,x["会计年度_2"]) else "False" ,axis=1 )
    kd_compare["月份_"]=kd_compare.apply(lambda x:  "True" if    operator.eq( x["期间_1"] ,x["期间_2"]) else "False" ,axis=1 )
    kd_compare["凭证号_"]=kd_compare.apply(lambda x:  "True" if   abs(int(x["凭证号_1"])-int(x["凭证号_2"]))<=0.01 else "False" ,axis=1 )
    kd_compare["摘要_"]=kd_compare.apply(lambda x:  "True" if    operator.eq( x["摘要_1"] ,x["摘要_2"]) else "False" ,axis=1 )
    kd_compare["科目编码_"]=kd_compare.apply(lambda x:  "True" if    operator.eq( x["对应科目代码"] ,x["科目编码"]) else "False" ,axis=1 )
    kd_compare["科目全名_"]=kd_compare.apply(lambda x:  "True" if    operator.eq( x["对应科目全称"] ,x["科目全名"]) else "False" ,axis=1 )
    kd_compare["核算维度_"]=kd_compare.apply(lambda x:  "True" if    operator.eq( x["核算维度_1"] ,x["核算维度_2"]) else "False" ,axis=1 )
    kd_compare["币别_"]=kd_compare.apply(lambda x:  "True" if    operator.eq( x["币别_1"] ,x["币别_2"]) else "False" ,axis=1 )
    kd_compare["原币金额_"]=kd_compare.apply(lambda x:  "True" if    operator.eq( x["原币金额_1"] ,x["原币金额_2"]) else "False" ,axis=1 )


    # print("kkkkkk")
    # print(kd_compare.head(50).to_markdown())

    kd_compare["借方金额_"]=kd_compare.apply(lambda x:  "True" if    operator.eq( x["借方金额_1"] ,x["借方金额_2"]) else "False" ,axis=1 )
    kd_compare["贷方金额_"]=kd_compare.apply(lambda x:  "True" if    operator.eq( x["贷方金额_1"] ,x["贷方金额_2"]) else "False" ,axis=1 )

    kd_compare.fillna("",inplace=True)

    kd_compare["科目编码"]=kd_compare["科目编码"].astype(str)
    kd_compare["科目编码"] = kd_compare["科目编码"].apply(lambda x: kemu_clear(x))

    kd_compare["年度_"] = kd_compare.apply(lambda x: x["年度_"] if len(x["科目编码"])>0 else "", axis=1)
    kd_compare["月份_"] = kd_compare.apply(lambda x: x["月份_"] if len(x["科目编码"])>0 else "", axis=1)
    kd_compare["凭证号_"] = kd_compare.apply(lambda x: x["凭证号_"] if len(x["科目编码"])>0 else "", axis=1)
    kd_compare["摘要_"] = kd_compare.apply(lambda x: x["摘要_"] if len(x["科目编码"])>0 else "", axis=1)
    kd_compare["科目编码_"] = kd_compare.apply(lambda x: x["科目编码_"] if len(x["科目编码"])>0 else "", axis=1)
    kd_compare["科目全名_"] = kd_compare.apply(lambda x: x["科目全名_"] if len(x["科目编码"])>0 else "", axis=1)
    kd_compare["核算维度_"] = kd_compare.apply(lambda x: x["核算维度_"] if len(x["科目编码"])>0 else "", axis=1)
    kd_compare["币别_"] = kd_compare.apply(lambda x: x["币别_"] if len(x["科目编码"])>0 else "", axis=1)
    kd_compare["原币金额_"] = kd_compare.apply(lambda x: x["原币金额_"] if len(x["科目编码"])>0 else "", axis=1)
    kd_compare["借方金额_"] = kd_compare.apply(lambda x: x["借方金额_"] if len(x["科目编码"]) > 0 else "", axis=1)
    kd_compare["贷方金额_"] = kd_compare.apply(lambda x: x["贷方金额_"] if len(x["科目编码"]) > 0 else "", axis=1)

    kd_compare.rename(columns={"对应科目代码": "对照表K对应星空科目代码", "对应科目名称": "对照表K对应星空科目名称"}, inplace=True)
    kd_compare.rename(columns={ "对应科目全称": "对照表K对应科目全称","核算维度_1":"对照表K对应星空维度"}, inplace=True)
    kd_compare.rename(columns={ "科目全名": "对照表星空科目全称","核算维度_2":"核算维度"}, inplace=True)

    kd_compare.loc[kd_compare.iseq.astype(int)>1,'日期']=""
    kd_compare.loc[kd_compare.iseq.astype(int)>1,'会计年度_1']=""
    kd_compare.loc[kd_compare.iseq.astype(int)>1,'会计期间']=""
    kd_compare.loc[kd_compare.iseq.astype(int)>1,'期间_1']=""
    kd_compare.loc[kd_compare.iseq.astype(int)>1,'凭证字号']=""
    kd_compare.loc[kd_compare.iseq.astype(int)>1,'凭证字_1']=""
    kd_compare.loc[kd_compare.iseq.astype(int)>1,'凭证号_1']=""
    kd_compare.loc[kd_compare.iseq.astype(int)>1,'会计年度_2']=""
    kd_compare.loc[kd_compare.iseq.astype(int)>1,'期间_2']=""
    kd_compare.loc[kd_compare.iseq.astype(int)>1,'凭证号_2']=""

    # print("testtest")
    # print(kd_compare.head(50).to_markdown())

    kd_compare2=kd_compare[["审核","复核","过账","日期","会计期间","会计年度_1","凭证字号","凭证字_1","凭证号_1","摘要_1","科目代码","科目名称","币别_1"
        ,"汇率","原币金额_1","借方","贷方","对照表K对应星空科目代码","对照表K对应星空科目名称","对照表K对应科目全称","对照表K对应星空维度"
                      ,"会计年度_2","期间_2","凭证字_2","凭证号_2","摘要_2","科目编码","对照表星空科目全称","核算维度","币别_2","原币金额_2",
                      "借方金额_2","贷方金额_2","年度_","月份_","凭证号_","摘要_","科目编码_","科目全名_","核算维度_","币别_","原币金额_","借方金额_","贷方金额_"]].copy()

    kd_compare2.columns=["审核","复核","过账","日期","会计期间","年度","凭证字号","凭证字","凭证号","摘要","科目代码","科目名称","币别"
        ,"汇率","原币金额","借方","贷方","对照表K对应星空科目代码","对照表K对应星空科目名称","对照表K对应科目全称","对照表K对应星空维度"
                      ,"会计年度","期间","凭证字","凭证号","摘要","科目编码","对照表星空科目全称","核算维度","币别","原币金额",
                      "借方金额","贷方金额","年度_","月份_","凭证号_","摘要_","科目编码_","科目全名_","核算维度_","币别_","原币金额_","借方金额_","贷方金额_"]

    # print(kd_compare2.head(50).to_markdown())

    # 请求选择一个用以保存的文件
    save_file = filedialog.asksaveasfilename(initialdir=os.getcwd(),
                                          title="第四步：请输入要保存的对比结果文件名:",
                                          filetypes=my_filetypes)

    # save_file="D:\数据处理\财务部\kis转k3cloud\ 盈养泉_测试2.xlsx"
    # save_file=r"/Users/mac/Desktop/营养泉星空云2.xlsx"

    if save_file.find("xls")<0:
        save_file=save_file+".xlsx"

    print("对比结果保存到：",save_file)
    # if  save_file.find("xls")<0:
    if len(save_file) < 6:
        sys.exit()

    kd_compare2.to_excel(save_file)

    # compare_vouch(kis_vouch,cloud_vouch)
    # print("文件生成完毕！")
    # input("")
    # pyinstaller -p D:\Anaconda3\envs\PyProject -F .\K3CloudCompare.py
    print("ok")