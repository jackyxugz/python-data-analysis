# __coding=utf8__
# /** 作者：zengyanghui **/
import pandas as pd
import numpy as np
import xlrd
import xlwt
import sys
import os
import os.path
import time
import tabulate

def cal_error_percent(fn,oms):
    # print(fn)
    # print(oms)
    # if (len((str(fn)))==0 & len((str(oms)))!=0):
    #     return "100%"
    # elif (len((str(fn)))!=0 & len((str(oms)))==0):
    #     return "100%"
    # elif (len((str(fn)))==0 & len((str(oms)))==0):
    #     return "0%"
    if (float(oms) != 0 and float(fn) == 0):
        return "100%"
    elif (float(oms) == 0 and float(fn) != 0):
        return "100%"
    elif (float(oms) == 0 and float(fn) == 0):
        return "0.00%"
    else:
        return   "{:.2f}".format(abs(float(fn)-float(oms))/max(fn,oms)*100.00)+"%"


def report_v2h(filename,iyear):
    # 平台	店铺名称	年度	月份	财务-订单数量	财务-订单金额	导出-订单数量	导出-订单金额	数量差异（财务/订单）	金额差异（财务/订单）
    df=pd.read_excel(filename,sheet_name="账单对比表",skiprows=1)
    print(df.head(30).to_markdown())
    # df=df[df["年"].isin([iyear])]
    df = df.apply(lambda x:x.astype(str).replace("nan","0"))
    df["年"] = df["年"].apply(lambda x:x.replace("年",""))
    df["年"] = df["年"].astype(int)
    df["月份"] = df["月份"].apply(lambda x:x.replace("月","").replace("（补2.3）","").replace("补",""))
    df["月份"] = df["月份"].astype(float)
    print(df.head(30).to_markdown())
    # df["财务-订单数量"].fillna(0, inplace=True)
    df["财务收入金额"].fillna(0, inplace=True)
    df["财务收入金额"] = df["财务收入金额"].astype(float)

    # df["导出-订单数量"].fillna(0,inplace=True)
    df["oms收入金额"].fillna(0,inplace=True)
    df["oms收入金额"] = df["oms收入金额"].astype(float)
        
    # df["财务-订单数量"]=df["财务-订单数量"].astype(int)
    # df["导出-订单数量"]=df["导出-订单数量"].astype(int)

    df2=df.copy()

    for i in range(1,13):
        # df2["订单行(财务/oms)-{}".format(i)]=df2.apply(lambda x:  "{:.2f}/{:.2f}".format(x["财务-订单数量"],x["导出-订单数量"]) if int(x["月份"])==i else "" ,axis=1)
        # df2["行差异(财务/订单)-{}".format(i)]=df2.apply(lambda x: cal_error_percent(x["财务-订单数量"],x["导出-订单数量"]) if int(x["月份"])==i else "" ,axis=1)

        df2["账单收入金额(财务/oms)-{}".format(i)]=df2.apply(lambda x:  "{:.2f}/{:.2f}".format(x["财务收入金额"],x["oms收入金额"]) if int(x["月份"])==i else "" ,axis=1)
        df2["金额差异(财务/oms)-{}".format(i)]=df2.apply(lambda x: cal_error_percent(x["财务收入金额"],x["oms收入金额"]) if int(x["月份"])==i else "" ,axis=1)


    del df2["主体"]
    del df2["Unnamed: 5"]
    del df2["oms回款金额"]
    del df2["oms退款金额"]
    del df2["oms回款金额.1"]
    del df2["oms退款金额.1"]
    # del df2["oms收入金额"]
    del df2["oms收入金额（不含税）"]
    del df2["财务回款金额"]
    del df2["财务退款金额"]
    del df2["财务回款金额.1"]
    del df2["财务退款金额.1"]
    # del df2["财务收入金额"]
    del df2["财务收入金额（不含税）"]
    del df2["差异回款金额"]
    del df2["差异退款金额"]
    del df2["差异回款金额.1"]
    del df2["差异退款金额.1"]
    del df2["差异收入金额（不含税）"]
    del df2["原因"]

    df3=df2.groupby(["平台","店铺名称","年"]).agg({
        # "订单行(财务/oms)-1":np.max,"行差异(财务/订单)-1":np.max,
        "账单收入金额(财务/oms)-1":np.max,"金额差异(财务/oms)-1":np.max,
        # "订单行(财务/oms)-2":np.max,"行差异(财务/订单)-2":np.max,
        "账单收入金额(财务/oms)-2":np.max,"金额差异(财务/oms)-2":np.max,
        # "订单行(财务/oms)-3":np.max,"行差异(财务/订单)-3":np.max,
        "账单收入金额(财务/oms)-3":np.max,"金额差异(财务/oms)-3":np.max,
        # "订单行(财务/oms)-4":np.max,"行差异(财务/订单)-4":np.max,
        "账单收入金额(财务/oms)-4":np.max,"金额差异(财务/oms)-4":np.max,
        # "订单行(财务/oms)-5":np.max,"行差异(财务/订单)-5":np.max,
        "账单收入金额(财务/oms)-5":np.max,"金额差异(财务/oms)-5":np.max,
        # "订单行(财务/oms)-6":np.max,"行差异(财务/订单)-6":np.max,
        "账单收入金额(财务/oms)-6":np.max,"金额差异(财务/oms)-6":np.max,
        # "订单行(财务/oms)-7":np.max,"行差异(财务/订单)-7":np.max,
        "账单收入金额(财务/oms)-7":np.max,"金额差异(财务/oms)-7":np.max,
        # "订单行(财务/oms)-8":np.max,"行差异(财务/订单)-8":np.max,
        "账单收入金额(财务/oms)-8":np.max,"金额差异(财务/oms)-8":np.max,
        # "订单行(财务/oms)-9":np.max,"行差异(财务/订单)-9":np.max,
        "账单收入金额(财务/oms)-9":np.max,"金额差异(财务/oms)-9":np.max,
        # "订单行(财务/oms)-10":np.max,"行差异(财务/订单)-10":np.max,
        "账单收入金额(财务/oms)-10":np.max,"金额差异(财务/oms)-10":np.max,
        # "订单行(财务/oms)-11":np.max,"行差异(财务/订单)-11":np.max,
        "账单收入金额(财务/oms)-11":np.max,"金额差异(财务/oms)-11":np.max,
        # "订单行(财务/oms)-12":np.max,"行差异(财务/订单)-12":np.max,
        "账单收入金额(财务/oms)-12":np.max,"金额差异(财务/oms)-12":np.max}).reset_index()
    print(df3.to_markdown())
    df3.to_excel( r"/Users/maclove/Downloads/2019财务和oms账单收入金额差距_横向-11-11.xlsx")


if __name__ == "__main__":
    print("开始:", time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()))
    # print("abc".find("x"))

    # caiwu_xushizhang()

    # combine_excel()
    # groupby_amt()
    # math_file()

    # report_v2h( r"work/财务和导出订单的数量和金额差距(2).xlsx",2019)
    report_v2h(r"/Users/maclove/Downloads/账单对比表-2021_1120.xlsx", 2020)

    print("ok")