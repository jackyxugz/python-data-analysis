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
from sqlalchemy import create_engine
import pymysql
import itertools
import re

pymysql.install_as_MySQLdb()

OMS_PROD_DB_NAME = "megaorderbill"
OMS_PROD_DB_USER = "chenxiaoselect"
OMS_PROD_DB_PWD = "NTC1abr2tqa6bev-hmr"
OMS_PROD_DB_HOST = "megaoms.rwlb.rds.aliyuncs.com"
OMS_PROD_DB_PORT = "3306"

# 本地存储目录
LOCAL_FILE_PATH = r'D:\ITDD10'
# 远程文件目录
REMOTE_FILE_PATH = r'Z:\it审计处理需求\IT审计\处理后文件\OMS通用格式账单\2021'
platform_list = ["WPH", "BD", "XHS", "TAOBAO", "TMALL", "DY", "JD", "JN", "KAOLA", "KS", "PDD", "YZ", "ALIBABA", "WM",
                 "FY", "MD", "ZMB"]
bill_year = 2021
bill_month = range(1, 13)


def list_all_files(rootdir, filekey_list):
    # filekey_list="2020|2019"  or
    # filekey_list="2019(.*)海外旗舰"  and
    # filekey_list = "(?!京东)"

    if len(filekey_list) > 0:
        filekey_list = filekey_list.replace(",", " ").replace("，", " ")
        filekey = filekey_list.split(" ")
        pass
    else:
        filekey = ''

    _files = []
    list = os.listdir(rootdir)  # 列出文件夹下所有的目录与文件
    for i in range(0, len(list)):
        path = os.path.join(rootdir, list[i])
        if os.path.isdir(path):
            _files.extend(list_all_files(path, filekey_list))
        if os.path.isfile(path):
            if (path.find("~") < 0) and (path.find(".DS_Store") < 0) and (path.find("._") < 0):  # 带~符号表示临时文件，不读取
                # if len(filekey_list) > 0:
                #     t=re.search(filekey_list,path)
                #     if t:  # 如果匹配成功
                #         _files.append(path)

                if len(filekey) > 0:
                    break_flag = False
                    for key in filekey:
                        if not break_flag:
                            # print(path)

                            # 简化版的不包含(类似正则表达式)  !京东 = 不包含京东
                            if ((len(key.replace("!", "")) + 1 == len(key)) and (key.find("?") < 0) and (
                                    key.find(".") < 0) and (key.find("(") < 0) and (key.find(")") < 0)):
                                if path.find(key.replace("!", "")) >= 0:
                                    # 要求不包含，结果找到了！
                                    print("要求不包含{}，结果找到了！".format(key.replace("!", "")))
                                    break_flag = True
                            else:
                                t = re.search(key, path)
                                if t:  # 如果匹配成功
                                    pass
                                else:
                                    # 只要有一项匹配不成功，则自动退出，认为不符合条件
                                    break_flag = True

                    if not break_flag:
                        _files.append(path)

                else:
                    _files.append(path)

    # print(_files)
    return _files


def get_files_df(rootdir, filekey):
    filelist = list_all_files(rootdir, filekey)
    # print(filelist)
    if len(filelist) > 0:
        mySeries = pd.Series(filelist)
        df = pd.DataFrame(mySeries)
        df.columns = ["filename"]
        # df["filename"] = df["filename"].apply(lambda x: x.strip().replace(" ", ""))

        # print(df.to_markdown())
        print(df.to_markdown())
        return df
    else:
        print("没有发现符合条件的文件！")
        # 创建一个空表


def combine_platform_excel(rootdir, filekey):
    # title_key,bottom_key
    platform = ""
    iyear = 1999
    imonth = 0
    df_files = get_files_df(rootdir, filekey)
    df_box = []
    files_count = df_files.shape[0]
    for index, file in df_files.iterrows():
        # 根据文件名，匹配不同的政策
        k = 0
        df = read_bill_excel(file["filename"])
        # pd.read_excel(file["filename"])

        print("测试...")
        print(df.head(3))

        # if df.shape[0]==0:
        #     print(file["filename"]+" 格式错误！")
        #     pass
        #
        # if  "platform" not in df.columns:
        #     print(file["filename"]+" 格式错误！")
        #     pass

        if df.empty:
            print("pass 空表")
            pass
        else:

            # print(file["filename"] )
            # print( policy_list)
            # 从数据表中获取平台信息
            platform = df["platform"].iloc[0]
            # iyear=df["iyear"].iloc[0]
            # imonth=df["imonth"].iloc[0]

            # for pf in platform_list:
            if platform == "TAOBAO":
                if "df_TAOBAO" in vars():
                    df_TAOBAO = df_TAOBAO.append(df)
                else:
                    df_TAOBAO = df
            elif platform == "TMALL":
                if "df_TMALL" in vars():
                    df_TMALL = df_TMALL.append(df)
                else:
                    df_TMALL = df
            elif platform == "DY":
                if "df_DY" in vars():
                    df_DY = df_DY.append(df)
                else:
                    df_DY = df
            elif platform == "JD":
                if "df_JD" in vars():
                    df_JD = df_JD.append(df)
                else:
                    df_JD = df

            elif platform == "JN":
                if "df_JN" in vars():
                    df_JN = df_JN.append(df)
                else:
                    df_JN = df

            elif platform == "KAOLA":
                if "df_KAOLA" in vars():
                    df_KAOLA = df_KAOLA.append(df)
                else:
                    df_KAOLA = df
            elif platform == "KS":
                if "df_KS" in vars():
                    df_KS = df_KS.append(df)
                else:
                    df_KS = df

            elif platform == "PDD":
                if "df_PDD" in vars():
                    df_PDD = df_PDD.append(df)
                else:
                    df_PDD = df

            elif platform == "WPH":
                if "df_WPH" in vars():
                    df_WPH = df_WPH.append(df)
                else:
                    df_WPH = df

            elif platform == "XHS":
                if "df_XHS" in vars():
                    df_XHS = df_XHS.append(df)
                else:
                    df_XHS = df

            elif platform == "YZ":
                if "df_YZ" in vars():
                    df_YZ = df_YZ.append(df)
                else:
                    df_YZ = df

            elif platform == "ALIBABA":
                if "df_ALIBABA" in vars():
                    df_ALIBABA = df_ALIBABA.append(df)
                else:
                    df_ALIBABA = df

            elif platform == "WM":
                if "df_WM" in vars():
                    df_WM = df_WM.append(df)
                else:
                    df_WM = df

            elif platform == "FY":
                if "df_FY" in vars():
                    df_FY = df_FY.append(df)
                else:
                    df_FY = df

            elif platform == "MD":
                if "df_MD" in vars():
                    df_MD = df_MD.append(df)
                else:
                    df_MD = df

            elif platform == "BD":
                if "df_BD" in vars():
                    df_BD = df_BD.append(df)
                else:
                    df_BD = df
            elif platform == "ZMB":
                if "df_ZMB" in vars():
                    df_ZMB = df_ZMB.append(df)
                else:
                    df_ZMB = df

            print("进度表：  {}/{}  ".format(index + 1, files_count))

    # 保存的文件名
    # save_filename=LOCAL_FILE_PATH+os.sep+"fn_EXCEL{}"
    if "df_TAOBAO" in vars():
        save_filename = LOCAL_FILE_PATH + os.sep + "fn_EXCEL_{}.pkl".format("TAOBAO")
        df_TAOBAO.to_pickle(save_filename)

    if "df_TMALL" in vars():
        save_filename = LOCAL_FILE_PATH + os.sep + "fn_EXCEL_{}.pkl".format("TMALL")
        df_TAOBAO.to_pickle(save_filename)

    if "df_DY" in vars():
        save_filename = LOCAL_FILE_PATH + os.sep + "fn_EXCEL_{}.pkl".format("DY")
        df_DY.to_pickle(save_filename)

    if "df_JD" in vars():
        save_filename = LOCAL_FILE_PATH + os.sep + "fn_EXCEL_{}.pkl".format("JD")
        df_JD.to_pickle(save_filename)

    if "df_JN" in vars():
        save_filename = LOCAL_FILE_PATH + os.sep + "fn_EXCEL_{}.pkl".format("JN")
        df_JN.to_pickle(save_filename)

    if "df_KAOLA" in vars():
        save_filename = LOCAL_FILE_PATH + os.sep + "fn_EXCEL_{}.pkl".format("KAOLA")
        df_KAOLA.to_pickle(save_filename)

    if "df_KS" in vars():
        save_filename = LOCAL_FILE_PATH + os.sep + "fn_EXCEL_{}.pkl".format("KS")
        df_KS.to_pickle(save_filename)

    if "df_PDD" in vars():
        save_filename = LOCAL_FILE_PATH + os.sep + "fn_EXCEL_{}.pkl".format("PDD")
        df_PDD.to_pickle(save_filename)

    if "df_WPH" in vars():
        save_filename = LOCAL_FILE_PATH + os.sep + "fn_EXCEL_{}.pkl".format("WPH")
        df_WPH.to_pickle(save_filename)

    if "df_XHS" in vars():
        save_filename = LOCAL_FILE_PATH + os.sep + "fn_EXCEL_{}.pkl".format("XHS")
        df_XHS.to_pickle(save_filename)

    if "df_YZ" in vars():
        save_filename = LOCAL_FILE_PATH + os.sep + "fn_EXCEL_{}.pkl".format("YZ")
        df_YZ.to_pickle(save_filename)

    if "df_ALIBABA" in vars():
        save_filename = LOCAL_FILE_PATH + os.sep + "fn_EXCEL_{}.pkl".format("ALIBABA")
        df_ALIBABA.to_pickle(save_filename)

    if "df_WM" in vars():
        save_filename = LOCAL_FILE_PATH + os.sep + "fn_EXCEL_{}.pkl".format("WM")
        df_WM.to_pickle(save_filename)

    if "df_FY" in vars():
        save_filename = LOCAL_FILE_PATH + os.sep + "fn_EXCEL_{}.pkl".format("FY")
        df_FY.to_pickle(save_filename)

    if "df_MD" in vars():
        save_filename = LOCAL_FILE_PATH + os.sep + "fn_EXCEL_{}.pkl".format("MD")
        df_MD.to_pickle(save_filename)

    if "df_BD" in vars():
        save_filename = LOCAL_FILE_PATH + os.sep + "fn_EXCEL_{}.pkl".format("BD")
        df_BD.to_pickle(save_filename)

    if "df_ZMB" in vars():
        save_filename = LOCAL_FILE_PATH + os.sep + "fn_EXCEL_{}.pkl".format("ZMB")
        df_ZMB.to_pickle(save_filename)

    print("合并并按平台分拆excel成功！")
    # return ""


def combine_excel(filedir, keyword):
    # df_box = self.get_table_box(filepath,keyword)
    print('请选择要合并的文件目录（请确定所有excel的表头都是相同的哦！）:')

    if len(filedir) > 0:
        pass
    else:
        filedir = filedialog.askdirectory()  # 获取文件夹
        print("你选择的路径是：", filedir)

    if len(filedir) == 0:
        print("你没有输入任何目录 :(")
        sys.exit()
        return

    global default_dir
    default_dir = filedir

    if len(keyword) > 0:
        pass
    else:
        # print('请输入要合并的文件名称或者后缀匹配符:（比如都必须包含 序时账 三个字，那么就请输入  "序时账" ， 不输入就表示所有excel都要合并！）')
        print(
            '筛选文件的规则:  \r\n1、京东 csv 表示选择文件完整路径中包含 "京东"和"csv"的文件  \r\n2、比如 淘宝 !天猫  表示只要淘宝，不要天猫  \r\n3、淘宝|天猫 表示 包含淘宝或者天猫    \r\n4、空格中间是and关系，每个项目都支持正则表达式 比如：2019(.*)海外旗舰  \r\n4、什么都不输入，表示默认选择目录下所有文件! \r\n请输入:')
        keyword = input()

        if len(filedir) == 0:
            print("你没有输入任何关键词 :(")
            keyword = ''
            # sys.exit()
            # return

    print("你希望在'{}'目录下找到所有的包含“{}”文件，然后合并。".format(filedir, keyword))

    combine_platform_excel(filedir, keyword)

    # print("生成完毕，现在关闭吗？yes/no")
    # byebye = input()
    # print('bybye:', byebye)


# 从数据库下载账单数据
def download_bill_records(p_engine, p_bill_year, p_bill_month):
    #  created,
    select_sql = r"""   select  {bill_year_show} iyear,{bill_month_show} imonth,platform,shopcode,  tid,                      
                            sum(case when IS_REFUNDAMOUNT = 1 then EXPEND_AMOUNT else 0 end) expend_amount,
                            sum(case when IS_AMOUNT = 1 then INCOME_AMOUNT else 0 end) income_amount
                        from `megaorder-{v_bill_year}`.order_info_bill_{v_bill_month}
                        where ((IS_REFUNDAMOUNT=1)  or (is_amount = 1))
                        group by platform, shopcode,  tid
                   ;""".format(
        bill_year_show=p_bill_year, bill_month_show=p_bill_month, v_bill_year=str(p_bill_year).zfill(4),
        v_bill_month=str(p_bill_month).zfill(2))
    try:
        read_begin_time = time.time()
        df = pd.read_sql_query(select_sql, p_engine)
        read_end_time = time.time()
        print(p_bill_year, '-', p_bill_month, '读取数据用时:', read_end_time - read_begin_time,
              '秒')

    except Exception as e:
        print('select_sql:\n', select_sql)
        print(e)
        raise 'error'

    return df


def read_bill_excel(filename):
    df = pd.read_excel(filename)
    print("抽查数据")
    print(df.head(3).to_markdown())
    # TID	SHOPNAME	PLATFORM	SHOPCODE	BILLPLATFORM	CREATED	TITLE	TRADE_TYPE	BUSINESS_NO	INCOME_AMOUNT	EXPEND_AMOUNT	TRADING_CHANNELS	BUSINESS_DESCRIPTION	remark	IS_REFUNDAMOUNT	IS_AMOUNT	OID	SOURCEDATA	RECIPROCAL_ACCOUNT	BATCHNO	currency	overseas_income	overseas_expend	currency_cny_rate	filename
    # df=df["tid","created","billplatform","shopcode","income_amount","expend_amount","is_amount","is_refundamount"]

    if ("TID" not in df.columns):
        print(filename + " TID 没有发现")
        return pd.DataFrame()

    if ("CREATED" not in df.columns):
        print(filename + " CREATED 没有发现")
        return pd.DataFrame()

    if ("BILLPLATFORM" not in df.columns):
        print("BILLPLATFORM 没有发现")
        # return pd.DataFrame()

    if ("SHOPCODE" not in df.columns):
        print(filename + " SHOPCODE 没有发现")
        return pd.DataFrame()

    if ("INCOME_AMOUNT" not in df.columns):
        print(filename + " INCOME_AMOUNT 没有发现")
        return pd.DataFrame()

    if ("EXPEND_AMOUNT" not in df.columns):
        print(filename + " EXPEND_AMOUNT 没有发现")
        return pd.DataFrame()

    if ("IS_AMOUNT" not in df.columns):
        print(filename + " IS_AMOUNT 没有发现")
        return pd.DataFrame()

    if ("IS_REFUNDAMOUNT" not in df.columns):
        print(filename + " IS_REFUNDAMOUNT 没有发现")
        return pd.DataFrame()

    if ("BILLPLATFORM" in df.columns):
        df = df[["TID", "CREATED", "BILLPLATFORM", "SHOPCODE", "INCOME_AMOUNT", "EXPEND_AMOUNT", "IS_AMOUNT",
                 "IS_REFUNDAMOUNT"]]
    else:
        df = df[["TID", "CREATED", "PLATFORM", "SHOPCODE", "INCOME_AMOUNT", "EXPEND_AMOUNT", "IS_AMOUNT",
                 "IS_REFUNDAMOUNT"]]

    df.columns = ["tid", "created", "platform", "shopcode", "income_amount", "expend_amount", "is_amount",
                  "is_refundamount"]
    df["is_amount"] = df["is_amount"].astype(str)
    df["is_refundamount"] = df["is_refundamount"].astype(str)
    # df=df[ df["is_amount"].str.contains("1") | df["is_refundamount"].str.contains("1") ]
    df = df[(df["is_amount"].str.contains('1')) | (df["is_refundamount"].str.contains('1'))]
    df["iyear"] = df["created"].apply(lambda x: x[:4])
    df["imonth"] = df["created"].apply(lambda x: x[5:7])
    print("抽查数据2")
    print(df.head(3).to_markdown())
    del df["created"]
    df = df.groupby(["iyear", "imonth", "tid", "platform", "shopcode"]).agg(
        {"income_amount": sum, "expend_amount": sum})  # "created",
    df = pd.DataFrame(df).reset_index()
    # 字段改名
    df.columns = ["iyear", "imonth", "tid", "platform", "shopcode", "income_amount", "expend_amount"]  # "created",

    return df


def combine_db():
    OMS_engine = create_engine(
        'mysql://{}:{}@{}:{}/{}'.format(OMS_PROD_DB_USER, OMS_PROD_DB_PWD, OMS_PROD_DB_HOST, OMS_PROD_DB_PORT,
                                        OMS_PROD_DB_NAME),
        echo=True,
        isolation_level='AUTOCOMMIT')

    # for val in itertools.product(bill_month):
    #     df = download_bill_records(OMS_engine, bill_year, val[0])
    #     filename = LOCAL_FILE_PATH +os.sep+ 'bill-' + str(bill_year) + '_' + str(val[0]).zfill(2) + '.pkl'
    #     df.to_pickle(filename)
    #     print("下载{}-{}账单成功".format(bill_year, str(val[0]).zfill(2)))

    for val in itertools.product(bill_month):
        filename = LOCAL_FILE_PATH + os.sep + 'bill-' + str(bill_year) + '_' + str(val[0]).zfill(2) + '.pkl'
        df_0 = pd.read_pickle(filename)

        for platform in platform_list:
            print(filename + " 正在汇总平台 ", platform)
            df = df_0[df_0["platform"].isin([platform])]
            print("抽查抽查")
            print(df.head(10).to_markdown())
            print("记录数:", df.shape[0])

            if platform == "TAOBAO":
                if "df_TAOBAO" in vars():
                    df_TAOBAO = df_TAOBAO.append(df)
                else:
                    df_TAOBAO = df

            if platform == "TMALL":
                if "df_TMALL" in vars():
                    df_TMALL = df_TMALL.append(df)
                else:
                    df_TMALL = df

            if platform == "DY":
                if "df_DY" in vars():
                    df_DY = df_DY.append(df)
                else:
                    df_DY = df

            if platform == "JD":
                if "df_JD" in vars():
                    df_JD = df_JD.append(df)
                else:
                    df_JD = df

            if platform == "JN":
                if "df_JN" in vars():
                    df_JN = df_JN.append(df)
                else:
                    df_JN = df

            if platform == "KAOLA":
                if "df_KAOLA" in vars():
                    df_KAOLA = df_KAOLA.append(df)
                else:
                    df_KAOLA = df

            if platform == "KS":
                if "df_KS" in vars():
                    df_KS = df_KS.append(df)
                else:
                    df_KS = df

            if platform == "PDD":
                if "df_PDD" in vars():
                    df_PDD = df_PDD.append(df)
                else:
                    df_PDD = df

            if platform == "WPH":
                if "df_WPH" in vars():
                    df_WPH = df_WPH.append(df)
                else:
                    df_WPH = df

            if platform == "XHS":
                if "df_XHS" in vars():
                    df_XHS = df_XHS.append(df)
                else:
                    df_XHS = df

            if platform == "YZ":
                if "df_YZ" in vars():
                    df_YZ = df_YZ.append(df)
                else:
                    df_YZ = df

            if platform == "ALIBABA":
                if "df_ALIBABA" in vars():
                    df_ALIBABA = df_ALIBABA.append(df)
                else:
                    df_ALIBABA = df

            if platform == "WM":
                if "df_WM" in vars():
                    df_WM = df_WM.append(df)
                else:
                    df_WM = df

            if platform == "FY":
                if "df_FY" in vars():
                    df_FY = df_FY.append(df)
                else:
                    df_FY = df

            if platform == "MD":
                if "df_MD" in vars():
                    df_MD = df_MD.append(df)
                else:
                    df_MD = df

            if platform == "BD":
                if "df_BD" in vars():
                    df_BD = df_BD.append(df)
                else:
                    df_BD = df

            if platform == "ZMB":
                if "df_ZMB" in vars():
                    df_ZMB = df_ZMB.append(df)
                else:
                    df_ZMB = df

        print("itdb_进度表：   ", str(val[0]).zfill(2))

        # 保存的文件名 LOCAL_FILE_PATH + 'bill-' + str(bill_year) + '_' + str(val[0]).zfill(2) + '.pkl'
        save_filename = LOCAL_FILE_PATH + os.sep + "it_DB"
        if "df_TAOBAO" in vars():
            save_filename = LOCAL_FILE_PATH + os.sep + "it_DB_{}.pkl".format("TAOBAO")
            df_TAOBAO.to_pickle(save_filename)

        if "df_TMALL" in vars():
            save_filename = LOCAL_FILE_PATH + os.sep + "it_DB_{}.pkl".format("TMALL")
            df_TAOBAO.to_pickle(save_filename)

        if "df_DY" in vars():
            save_filename = LOCAL_FILE_PATH + os.sep + "it_DB_{}.pkl".format("DY")
            df_DY.to_pickle(save_filename)

        if "df_JD" in vars():
            save_filename = LOCAL_FILE_PATH + os.sep + "it_DB_{}.pkl".format("JD")
            df_JD.to_pickle(save_filename)

        if "df_JN" in vars():
            save_filename = LOCAL_FILE_PATH + os.sep + "it_DB_{}.pkl".format("JN")
            df_JN.to_pickle(save_filename)

        if "df_KAOLA" in vars():
            save_filename = LOCAL_FILE_PATH + os.sep + "it_DB_{}.pkl".format("KAOLA")
            df_KAOLA.to_pickle(save_filename)

        if "df_KS" in vars():
            save_filename = LOCAL_FILE_PATH + os.sep + "it_DB_{}.pkl".format("KS")
            df_KS.to_pickle(save_filename)

        if "df_PDD" in vars():
            save_filename = LOCAL_FILE_PATH + os.sep + "it_DB_{}.pkl".format("PDD")
            df_PDD.to_pickle(save_filename)

        if "df_WPH" in vars():
            save_filename = LOCAL_FILE_PATH + os.sep + "it_DB_{}.pkl".format("WPH")
            df_WPH.to_pickle(save_filename)

        if "df_XHS" in vars():
            save_filename = LOCAL_FILE_PATH + os.sep + "it_DB_{}.pkl".format("XHS")
            df_XHS.to_pickle(save_filename)

        if "df_YZ" in vars():
            save_filename = LOCAL_FILE_PATH + os.sep + "it_DB_{}.pkl".format("YZ")
            df_YZ.to_pickle(save_filename)

        if "df_ALIBABA" in vars():
            save_filename = LOCAL_FILE_PATH + os.sep + "it_DB_{}.pkl".format("ALIBABA")
            df_ALIBABA.to_pickle(save_filename)

        if "df_WM" in vars():
            save_filename = LOCAL_FILE_PATH + os.sep + "it_DB_{}.pkl".format("WM")
            df_WM.to_pickle(save_filename)

        if "df_FY" in vars():
            save_filename = LOCAL_FILE_PATH + os.sep + "it_DB_{}.pkl".format("FY")
            df_FY.to_pickle(save_filename)

        if "df_MD" in vars():
            save_filename = LOCAL_FILE_PATH + os.sep + "it_DB_{}.pkl".format("MD")
            df_MD.to_pickle(save_filename)

        if "df_BD" in vars():
            save_filename = LOCAL_FILE_PATH + os.sep + "it_DB_{}.pkl".format("BD")
            df_BD.to_pickle(save_filename)

        if "df_ZMB" in vars():
            save_filename = LOCAL_FILE_PATH + os.sep + "it_DB_{}.pkl".format("ZMB")
            df_ZMB.to_pickle(save_filename)


# 生成两张表格的差异分析报告
def get_defferent_beyondcompare(df_1, df_2, key_columns, value_columns, number_columns, leftname, rightname,
                                bdrop_duplicates):
    # 假设不重复，假设字段顺序完全一致

    if bdrop_duplicates and len(value_columns) == 1:
        print("开启合并对比模式！")

    # df1 = df_1.reset_index().copy()
    # df2 = df_2.reset_index().copy()

    print("{}原始记录数: {} VS  {}原始记录数: {}".format(leftname, df_1.shape[0], rightname, df_2.shape[0]))

    # 标记原始表的序号
    df_1["originaliid"] = df_1.index
    df_2["originaliid"] = df_2.index

    # list 做减法操作，剔除重复的字段
    other_columns = list(set(df_1.columns.to_list()) - set(key_columns) - set(value_columns))
    # np.array(df_1.columns.to_list())
    # print("备注性质的字段有：",other_columns)

    all_columns = key_columns + value_columns + ["originaliid"]
    # print("all_columns",all_columns)

    if bdrop_duplicates and len(value_columns) > 1:
        raise "分组合并暂不支持多项值！"

    # 如果只有一个数字列，支持分组求和比对
    if bdrop_duplicates and len(value_columns) == 1:
        all_columns = key_columns + value_columns

        value_col = "".join(value_columns)

        df_1 = df_1.groupby(key_columns).agg({value_col: np.sum})
        df_1 = pd.DataFrame(df_1).reset_index()
        df_1.columns = all_columns
        # del df_1["index"]

        df_2 = df_2.groupby(key_columns).agg({value_col: np.sum})
        df_2 = pd.DataFrame(df_2).reset_index()
        df_2.columns = all_columns
        # del df_2["index"]

    # print("抽查:")
    # print(df_1.head(5).to_markdown())
    # print(df_2.head(5).to_markdown())

    df1 = df_1[all_columns].copy()
    df2 = df_2[all_columns].copy()

    df1["value_combine"] = ""
    df2["value_combine"] = ""
    for col in value_columns:
        df1["value_combine"] = df1["value_combine"] + "|" + df1[col].astype(str)
        df2["value_combine"] = df2["value_combine"] + "|" + df2[col].astype(str)

    df1["value_combine"] = df1["value_combine"].apply(lambda x: x[1:])
    df2["value_combine"] = df2["value_combine"].apply(lambda x: x[1:])

    df_left = df1.copy()
    df_right = df2.copy()

    for col in value_columns:
        del df_left[col]
        del df_right[col]

    print("整理后的左表:", df_left.shape[0], "右表:", df_right.shape[0])

    # print("左表抽样：")
    # print(df_left.head(1).to_markdown())
    #
    # print("右表抽样：")
    # print(df_right.head(1).to_markdown())

    # value_column= "".join(df_left.columns[-1:])
    # print("value_column=",value_column)

    # 附加一个索引列
    df_left["uniqueindex"] = df_left.iloc[:, 0]
    df_right["uniqueindex"] = df_right.iloc[:, 0]

    # 拼接生成索引行
    # print(df_left.columns)
    # 第一列先设置好，最后一列是计算列，不加入索引
    for c in key_columns[1:]:
        # print('字段:',c)
        df_left[c] = df_left[c].astype(str)
        df_right[c] = df_right[c].astype(str)
        df_left["uniqueindex"] = df_left.apply(lambda x: str(x["uniqueindex"]) + "|" + str(x[c]), axis=1)
        df_right["uniqueindex"] = df_right.apply(lambda x: str(x["uniqueindex"]) + "|" + str(x[c]), axis=1)

    # print("左表抽样2：")
    # print(df_left.head(1).to_markdown())
    #
    # print("右表抽样2：")
    # print(df_right.head(1).to_markdown())

    if bdrop_duplicates:  # 有压缩，就会失去原有的索引编号
        df_result = pd.DataFrame(
            columns=["uniqueindex_1", "value_combine_1", "match", "uniqueindex_2", "value_combine_2"])
        # 左边多
        # print("列名：",["uniqueindex"]+value_columns)
        df_leftmore = df_left[~df_left["uniqueindex"].isin(df_right["uniqueindex"])][
            ["uniqueindex", "value_combine"]].copy()
        df_leftmore = pd.DataFrame(df_leftmore)
        df_leftmore["match"] = "+-"
        df_leftmore["other"] = ""
        df_leftmore["value_combine" + "_2"] = ""
        print("左边多:", df_leftmore.shape[0])
        temp_df = df_leftmore[
            ["uniqueindex", "value_combine", "match", "other", "value_combine" + "_2"]].copy()
        temp_df.columns = ["uniqueindex_1", "value_combine" + "_1", "match", "uniqueindex_2",
                           "value_combine" + "_2"]
        # print(temp_df.head(10).to_markdown())
        df_result = df_result.append(temp_df)

        # 右边多
        df_rightmore = df_right[~df_right["uniqueindex"].isin(df_left["uniqueindex"])][
            ["uniqueindex", "value_combine"]].copy()
        df_rightmore = pd.DataFrame(df_rightmore)
        df_rightmore["match"] = "-+"
        df_rightmore["other"] = ""
        df_rightmore["value_combine" + "_1"] = ""
        print("右边多:", df_rightmore.shape[0])
        temp_df = df_rightmore[
            ["other", "value_combine" + "_1", "match", "uniqueindex", "value_combine"]].copy()
        temp_df.columns = ["uniqueindex_1", "value_combine" + "_1", "match", "uniqueindex_2",
                           "value_combine" + "_2"]
        # print(temp_df.head(10).to_markdown())
        df_result = df_result.append(temp_df)

        # 两张表联合查询
        df_inner = df_left.merge(df_right, how="inner", on="uniqueindex")
        df_inner["other"] = df_inner["uniqueindex"]
        # print("联合查询")
        # print(df_inner.head(10).to_markdown())

        # 两边相等
        # df_inner_equal=df_inner.loc[  abs(df_inner[value_column+"_x"].astype(float)-df_inner[value_column+"_y"].astype(float))<0.01].copy()
        df_inner_equal = df_inner.loc[df_inner["value_combine" + "_x"] == df_inner["value_combine" + "_y"]].copy()
        df_inner_equal = pd.DataFrame(df_inner_equal)
        df_inner_equal["match"] = "相等"
        temp_df = df_inner_equal[["uniqueindex", "value_combine" + "_x", "match", "other",
                                  "value_combine" + "_y"]].copy()
        temp_df.columns = ["uniqueindex_1", "value_combine" + "_1", "match", "uniqueindex_2",
                           "value_combine" + "_2"]

        print("两边相等:", temp_df.shape[0])
        # print(temp_df.head(10).to_markdown())
        df_result = df_result.append(temp_df)

        # 两边不等于
        # df_inner_not_equal = df_inner.loc[
        #     abs(df_inner[value_column+"_x"].astype(float) - df_inner[value_column+"_y"].astype(float)) >= 0.01].copy()

        df_inner_not_equal = df_inner.loc[df_inner["value_combine" + "_x"] != df_inner["value_combine" + "_y"]].copy()
        # print(df_inner_not_equal.head(5).to_markdown())

        df_inner_not_equal["match"] = "<>"
        temp_df = df_inner_not_equal[
            ["uniqueindex", "value_combine_x", "match", "other", "value_combine_y"]].copy()
        temp_df.columns = ["uniqueindex_1", "value_combine" + "_1", "match", "uniqueindex_2",
                           "value_combine" + "_2"]

        print("两边不相等:", temp_df.shape[0])
        # print(df_inner_not_equal[["uniqueindex", "match", "uniqueindex"]].head(100).to_markdown())
        # print(temp_df.head(3).to_markdown())
        df_result = df_result.append(temp_df)

        df_result.fillna("", inplace=True)
        # print("debug1")
        # print(df_result.head(3).to_markdown())

        # 找回原来的字段列表，形成uniqueindex
        df1["uniqueindex"] = df1.iloc[:, 0]
        df2["uniqueindex"] = df2.iloc[:, 0]
        for c in key_columns[0:-2]:
            # print('字段:', c)
            # print(df1.head(2).to_markdown())
            if c != "product":
                df1[c] = df1[c].astype(str)
                df2[c] = df2[c].astype(str)
                df1["uniqueindex"] = df1.apply(lambda x: x["uniqueindex"] + "|" + x[c], axis=1)
                df2["uniqueindex"] = df2.apply(lambda x: x["uniqueindex"] + "|" + x[c], axis=1)

        # print("检查原始表")
        # print(df1.head(5).to_markdown())
        #
        # print("检查结果表")
        # print(df_result.head(5).to_markdown())

        # 把列名解压缩
        # print("列名解压缩")
        # print(df_left.columns)

        df_result["value_combine_1"] = df_result["value_combine_1"].astype(str)
        df_result["value_combine_2"] = df_result["value_combine_2"].astype(str)

        # print("拆解前")
        # print(df_result.head(5).to_markdown())

        # print(df_left.columns)
        # print(df_left.columns[0:-2])

        # print("分析字段：",df_left.columns.to_list(),df_left.columns[0:-3].to_list())
        # print(len(df_left.columns[0:-2]))
        # print(len(value_columns))

        # print("拆解出关键索引字段",df_left.columns[0:-2])
        if len(df_left.columns[0:-2]) == 1:
            col_name = "".join(df_left.columns[0:-2])
            df_result[col_name + "_1"] = df_result["uniqueindex_1"]
            df_result[col_name + "_2"] = df_result["uniqueindex_2"]
        else:
            col_index = 0
            for c in df_left.columns[0:-2]:
                # print("跟踪：",c,col_index)
                df_result[c + "_1"] = df_result.apply(
                    lambda x: x["uniqueindex_1"].split("|")[col_index] if str(x["uniqueindex_1"]).split(
                        "|").__len__() > col_index else x["uniqueindex_1"],
                    axis=1)
                df_result[c + "_2"] = df_result.apply(
                    lambda x: x["uniqueindex_2"].split("|")[col_index] if str(x["uniqueindex_2"]).split(
                        "|").__len__() > col_index else x["uniqueindex_2"],
                    axis=1)
                col_index = col_index + 1

        # print("拆解出数据字段")
        # print(df_result.head(5).to_markdown())
        if len(value_columns) == 1:
            col_name = "".join(value_columns)
            df_result[col_name + "_1"] = df_result["value_combine_1"]
            df_result[col_name + "_2"] = df_result["value_combine_2"]
        else:
            col_index = 0
            for c in value_columns:
                # print(col_index,c)
                df_result[c + "_1"] = df_result.apply(
                    lambda x: x["value_combine_1"].split("|")[col_index] if x["value_combine_1"].split(
                        "|").__len__() > col_index else x["value_combine_1"],
                    axis=1)
                df_result[c + "_2"] = df_result.apply(
                    lambda x: x["value_combine_2"].split("|")[col_index] if x["value_combine_2"].split(
                        "|").__len__() > col_index else x["value_combine_2"],
                    axis=1)
                col_index = col_index + 1

        # print("找回备注字段")
        # print(df_result.head(5).to_markdown())

        # print("找回备注字段的结果是：")
        # print(df_result.head(5).to_markdown())

        # print("修改 key_columns")
        str_column1 = ""
        str_column2 = ""
        for c in key_columns:
            str_column1 = str_column1 + "," + c + "_1"
            str_column2 = str_column2 + "," + c + "_2"

        for c in value_columns:
            str_column1 = str_column1 + "," + c + "_1"
            str_column2 = str_column2 + "," + c + "_2"

        str_column = str_column1 + ",match" + str_column2

        # print("字段列表：",str_column)
        str_column = str_column[1:].strip()  # 去掉开始的逗号
        # str_column = str_column.replace("uniqueindex","pono_")
        # print("字段列表：",str_column)
        series_column = pd.Series(str_column.split(","))
        # print(df_result[series_column].head(3).to_markdown())

        for col in df_result.columns:
            if col != "match":
                df_result[col] = df_result[col].apply(
                    lambda x: "'{}".format(x) if len(str(x)) > 0 else '')  # 强制转字符串,避免转数字
                # 设置数字型字段的格式
                for col2 in number_columns:
                    # print("col=col2 ",col,col2)
                    if col.replace("_1", "").replace("_2", "") == col2:
                        # print("col2:",col2)
                        df_result[col] = df_result[col].apply(lambda x: x.replace("'", ""))
                        # df_result[col]=df_result[col].astype(float)

        print("检查最后的比对结果：")
        print(df_result[series_column].head(3).to_markdown())

    else:
        df_result = pd.DataFrame(
            columns=["originaliid_1", "uniqueindex_1", "value_combine_1", "match", "originaliid_2", "uniqueindex_2",
                     "value_combine_2"])
        # 左边多
        # print("列名：",["uniqueindex"]+value_columns)
        df_leftmore = df_left[~df_left["uniqueindex"].isin(df_right["uniqueindex"])][
            ["originaliid", "uniqueindex", "value_combine"]].copy()
        df_leftmore = pd.DataFrame(df_leftmore)
        df_leftmore["match"] = "+-"
        df_leftmore["other"] = ""
        df_leftmore["value_combine" + "_2"] = ""
        print("左边多:", df_leftmore.shape[0])
        temp_df = df_leftmore[
            ["originaliid", "uniqueindex", "value_combine", "match", "other", "other", "value_combine" + "_2"]].copy()
        temp_df.columns = ["originaliid_1", "uniqueindex_1", "value_combine" + "_1", "match", "originaliid_2",
                           "uniqueindex_2",
                           "value_combine" + "_2"]
        # print(temp_df.head(10).to_markdown())
        df_result = df_result.append(temp_df)

        # 右边多
        df_rightmore = df_right[~df_right["uniqueindex"].isin(df_left["uniqueindex"])][
            ["originaliid", "uniqueindex", "value_combine"]].copy()
        df_rightmore = pd.DataFrame(df_rightmore)
        df_rightmore["match"] = "-+"
        df_rightmore["other"] = ""
        df_rightmore["value_combine" + "_1"] = ""
        print("右边多:", df_rightmore.shape[0])
        temp_df = df_rightmore[
            ["other", "other", "value_combine" + "_1", "match", "originaliid", "uniqueindex", "value_combine"]].copy()
        temp_df.columns = ["originaliid_1", "uniqueindex_1", "value_combine" + "_1", "match", "originaliid_2",
                           "uniqueindex_2",
                           "value_combine" + "_2"]
        # print(temp_df.head(10).to_markdown())
        df_result = df_result.append(temp_df)

        # 两张表联合查询
        df_inner = df_left.merge(df_right, how="inner", on="uniqueindex")
        df_inner["other"] = df_inner["uniqueindex"]
        # print("联合查询")
        # print(df_inner.head(10).to_markdown())

        # 两边相等
        # df_inner_equal=df_inner.loc[  abs(df_inner[value_column+"_x"].astype(float)-df_inner[value_column+"_y"].astype(float))<0.01].copy()
        df_inner_equal = df_inner.loc[df_inner["value_combine" + "_x"] == df_inner["value_combine" + "_y"]].copy()
        df_inner_equal = pd.DataFrame(df_inner_equal)
        df_inner_equal["match"] = "相等"
        temp_df = df_inner_equal[
            ["originaliid_x", "uniqueindex", "value_combine" + "_x", "match", "originaliid_y", "other",
             "value_combine" + "_y"]].copy()
        temp_df.columns = ["originaliid_1", "uniqueindex_1", "value_combine" + "_1", "match", "originaliid_2",
                           "uniqueindex_2",
                           "value_combine" + "_2"]

        print("两边相等:", temp_df.shape[0])
        # print(temp_df.head(10).to_markdown())
        df_result = df_result.append(temp_df)

        # 两边不等于
        # df_inner_not_equal = df_inner.loc[
        #     abs(df_inner[value_column+"_x"].astype(float) - df_inner[value_column+"_y"].astype(float)) >= 0.01].copy()

        df_inner_not_equal = df_inner.loc[df_inner["value_combine" + "_x"] != df_inner["value_combine" + "_y"]].copy()
        # print(df_inner_not_equal.head(5).to_markdown())

        df_inner_not_equal["match"] = "<>"
        temp_df = df_inner_not_equal[
            ["originaliid_x", "uniqueindex", "value_combine_x", "match", "originaliid_y", "other",
             "value_combine_y"]].copy()
        temp_df.columns = ["originaliid_1", "uniqueindex_1", "value_combine" + "_1", "match", "originaliid_2",
                           "uniqueindex_2",
                           "value_combine" + "_2"]

        print("两边不相等:", temp_df.shape[0])
        # print(df_inner_not_equal[["uniqueindex", "match", "uniqueindex"]].head(100).to_markdown())
        # print(temp_df.head(3).to_markdown())
        df_result = df_result.append(temp_df)

        df_result.fillna("", inplace=True)
        # print("debug1")
        # print(df_result.head(3).to_markdown())

        # 找回原来的字段列表，形成uniqueindex
        df1["uniqueindex"] = df1.iloc[:, 0]
        df2["uniqueindex"] = df2.iloc[:, 0]
        for c in key_columns[0:-2]:
            # print('字段:', c)
            # print(df1.head(2).to_markdown())
            if c != "product":
                df1[c] = df1[c].astype(str)
                df2[c] = df2[c].astype(str)
                df1["uniqueindex"] = df1.apply(lambda x: str(x["uniqueindex"]) + "|" + str(x[c]), axis=1)
                df2["uniqueindex"] = df2.apply(lambda x: str(x["uniqueindex"]) + "|" + str(x[c]), axis=1)

        # print("检查原始表")
        # print(df1.head(5).to_markdown())
        #
        # print("检查结果表")
        # print(df_result.head(5).to_markdown())

        # 把列名解压缩
        # print("列名解压缩")
        # print(df_left.columns)

        df_result["value_combine_1"] = df_result["value_combine_1"].astype(str)
        df_result["value_combine_2"] = df_result["value_combine_2"].astype(str)

        # print("拆解前")
        # print(df_result.head(5).to_markdown())

        # print(df_left.columns)
        # print(df_left.columns[0:-2])

        # print("分析字段：",df_left.columns.to_list(),df_left.columns[0:-3].to_list())
        # print(len(df_left.columns[0:-2]))
        # print(len(value_columns))

        # 拆解出关键索引字段
        if len(df_left.columns[0:-3]) == 1:
            col_name = "".join(df_left.columns[0:-2])
            df_result[col_name + "_1"] = df_result["uniqueindex_1"]
            df_result[col_name + "_2"] = df_result["uniqueindex_2"]
        else:
            col_index = 0
            for c in df_left.columns[0:-3]:
                # print("跟踪：",c,col_index)
                df_result[c + "_1"] = df_result.apply(
                    lambda x: x["uniqueindex_1"].split("|")[col_index] if str(x["uniqueindex_1"]).split(
                        "|").__len__() > col_index else x["uniqueindex_1"],
                    axis=1)
                df_result[c + "_2"] = df_result.apply(
                    lambda x: x["uniqueindex_2"].split("|")[col_index] if str(x["uniqueindex_2"]).split(
                        "|").__len__() > col_index else x["uniqueindex_2"],
                    axis=1)
                col_index = col_index + 1

        # print("拆解出数据字段")
        # print(df_result.head(5).to_markdown())
        if len(value_columns) == 1:
            col_name = "".join(value_columns)
            df_result[col_name + "_1"] = df_result["value_combine_1"]
            df_result[col_name + "_2"] = df_result["value_combine_2"]
        else:
            col_index = 0
            for c in value_columns:
                # print(col_index,c)
                df_result[c + "_1"] = df_result.apply(
                    lambda x: x["value_combine_1"].split("|")[col_index] if x["value_combine_1"].split(
                        "|").__len__() > col_index else x["value_combine_1"],
                    axis=1)
                df_result[c + "_2"] = df_result.apply(
                    lambda x: x["value_combine_2"].split("|")[col_index] if x["value_combine_2"].split(
                        "|").__len__() > col_index else x["value_combine_2"],
                    axis=1)
                col_index = col_index + 1

        df_result["originaliid_1"].fillna("-1", inplace=True)
        df_result["originaliid_2"].fillna("-1", inplace=True)

        df_result["originaliid_1"] = df_result["originaliid_1"].apply(lambda x: "-1" if len(str(x)) == 0 else x)
        df_result["originaliid_2"] = df_result["originaliid_2"].apply(lambda x: "-1" if len(str(x)) == 0 else x)

        # print("找回备注字段")
        # print(df_result.head(5).to_markdown())

        df_result["originaliid_1"] = df_result["originaliid_1"].astype(int)
        df_result["originaliid_2"] = df_result["originaliid_2"].astype(int)

        df_result = df_result.merge(df_1[other_columns], how="left", left_on="originaliid_1", right_on="originaliid")
        df_result = df_result.merge(df_2[other_columns], how="left", left_on="originaliid_2", right_on="originaliid")

        # print("找回备注字段的结果是：")
        # print(df_result.head(5).to_markdown())

        # print("修改 key_columns")
        str_column1 = ""
        str_column2 = ""
        for c in key_columns:
            str_column1 = str_column1 + "," + c + "_1"
            str_column2 = str_column2 + "," + c + "_2"

        for c in other_columns:
            str_column1 = str_column1 + "," + c + "_1"
            str_column2 = str_column2 + "," + c + "_2"
            df_result.rename(columns={c + "_x": c + "_1"}, inplace=True)
            df_result.rename(columns={c + "_y": c + "_2"}, inplace=True)

        for c in value_columns:
            str_column1 = str_column1 + "," + c + "_1"
            str_column2 = str_column2 + "," + c + "_2"

        str_column = str_column1 + ",match" + str_column2

        # 去掉自建的索引列
        str_column = str_column.replace("originaliid_1,", "").replace("originaliid_2,", "")

        # print("字段列表：",str_column)
        str_column = str_column[1:].strip()  # 去掉开始的逗号
        # str_column = str_column.replace("uniqueindex","pono_")
        # print("字段列表：",str_column)

        # print("重新按照默认的字段顺序进行调整！")
        new_column1 = df_1.columns.to_list()
        new_column2 = []
        # new_column2=list(set(new_column1).intersection(set(  str_column.replace("_1","").replace("_2","").split(",")  )))
        for c in new_column1:
            if c in str_column.replace("_1", "").replace("_2", "").split(","):
                new_column2.append(c)

        # print("重叠的字段有：",new_column2)

        str_new_column1 = ""
        str_new_column2 = ""
        for c in new_column2:
            str_new_column1 = str_new_column1 + "".join(c) + "_1,"
            str_new_column2 = str_new_column2 + "".join(c) + "_2,"

        str_new_column3 = str_new_column1 + "match," + str_new_column2[:-1] + ",uniqueindex_1,uniqueindex_2"
        # print("重新拼接后的字段：",str_new_column3)

        series_column = pd.Series(str_new_column3.split(","))
        # print(df_result[series_column].head(3).to_markdown())
        # print(df_result.head(3).to_markdown())

        # for col in df_result.columns:
        #     if col != "match":
        #         print("test",col)
        #         df_result[col] = df_result[col].apply(
        #             lambda x:"'{}".format(x) if len(str(x)) > 0 else '')  # 强制转字符串,避免转数字
        #         # 设置数字型字段的格式
        #         for col2 in number_columns:
        #             # print("col=col2 ",col,col2)
        #             if col.replace("_1","").replace("_2","") == col2:
        #                 # print("col2:",col2)
        #                 df_result[col] = df_result[col].apply(lambda x:x.replace("'",""))
        #                 # df_result[col]=df_result[col].astype(float)

        # number_columns
        # df_result[series_column].to_excel(r"/Users/vicetone/lclproject/python/ITDD/data/财务数据/卖家联合/比对结果(左 {},右 {}).xlsx".format(leftname,rightname))

        print(df_result[series_column].head(3).to_markdown())

    return df_result[series_column]


# def cal_platform_different(platform):


if __name__ == '__main__':
    print('开始计算...')
    all_begin_time = time.time()
    print("开始:", time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()))

    # 账单合并，按平台，按年月区分，生成pickle文件
    # combine_excel(REMOTE_FILE_PATH, ".xlsx|.xls 账单")
    # sys.exit()

    # 合并所有从数据库导出的账单明细记录，按平台分拆
    # combine_db()

    # 保存差异
    chayi = []
    for platform in platform_list:
        # platform="TAOBAO"
        if os.path.exists(LOCAL_FILE_PATH + os.sep + "it_DB_{}.pkl".format(platform)):
            if os.path.exists(LOCAL_FILE_PATH + os.sep + "fn_EXCEL_{}.pkl".format(platform)):
                df_it = pd.read_pickle(LOCAL_FILE_PATH + os.sep + "it_DB_{}.pkl".format(platform))
                df_fn = pd.read_pickle(LOCAL_FILE_PATH + os.sep + "fn_EXCEL_{}.pkl".format(platform))

                df_it["tid"] = df_it["tid"].astype(str)
                df_fn["tid"] = df_fn["tid"].astype(str)

                print("统计平台 ", platform)

                print("fn:")
                print(df_fn.head(5).to_markdown())

                print(df_fn[df_fn["tid"].str.contains("84515958894398")])

                print("it:")
                print(df_it.head(5).to_markdown())

                dd = get_defferent_beyondcompare(df_fn, df_it, ["iyear", "imonth", "platform", "shopcode", "tid"],
                                                 ["income_amount", "expend_amount"],
                                                 ["income_amount", "expend_amount"], "财务", "it", False)
                print("差异:")
                dd.fillna("", inplace=True)
                # print(dd.head(50).to_markdown())

                dd.to_pickle(LOCAL_FILE_PATH + os.sep + "_compare_{}.pkl".format(platform))
                for imonth in range(1, 13):
                    # 分月统计差异
                    dm = dd[dd["imonth_1"].isin([str(imonth).zfill(2)])]
                    if dm.shape[0] > 200000:
                        # dm.to_csv(LOCAL_FILE_PATH + os.sep + "差异_{}_{}.csv".format(platform, str(imonth).zfill(2)))
                        pagecount = 200000
                        pagecnt = int(dm.shape[0] / pagecount) + 1
                        for i in range(0, int(dm.shape[0] / pagecount) + 1):
                            # print("分页：{}  from:{} to:{}".format(i+1, i * 500000, (i + 1) * 500000))
                            print("分页：{}  from:{} to:{} ，记录数: {} ".format(i + 1, i * pagecount, (i + 1) * pagecount,
                                                                          dm[i * pagecount:(i + 1) * pagecount].shape[
                                                                              0]))

                            # print("合并生成:",
                            #       LOCAL_FILE_PATH + os.sep + "{}_合并表格_{}.{}.{}.xlsx".format(plat, index, pagecnt, i + 1))
                            dm[i * pagecount:(i + 1) * pagecount].to_excel(
                                LOCAL_FILE_PATH + os.sep + "差异_{}_{}.{}.xlsx".format(platform, str(imonth).zfill(2),
                                                                                     i + 1))

                    else:
                        dm.to_excel(LOCAL_FILE_PATH + os.sep + "差异_{}_{}.xlsx".format(platform, str(imonth).zfill(2)))

                    cnt1 = dm[dm["match"].str.contains("\+-")].shape[0]
                    cnt2 = dm[dm["match"].str.contains("\-+")].shape[0]
                    cnt3 = dm[dm["match"].str.contains("相等")].shape[0]
                    cnt4 = dm[dm["match"].str.contains("\<>")].shape[0]
                    chayi.append([platform, imonth, cnt1, cnt2, cnt3, cnt4])

                    print(chayi)

    chayi = pd.DataFrame(chayi).reset_index()
    chayi.columns = ["id", "平台", "月份", "左边多", "右边多", "相等", "不等"]
    chayi.to_excel(LOCAL_FILE_PATH + os.sep + "差异汇总.xlsx")

    all_end_time = time.time()
    print("结束:", time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()))
    print('计算完成，从', bill_year, str(bill_month[0]).zfill(2), ' - ', bill_year,
          str(bill_month[len(bill_month) - 1]).zfill(2), '总消耗时间:', all_end_time - all_begin_time, '秒')
