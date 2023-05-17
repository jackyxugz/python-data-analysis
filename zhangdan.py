# coding=utf-8

import  os
import pandas as pd
import  time
import datetime
# from sqlalchemy import create_engine
# import pymysql
import os.path
import xlwt
import zipfile
# import unrar
from pathlib import Path
# import msoffcrypto
import io
# import rarfile
import re

'''
pip install unrar
pip install rarfile
pip install msoffcrypto-tool
'''

# 设置显示的最大宽度
pd.set_option('display.max_colwidth', 1024)


# def read_excel_with_password(filename,password):
#     decrypted = io.BytesIO()
#     with open(filename, "rb") as f:
#         file = msoffcrypto.OfficeFile(f)
#         file.load_key(password=password)  # Use password
#         file.decrypt(decrypted)
#
#     df = pd.read_excel(decrypted)
#     print(df)


def un_zip_rar(zip_filename):
    # filename=zip_filename.split("\\")[-1]
    # folder_abs=zip_filename.replace(filename,"")

    if zip_filename.find(".zip")>0:
        folder_abs = zip_filename.replace(".zip", "")+"_zip"
        zip_file = zipfile.ZipFile(zip_filename)
        zip_list = zip_file.namelist()  # 得到压缩包里所有文件

        for fileM in zip_list:
            #f = unicode(f, 'cp936')
            # 循环解压文件到指定目录
            new_filename=folder_abs + '\\' + fileM.encode('cp437').decode('gbk')

            if True:
            #if os.path.isdir(folder_abs+ '\\'):
                if os.path.exists(new_filename):
                    # 文件已存在，不需要重复解压缩了
                    # print("{}文件已经存在！".format(new_filename))
                    pass
                else:
                    # print("{}文件不存在！".format(new_filename))
                    extracted_path = Path(zip_file.extract(fileM, folder_abs))
                    # f = f.encode(encoding='UTF-8', errors='strict')
                    # 文件重命名，将中文的文件名还原
                    extracted_path.rename(new_filename)
            else:
                print("目录错误！")

        zip_file.close()  # 关闭文件，必须有，释放内存
    elif zip_filename.find(".rar")>0:
        folder_abs = zip_filename.replace(".rar", "")+"_rar"

        # print(folder_abs)
        # rf = rarfile.RarFile(zip_filename, mode='r')  # mode的值只能为'r'
        #
        # # rf_list = rf.namelist()  # 得到压缩包里所有的文件
        # # print('rar文件内容', rf_list)
        # # for f in rf_list:
        # #     rf.extract(f, folder_abs)  # 循环解压，将文件解压到指定路径
        #
        # # 一次性解压所有文件到指定目录
        # rf.extractall(folder_abs) # 不传path，默认为当前目录
        # 解压缩rar到指定文件夹

        # 创建解压缩目录
        if os.path.isdir(folder_abs):
            pass
        else:
            os.mkdir(folder_abs)
        os.chdir(folder_abs)

        # 调用本地rar.exe进行解压缩
        rar_command1 = "WinRAR.exe x -ibck %s %s" % (zip_filename, folder_abs)

        # 这里需要指定本机的winrar.exe目录！！！！！！！！！！！！
        rar_command2 = r'"C:\Program Files\WinRAR\WinRAR.exe"   X -o+   -ibck %s %s' % (zip_filename, folder_abs)
        if os.system(rar_command1) == 0:
            print
            "Path OK."
        else:
            if os.system(rar_command2) != 0:
                print
                "Error."
            else:
                print
                "Exe OK"


    return list_all_files(folder_abs)


# def un_rar(file_name):
#     """unrar zip file"""
#     rar = rarfile.RarFile(file_name)
#     if os.path.isdir(file_name + "_files"):
#         pass
#     else:
#         os.mkdir(file_name + "_files")
#     os.chdir(file_name + "_files")
#     rar.extractall()
#     rar.close()


def list_all_files(rootdir):
    _files = []
    list = os.listdir(rootdir) #列出文件夹下所有的目录与文件
    for i in range(0,len(list)):
        path = os.path.join(rootdir,list[i])
        if os.path.isdir(path):
           _files.extend(list_all_files(path))
        if os.path.isfile(path):
          if path.find("~")<0:
            if path.find("xlsx")>0:
                _files.append(path)
            elif path.find("xls")>0:
                _files.append(path)
            elif path.find(".csv")>0:
                _files.append(path)

    return _files

def un_all_zip_rar(rootdir):
    # 解压缩目录下所有的zip和rar文件
    _files = []
    list = os.listdir(rootdir)  # 列出文件夹下所有的目录与文件
    for i in range(0, len(list)):
        path = os.path.join(rootdir, list[i])
        if os.path.isdir(path):
            #_files.extend(list_all_files(path))
            un_all_zip_rar(path)
        if os.path.isfile(path):
            if path.find("~") < 0:
                if path.find(".zip") > 0:
                    un_zip_rar(path)
                elif path.find(".rar") > 0:
                    un_zip_rar(path)


def get_all_files(rootdir):
    # rootdir = r"D:\数据处理\销售统计\test"
    # 先解压缩文件
    un_all_zip_rar(rootdir)
    filelist = list_all_files(rootdir)
    mySeries = pd.Series(filelist)
    df = pd.DataFrame(mySeries)
    df.columns = ["filename"]
    df["filename"] = df["filename"].apply(lambda x: re.sub('\s|\t|\n', '', x.strip().replace(" ", "")))

    # df.to_excel(r"data\ttt.xlsx")
    # print(df)
    return df

def download_bill(shopname,iyear,imonth):
    engine = create_engine(
        'mysql+pymysql://user203_select:li123456@10.20.16.203:3308/megaorder-{}'.format(iyear))

    # 1.查询语句
    sql = (
        " select  id,tid,created,income_amount,expend_amount  from order_info_bill_{:0>2d}  where  SHOPNAME='{}'  ").format(int(imonth),shopname)
    df = pd.read_sql_query(sql, engine)
    df.to_pickle(r"data\bill_{:0>2d}_{:0>2d}_{}.pkl".format(int(iyear),int(imonth),shopname))

def merge_bill(shopname):
    yearmonth_list=["2020.12","2021.01"]
    # "2020.01","2020.02","2021.02"
    # ,"2020.03","2020.04","2020.05","2020.06","2020.07","2020.08","2020.09","2020.10","2020.11","2020.12","2021.01","2021.02","2021.03","2021.04","2021.05"
    df=pd.DataFrame(columns=["id","tid","created","income_amount","expend_amount"])
    for ym in yearmonth_list:
        iyear=ym.split(".")[0]
        imonth = ym.split(".")[1]
        filename=r"data\bill_{:0>2d}_{:0>2d}_{}.pkl".format(int(iyear),int(imonth),shopname)
        if os.path.exists(filename):
            pass
        else:
            print("{}文件不存在！".format(filename))
            download_bill(shopname,iyear,imonth)

        _df=pd.read_pickle(filename)
        df=df.append(_df)

    # if os.path.isdir(shopname):
    #     pass
    # else:
    #     os.mkdir(shopname)

    df.to_pickle(r"data\bill_last_{}_{}.pkl".format( ''.join(yearmonth_list[-1:]),shopname))

def query_oms_bill(shopname,iyear,imonth):
    rownumber = 0
    filename=r"data\bill_last_2021.01_smilelab海外旗舰店.pkl"
    if os.path.exists(filename):
        pass
    else:
        print("合并文件不存在！")
        merge_bill(shopname)

    _df=pd.read_pickle(filename)
    _df["tid"]=_df["tid"].apply(lambda x:x.replace("\"","").strip())
    # df.fillna(0)
    # 剔除tid=''
    # print(df.shape[0])
    # df=df[df.tid.str.len()>0]
    # print(df.shape[0])

    # print(df.sort_values("income_amount",ascending=False).head(10).to_markdown())

    #

    df=_df.groupby(["tid"]).agg({"income_amount":np.sum,"expend_amount":np.sum,"created":np.max}).reset_index()
    df=pd.DataFrame(df)
    df["income_amount"]=df["income_amount"].astype(float)
    df["expend_amount"]=df["expend_amount"].astype(float)

    # 剔除income为0的记录
    df=df[df.income_amount>0]

    df["expend_amount"]=df["expend_amount"].apply(lambda x:-x)
    # print("抽查数据库数据~！")
    # print(df.head(10).to_markdown())
    df["iyear"]=df["created"].apply(lambda x: x.year)  #datetime.datetime.strptime(x,"%Y-%m-%d %H:%M:%S").year
    df["imonth"] = df["created"].apply(lambda x: x.month) # datetime.datetime.strptime(x,"%Y-%m-%d %H:%M:%S").month
    df.rename(columns={"income_amount": "income", "expend_amount": "fee"}, inplace=True)
    df.reset_index(inplace=True)

    # print("oms bill...")
    # print(df.head(10).to_markdown())

    df=df[  df.imonth.astype(int)==int(imonth)]
    df = df[df.iyear.astype(int) == int(iyear)]

    rownumber = _df[_df.tid.isin(df["tid"])].shape[0]
    #_df[_df.tid.isin(df["tid"])].to_excel(r"data\oms.xls")


    return df,rownumber


def query_settlement(settle_filename,fee_filename,iyear,imonth):
    rownumber=0
    # master
    df_settle=pd.DataFrame(columns=["Partner_transaction_id","Rmb_amount","Rmb_settlement","Payment_time"])
    for file in  settle_filename:
        # print(pd.read_csv(file,encoding="ansi").head(1).to_markdown())
        df_settle=df_settle.append(pd.read_csv(file,encoding="ansi",dtype = {'Partner_transaction_id' : str})[["Partner_transaction_id", "Rmb_amount", "Rmb_settlement", "Payment_time"]])

    #details
    df_fee =pd.DataFrame(columns=["Transaction_id", "Partner_transaction_id", "Rmb_gross_amount", "Fee_rmb_amount"])
    for file in  fee_filename:
        df_fee=df_fee.append(pd.read_csv(file,encoding="ansi",dtype = {'Partner_transaction_id' : str})[["Transaction_id","Partner_transaction_id", "Rmb_gross_amount", "Fee_rmb_amount"]])
    # if len(fee2_filename)>0:
    #     df_fee2 = pd.read_csv(fee2_filename)["Transaction_id", "Partner_transaction_id", "Rmb_gross_amount", "Fee_rmb_amount"]
    #     df_fee1.append(df_fee1)

    rownumber = df_settle.shape[0]+df_fee.shape[0]

    df=df_settle.merge(df_fee,how="inner",on="Partner_transaction_id")

    df["Rmb_amount"]=df["Rmb_amount"].astype(float)
    df["Fee_rmb_amount"] = df["Fee_rmb_amount"].astype(float)

    # df.rename(columns={"Partner_transaction_id": "tid"}, inplace=True)


    #df.to_excel(r"data\seetle.xls")

    df = df.groupby("Partner_transaction_id").agg({"Rmb_amount": np.average, "Fee_rmb_amount": np.sum }).reset_index()
    df=pd.DataFrame(df)

    df["iyear"]=iyear
    df["imonth"] = imonth
    #df["Partner_transaction_id"] = df["Partner_transaction_id"].apply(lambda x: x.replace("\"", "").strip())


    df.rename(columns={"Partner_transaction_id":"tid","Rmb_amount":"income","Fee_rmb_amount":"fee"},inplace=True)
    # df["ttid"]=df["Partner_transaction_id"]

    # print("ccccc")
    # print(df.dtypes)
    # print(df.head(3).to_markdown())


    return df,rownumber


def get_defferent(_df1,_df2,keyname,valuename):

    df1=_df1[[keyname,valuename]].copy()
    df2 = _df2[[keyname, valuename]].copy()

    df1[keyname] = df1[keyname].apply(lambda x: x.strip())
    df2[keyname] = df2[keyname].apply(lambda x: x.strip())

    valuename1=valuename+"_x"
    valuename2 = valuename + "_y"

    df = df1.merge(df2, how="left", on=keyname)
    df.fillna(0,inplace=True)
    if df[abs(df[valuename1].astype(float)-df[valuename2].astype(float))>=0.01].empty:
        pass
    else:
        print("右表 异常:",df[abs(df[valuename1].astype(float)-df[valuename2].astype(float))>=0.01].shape[0])
        print(df[abs(df[valuename1].astype(float)-df[valuename2].astype(float))>=0.01])
        return

    df = df1.merge(df2, how="right", on=keyname)
    df.fillna(0,inplace=True)
    #print(df.to_markdown())
    if df[abs(df[valuename1].astype(float) - df[valuename2].astype(float)) >= 0.01].empty:
        pass
    else:
        print("左表 异常:",df[abs(df[valuename1].astype(float) - df[valuename2].astype(float)) >= 0.01].shape[0])
        print(df[abs(df[valuename1].astype(float) - df[valuename2].astype(float)) >= 0.01])
        return

    if df2[~df2.tid.isin(df1["tid"])].empty:
        pass
    else:
        print("左表少这些")
        print(df2[~df2.tid.isin(df1["tid"])])
        return

    if df1[~df1.tid.isin(df2["tid"])].empty:
        pass
    else:
        print("右表少这些")
        print(df1[~df1.tid.isin(df2["tid"])])
        return

    print("左右表相等！")


def match_bill(shopname,settle_files,fee_files,iyear,imonth):
    #settle_files=[r"data\2088231304534403_settle_202101_50002020122400032007000074148385.csv",r"data\2088231304534403_settle_202101_50002021010500032007000074673732.csv"]
    #fee_files=[r"data\2088231304534403_fee_202101.csv"]
    df1,row1 = query_settlement(settle_files,fee_files,iyear,imonth)
    df2,row2 = query_oms_bill(shopname,iyear, imonth)

    # print(df1.head(10).to_markdown())
    # print(df2.head(10).to_markdown())
    income_sum1 = df1["income"].sum()
    income_sum2 = df2["income"].sum()
    fee_sum1 = df1["fee"].sum()
    fee_sum2 = df2["fee"].sum()

    # print(df1.head(3).to_markdown())
    # print(df2.head(3).to_markdown())

    if row1  == row2:
        if abs(income_sum1-income_sum2)<0.01:
                if (abs(fee_sum1 - fee_sum2)) < 0.01:
                    print("对账平衡 OK")
                else:
                 print("\n支出金额不等:oms:{:.2f}<>财务:{:.2f}\n".format(fee_sum2, fee_sum1))
                 get_defferent(df1, df2, "tid", "fee")
        else:
            print("\n收入不等:oms:{:.2f}<>财务:{:.2f}\n".format(income_sum2,income_sum1))
            get_defferent(df1, df2, "tid", "income")

    else:
        print("\n记录数字不等:{}<>{}".format(row2,row1))
        get_defferent(df1, df2, "tid", "income")
        get_defferent(df1, df2, "tid", "fee")

    # print(df2[df2.tid.isin(df1["tid"])])
    # df1.to_excel(r"data\biao.xlsx")
    # df2.to_excel(r"data\oms.xlsx")
    # print(df1.shape[0],df2.shape[0])
    print("\nOMS",df2.shape[0],income_sum2,df2.shape[0]*2,"{:.2f}\n".format(-fee_sum2),
          "\n财务",df1.shape[0],income_sum1,df1.shape[0]*2,"{:.2f}\n".format(-fee_sum1))

def match_all_bill(filename):
    df=pd.read_excel(filename)
    df["iyear"]=df["iyear"].astype(int)
    df["imonth"] = df["imonth"].astype(int)
    shopyearmonth=df.drop_duplicates(subset=['shopname', 'iyear','imonth'], keep='first')  # , inplace=True
    shopyearmonth=pd.DataFrame(shopyearmonth)
    for index, row in shopyearmonth.iterrows():
        shopname= row["shopname"]
        iyear = row["iyear"]
        imonth = row["imonth"]

        files_list = df[df.shopname.str.contains(shopname)]
        files_list = files_list[files_list.imonth == imonth]
        files_list = files_list[files_list.iyear == iyear]

        # print(files_list)

        settle_files=files_list[files_list.filetype.str.contains("settle")]["filename"].to_list()
        fee_files = files_list[files_list.filetype.str.contains("fee")]["filename"].to_list()
        #
        # print(settle_files)
        # print(fee_files)
        print(shopname,iyear,imonth)
        print(settle_files)
        print(fee_files)

        # match_bill(row["shopname"] ,settle_files,fee_files, row["iyear"], row["imonth"])



def read_taobao(filename):
    # 读取csv文件，忽略前4行
    df=pd.read_csv(filename,header=4, encoding = "ansi")
    df=pd.DataFrame(df)

    # 删除末尾行
    df=df[~df.类型.str.contains("账务汇总列表结束")]
    df = df[~df.类型.str.contains("导出时间")]

    income_count=df[df.类型.str.contains("合计")]["收入笔数"].iloc[0]
    income_amount=df[df.类型.str.contains("合计")]["收入金额（+元）"].iloc[0]
    expend_count=df[df.类型.str.contains("合计")]["支出笔数"].iloc[0]
    expend_amount=df[df.类型.str.contains("合计")]["支出金额（-元）"].iloc[0]

    # 数据类型强制转换
    income_count=int(income_count)
    expend_count = int(expend_count)

    # print("本地文件查询")
    # print("收入笔数:",income_count)
    # print("收入金额（+元）:", income_amount)
    # print("支出笔数:", expend_count)
    # print("支出金额（-元）:", expend_amount)
    # print(df.head(10).to_markdown())
    # print(df.head(10).to_markdown())

    return  income_count,income_amount,expend_count,expend_amount


def read_douyin(filename):
    # 读取excel文件
    df=pd.read_excel(filename)
    df=pd.DataFrame(df)[["订单收入(元)","订单支出(元)"]]

    income_count=df[df["订单收入(元)"].astype(float)>0].shape[0]
    income_amount=df["订单收入(元)"].sum()
    expend_count=df[df["订单支出(元)"].astype(float)>0].shape[0]
    expend_amount=df["订单支出(元)"].sum()

    # print("本地文件查询")
    # print("收入笔数:",income_count)
    # print("订单收入(元):", income_amount)
    # print("支出笔数:", expend_count)
    # print("订单支出(元):", expend_amount)
    # print(df.head(10).to_markdown())

    return  income_count,income_amount,expend_count,expend_amount

# def query_taobao(shopname,iyear,imonth):
#     engine = create_engine(
#         'mysql+pymysql://user203_select:li123456@10.20.16.203:3307/megaorder-{}'.format(iyear))
#     # 查询语句
#     sql = (" SELECT   COUNT(INCOME_AMOUNT) income_count,SUM(INCOME_AMOUNT) income_amount  FROM `order_info_bill_{:0>2d}` WHERE SHOPNAME = '{}' AND IFNULL(INCOME_AMOUNT,'')!='0.00';   ").format(imonth, shopname)
#     df = pd.read_sql_query(sql, engine)
#     df.fillna(0, inplace=True)
#     income_count=df["income_count"].iloc[0]
#     income_amount = df["income_amount"].iloc[0]
#
#     sql = (
#         " SELECT   COUNT(EXPEND_AMOUNT) expend_count,SUM(EXPEND_AMOUNT) expend_amount  FROM `order_info_bill_{:0>2d}` WHERE SHOPNAME = '{}' AND IFNULL(INCOME_AMOUNT,'')!='0.00';   ").format(
#         imonth, shopname)
#     df = pd.read_sql_query(sql, engine)
#     df.fillna(0, inplace=True)
#     expend_count = df["expend_count"].iloc[0]
#     expend_amount = df["expend_amount"].iloc[0]
#
#     # print("数据库查询")
#     # print("收入笔数:", income_count)
#     # print("收入金额:", income_amount)
#     # print("支出笔数:", expend_count)
#     # print("支出金额:", expend_amount)
#
#     return income_count,income_amount,expend_count,expend_amount
#
#
#
# def query_douyin(shopname,iyear,imonth):
#     engine = create_engine(
#         'mysql+pymysql://user203_select:li123456@10.20.16.203:3307/megaorder-{}'.format(iyear))
#     # 查询语句
#     sql = (" SELECT   COUNT(INCOME_AMOUNT) income_count,SUM(INCOME_AMOUNT) income_amount  FROM `order_info_bill_{:0>2d}` WHERE SHOPNAME = '{}' AND IFNULL(INCOME_AMOUNT,'')!='0.00'  AND BUSINESS_DESCRIPTION = '实际支付';   ").format(imonth, shopname)
#     df = pd.read_sql_query(sql, engine)
#     df.fillna(0, inplace=True)
#     income_count=df["income_count"].iloc[0]
#     income_amount = df["income_amount"].iloc[0]
#
#     sql = (
#         " SELECT   COUNT(EXPEND_AMOUNT) expend_count,SUM(EXPEND_AMOUNT) expend_amount  FROM `order_info_bill_{:0>2d}` WHERE SHOPNAME = '{}' AND IFNULL(INCOME_AMOUNT,'')!='0.00'  AND BUSINESS_DESCRIPTION = '平台服务费' ;   ").format(
#         imonth, shopname)
#     df = pd.read_sql_query(sql, engine)
#     df.fillna(0, inplace=True)
#     # df["expend_amount"] = df["expend_amount"].apply(lambda x: 0 if x is None else x)
#
#     expend_count = df["expend_count"].iloc[0]
#     expend_amount = df["expend_amount"].iloc[0]
#
#     # print("数据库查询")
#     # print("收入笔数:", income_count)
#     # print("收入金额:", income_amount)
#     # print("支出笔数:", expend_count)
#     # print("支出金额:", expend_amount)
#
#     return income_count,income_amount,expend_count,expend_amount
#
# def match_bill_taobao(shopname,iyear,imonth,filename):
#     income_count1, income_amount1, expend_count1, expend_amount1=read_taobao(filename)
#     income_count2, income_amount2, expend_count2, expend_amount2=query_taobao(shopname,iyear,imonth)
#
#     print("-------------统计结果-------------------")
#
#     if abs(income_count1-income_count2)<=0.001:
#         if abs(income_amount1 - income_amount2) <= 0.001:
#             if abs(expend_count1 - expend_count2) <= 0.001:
#                 if abs(expend_amount1 - expend_amount2) <= 0.001:
#                     print("{} {}-{} 对账平衡".format(shopname,iyear,imonth))
#                 else:
#                     print("{} {}-{} 支出金额不等:{}<>{}".format(shopname,iyear,imonth,expend_amount1,expend_amount2))
#             else:
#                 print("{} {}-{} 支出笔数不等:{}<>{}".format(shopname, iyear, imonth, expend_count1, expend_count2))
#         else:
#             print("{} {}-{} 收入金额不等:{}<>{}".format(shopname, iyear, imonth, income_amount1, income_amount2))
#     else:
#         print("{} {}-{} 收入笔数不等:{}<>{}".format(shopname, iyear, imonth, income_count1, income_count2))
#
#
# def match_bill_douyin(shopname,iyear,imonth,filename):
#     income_count1, income_amount1, expend_count1, expend_amount1=read_douyin(filename)
#     income_count2, income_amount2, expend_count2, expend_amount2=query_douyin(shopname,iyear,imonth)
#
#     print("-------------统计结果-------------------")
#
#     if abs(income_count1-income_count2)<=0:
#         if abs(income_amount1 - income_amount2) <= 0.001:
#             if abs(expend_count1 - expend_count2) <= 0.001:
#                 if abs(expend_amount1 - expend_amount2) <= 0.001:
#                     print("{} {}-{} 对账平衡".format(shopname,iyear,imonth))
#                 else:
#                     print("{} {}-{} 支出金额不等:{}<>{}".format(shopname,iyear,imonth,expend_amount1,expend_amount2))
#             else:
#                 print("{} {}-{} 支出笔数不等:{}<>{}".format(shopname, iyear, imonth, expend_count1, expend_count2))
#         else:
#             print("{} {}-{} 收入金额不等:{}<>{}".format(shopname, iyear, imonth, income_amount1, income_amount2))
#     else:
#         print("{} {}-{} 收入笔数不等:{}<>{}".format(shopname, iyear, imonth, income_count1, income_count2))
#
#
#
# def sycm(filename1,filename2):
#     # 读取excel文件
#     df1=pd.read_excel(filename1)
#     df1=pd.DataFrame(df1)[["订单编号","实付金额","订单状态"]]
#     df1=df1[df1.订单状态.str.contains("交易成功")]
#     df1.rename(columns={"订单编号":"orderId","实付金额":"payFee"},inplace=True)
#     df1["orderId"]=df1["orderId"].astype(str)
#     df1["orderId"]=df1["orderId"].apply(lambda  x: x.strip())
#
#     df2 = pd.read_csv(filename2,encoding="ansi" )
#     # ,dtype={"orderid":str}
#     df2 = pd.DataFrame(df2)[["orderId", "confirmFee"]]
#     df2.rename(columns={"confirmFee": "payFee"}, inplace=True)
#     df2["orderId"] = df2["orderId"].astype(str)
#     df2["orderId"] = df2["orderId"].apply(lambda x: x.strip())
#     df2=df2[df2.payFee>0]
#
#     # print(df1)
#     # print(df2)
#
#     print(df1.shape[0], df1["payFee"].sum())
#     print(df2.shape[0],df2["payFee"].sum())
#
#     df1=df1.sort_values(["orderId"])
#     df2 = df2.sort_values(["orderId"])
#
#
#     print(df2[~df2.orderId.isin(df1["orderId"])])
#     print(df1[~df1.orderId.isin(df2["orderId"])])
#     df = df2.merge(df1,how="left",on="orderId")
#     print(df)
#     print(df[df["payFee_x"].astype(float)-df["payFee_y"].astype(float)>0])
#
#     df1.to_excel(r"data\t1.xlsx")
#     df2.to_excel(r"data\t2.xlsx")
#
#
#
#

if __name__ == "__main__":
    print("开始:", time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()))

    #match_bill_taobao("bodyaid旗舰店",2021,1,r"data\20887312684950900156_202101_账务汇总(1).csv")
    #match_bill_douyin("萌洁齿旗舰店", 2021, 1, r"data\DentylActive官方旗舰店（抖音小店）账单202101-原表(1).xlsx")

    # match_bill(2021, 1)
    # sycm(r"data\202107162003547b69d054a237315efd835ccc0fbe447e.xlsx",r"data\data_[2021-06-01-2021-06-01]@001(1).csv")
    # match_bill("smilelab海外旗舰店", 2021, 1)

    # s=r"d:\\x\y\\z.zip"
    # print(s.split("\\")[-1])

    # un_zip_rar(r"D:\数据处理\销售统计\test\test_zip.zip")
    # un_zip_rar(r"D:\数据处理\销售统计\test\test.rar")
    #un_rar(r"D:\数据处理\销售统计\test\test.rar")

    rootdir = r"/Users/maclove/Downloads/财务数据/neworder-2019/拼多多/账单"
    df=get_all_files(rootdir)
    print(df)

    # match_all_bill(r"data\bill_file_list.xlsx")






