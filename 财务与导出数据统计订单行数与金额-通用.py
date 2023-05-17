# __coding=utf8__
# /** 作者：zengyanghui **/

import sys
import os
import pandas as pd
#显示所有列
pd.set_option('display.max_columns', None)
#显示所有行
pd.set_option('display.max_rows', None)
#设置value的显示长度为100，默认为50
pd.set_option('max_colwidth',200)

import numpy as np

import time
import os.path
import xlrd
import xlwt
import pprint

import tabulate

# import win32api
# import win32ui
# import win32con
#
# import win32com
# from win32com.shell import shell

# # 1表示打开文件对话框
# dlg = win32ui.CreateFileDialog(1)
# # 设置打开文件对话框中的初始显示目录
# dlg.SetOFNInitialDir('E:/Python')
# # 弹出文件选择对话框
# dlg.DoModal()
# # 获取选择的文件名称
# filename = dlg.GetPathName()
# print(filename)


# xx=shell.SHGetPathFromIDList()
# print(xx)




def list_all_files(rootdir, filekey_list):
    if len(filekey_list) > 0:
        filekey_list = filekey_list.replace(",", " ")
        filekey = filekey_list.split(" ")
    else:
        filekey = ''

    _files = []
    list = os.listdir(rootdir)  # 列出文件夹下所有的目录与文件
    for i in range(0, len(list)):
        path = os.path.join(rootdir, list[i])
        if os.path.isdir(path):
            _files.extend(list_all_files(path, filekey_list))
        if os.path.isfile(path):
            if ((path.find("~") < 0)  and  (path.find(".DS_Store") < 0)):  # 带~符号表示临时文件，不读取
                if len(filekey) > 0:
                    for key in filekey:
                        # print(path)
                        filename = "".join(path.split("\\")[-1:])
                        # print("文件名:",filename)

                        key=key.replace("！","!")

                        if key.find("!")>=0:
                            # print("反向选择:",key)
                            if filename.find(key.replace("!","")) >= 0:  # 此文件不要读取
                                # print("{} 不应该包含 {}，所以剔除:".format(filename,key ))
                                pass
                        elif filename.find(key) > 0:  # 只做文件名的过滤
                            _files.append(path)

                else:
                    _files.append(path)

    # print(_files)
    return _files


def read_excel(filename):
    # 原代码，备份：
    # print(filename)
    print(filename)
    if filename.find("xls") > 0:
        df = pd.read_excel(filename, dtype=str)
        df = df.apply(lambda x: x.astype(str).str.replace("¥", "").str.replace(":", "").str.replace("'", "").str.replace('"', ''))
    else:
        if filename.find("拼多多"):
            df = pd.read_csv(filename, dtype=str)
        else:
            df = pd.read_csv(filename, dtype=str, encoding="gb18030")
    print(df.head(1).to_markdown())
    df.drop_duplicates(inplace=True)
    print("定位1")
    if ((filename.find("淘宝") > 0) or (filename.find("TAOBAO") > 0) or (filename.find("天猫") > 0) or (filename.find("TMALL") > 0)):
        plat = "淘宝/天猫"
        if filename.find("财务") >= 0:
            if "订单状态" in df.columns:
                df = df[df["订单状态"].str.contains("交易成功")]
            else:
                dict = {"平台": "", "店铺名称": "", "年度": "", "月份": "", "订单数量": "", "订单金额": ""}
                df = pd.DataFrame(dict, index=[0])
                return df
        else:
            pass
        if "买家实际支付金额" in df.columns:
            df["订单金额"] = df["买家实际支付金额"]
        elif "金额" in df.columns:
            df["订单金额"] = df["金额"]
        elif "销售金额" in df.columns:
            df["订单金额"] = df["销售金额"]
        elif "买家应付货款" in df.columns:
            df["订单金额"] = df["买家应付货款"]
        if "订单创建时间" in df.columns:
            df["开始时间"] = df["订单创建时间"]
        else:
            dict = {"平台": "", "店铺名称": "", "年度": "", "月份": "", "订单数量": "", "订单金额": ""}
            df = pd.DataFrame(dict, index=[0])
            return df
        print("淘宝/天猫")

    elif ((filename.find("京东") > 0) or (filename.find("JD") > 0)):
        plat = "京东"
        if "订单号" in df.columns:
            # df.dropna(subset=["订单号"], inplace=True)
            df = df[~df["订单号"].str.contains("nan")]
            df.drop_duplicates(subset="订单号",inplace=True)
            print(df.tail(5).to_markdown())
        if "应付金额" in df.columns:
            df["订单金额"] = df["应付金额"]
        elif "金额" in df.columns:
            df["订单金额"] = df["金额"]
        elif "销售金额" in df.columns:
            df["订单金额"] = df["销售金额"]
        if "下单时间" in df.columns:
            df["开始时间"] = df["下单时间"]
        elif "时间" in df.columns:
            df["开始时间"] = df["时间"]
        elif "日期" in df.columns:
            df["开始时间"] = df["日期"]
        elif "订单创建时间" in df.columns:
            df["开始时间"] = df["订单创建时间"]
        if "开始时间" in df.columns:
            pass
        else:
            dict = {"平台": "", "店铺名称": "", "年度": "", "月份": "", "订单数量": "", "订单金额": ""}
            df = pd.DataFrame(dict, index=[0])
            return df
        if filename.find("财务") >= 0:
            if "订单状态" in df.columns:
                df = df[df["订单状态"].str.contains("完成")]
            else:
                dict = {"平台": "", "店铺名称": "", "年度": "", "月份": "", "订单数量": "", "订单金额": ""}
                df = pd.DataFrame(dict, index=[0])
                return df
        if "店铺名称" in df.columns:
            pass
        else:
            if filename.find("2019") >= 0:
                df["店铺名称"] = "".join(filename.split(os.sep)[-2:-1])
                df["店铺名称"] = df["店铺名称"].replace("京东", "")
            elif filename.find("2020") >=0:
                shop = "".join(filename.split(os.sep)[-1:])
                df["店铺名称"] = "".join(shop.split("2020")[:1])
                df["店铺名称"] = df["店铺名称"].replace("京东", "")
        if "订单金额" in df.columns:
            pass
        else:
            dict = {"平台": "", "店铺名称": "", "年度": "", "月份": "", "订单数量": "", "订单金额": ""}
            df = pd.DataFrame(dict, index=[0])
            return df
        print("京东")
    elif ((filename.find("拼多多") > 0) or (filename.find("PDD") > 0)):
        plat = "拼多多"
        if "商家实收金额(元)" in df.columns:
            df["订单金额"] = df["商家实收金额(元)"]
        elif "销售金额" in df.columns:
            df["订单金额"] = df["销售金额"]
        if "支付时间" in df.columns:
            df["开始时间"] = df["支付时间"]
        elif "拼单成功时间" in df.columns:
            df["开始时间"] = df["拼单成功时间"]
        elif "订单创建时间" in df.columns:
            df["开始时间"] = df["订单创建时间"]
        if "售后状态" in df.columns:
            df = df[~df["售后状态"].str.contains("退款成功")]
        if "店铺名称" in df.columns:
            pass
        else:
            if filename.find("2019") >= 0:
                df["店铺名称"] = "".join(filename.split(os.sep)[-2:-1])
                df["店铺名称"] = df["店铺名称"].replace("拼多多", "")
            elif filename.find("2020") >=0:
                shop = "".join(filename.split(os.sep)[-1:])
                df["店铺名称"] = "".join(shop.split("2020")[:1])
                df["店铺名称"] = df["店铺名称"].replace("拼多多", "")
        if "拼多多店铺账务明细查询" in df.columns:
            dict = {"平台": "", "店铺名称": "", "年度": "", "月份": "", "订单数量": "", "订单金额": ""}
            df = pd.DataFrame(dict, index=[0])
            return df
        else:
            pass
        df = df[df["开始时间"].str.contains("2019")]
        print("拼多多")
    elif ((filename.find("抖音") > 0) or (filename.find("DY") > 0)):
        plat = "抖音"
        if "实收款（到付按此收费）" in df.columns:
            df["订单金额"] = df["实收款（到付按此收费）"]
        elif "应付金额（到付按此收费）" in df.columns:
            df["订单金额"] = df["应付金额（到付按此收费）"]
        elif "订单应付金额" in df.columns:
            df["订单金额"] = df["订单应付金额"]
        else:
            df["订单金额"] = df["销售金额"]
        if "订单提交时间" in df.columns:
            df["开始时间"] = df["订单提交时间"]
        else:
            df["开始时间"] = df["订单创建时间"]
        if "店铺名称" in df.columns:
            pass
        else:
            file = "".join(filename.split(os.sep)[-1:])
            if file.find("（")>=0:
                df["店铺名称"] = "".join(file.split("（")[:1])
            elif file.find("订单")>=0:
                df["店铺名称"] = "".join(file.split("订单")[:1])
        print("抖音")
    elif ((filename.find("快手") > 0) or (filename.find("KS") > 0)):
        for column_name in df.columns:
            df.rename(columns={column_name: column_name.replace(" ", "").replace("\n", "").strip()}, inplace=True)
        plat = "快手"
        if "实付款(元)" in df.columns:
            df["订单金额"] = df["实付款(元)"]
        elif "实付款" in df.columns:
            df["订单金额"] = df["实付款"]
        elif "销售金额" in df.columns:
            df["订单金额"] = df["销售金额"]
        df["开始时间"] = df["订单创建时间"]
        if "店铺名称" in df.columns:
            pass
        else:
            file = "".join(filename.split(os.sep)[-1:])
            shop = "".join(file.split("（")[:1])
            df["店铺名称"] = shop
        print("快手")
    elif ((filename.find("小红书") > 0) or (filename.find("XHS") > 0)):
        if ((filename.find("小红书") >= 0) & (filename.find("订单") >= 0)):
            print("小红书定位1")
            # try:
            #     filename = filename[:-7] + "账单.xlsx"
            #     df = pd.read_excel(filename, sheet_name="商品销售", dtype=str)
            #     df = df.apply(
            #         lambda x: x.astype(str).str.replace("¥", "").str.replace(":", "").str.replace("'", "").str.replace(
            #             '"', ''))
            # except Exception as e:
            dict = {"平台": "", "店铺名称": "", "年度": "", "月份": "", "订单数量": "", "订单金额": ""}
            df = pd.DataFrame(dict, index=[0])
            print("小红书定位2")
            return df
        elif ((filename.find("XHS") >= 0) & (filename.find("订单") >= 0)):
            if filename.find("2019") >= 0:
                df = df[df["主订单编号"].str.contains("P")]
                print("小红书定位3.1")
            elif filename.find("2020") >= 0:
                print("小红书定位3.2")
                pass
        else:
            try:
            # filename = filename[:-7] + "账单.xlsx"
                df = pd.read_excel(filename, sheet_name="商品销售", dtype=str)
                print(df.head(1).to_markdown())
                df = df.apply(
                    lambda x: x.astype(str).str.replace("¥", "").str.replace(":", "").str.replace("'", "").str.replace('"', ''))
                print("小红书定位4")
            except Exception as e:
                dict = {"平台": "", "店铺名称": "", "年度": "", "月份": "", "订单数量": "", "订单金额": ""}
                df = pd.DataFrame(dict, index=[0])
                print("小红书定位5")
                return df
        plat = "小红书"
        # print("小红书定位4")
        if "实付总额" in df.columns:
            df["订单金额"] = df["实付总额"]
        elif "收入总额" in df.columns:
            df["订单金额"] = df["收入总额"]
        else:
            df["订单金额"] = df["销售金额"]
        if "用户下单时间" in df.columns:
            df["开始时间"] = df["用户下单时间"]
        else:
            df["开始时间"] = df["订单创建时间"]
        if "店铺名称" in df.columns:
            pass
        else:
            df["店铺名称"] = "".join(filename.split(os.sep)[-2:-1])
            df["店铺名称"] = df["店铺名称"].replace("小红书", "")
        print("小红书")
    elif (filename.find("考拉") > 0 or filename.find("KAOLA") > 0):
        plat = "网易考拉"
        if "订单实付金额" in df.columns:
            df["订单金额"] = df["订单实付金额"]
        else:
            df["订单金额"] = df["销售金额"]
        if "下单时间" in df.columns:
            df["开始时间"] = df["下单时间"]
        else:
            df["开始时间"] = df["订单创建时间"]
        if "店铺名称" in df.columns:
            pass
        else:
            if filename.find("2019") >= 0:
                df["店铺名称"] = "".join(filename.split(os.sep)[-2:-1])
                df["店铺名称"] = df["店铺名称"].replace("网易考拉", "")
            elif filename.find("2020") >= 0:
                shop = "".join(filename.split(os.sep)[-1:])
                if shop.find("-") >= 0:
                    df["店铺名称"] = "".join(shop.split("-")[:1])
                    df["店铺名称"] = df["店铺名称"].replace("网易考拉", "")
                else:
                    df["店铺名称"] = "".join(shop.split("2020")[:1])
                    df["店铺名称"] = df["店铺名称"].replace("网易考拉", "")
        print("网易考拉")
    elif (filename.find("有赞") > 0 or filename.find("YZ") > 0):
        plat = "有赞"
        if "订单实付金额" in df.columns:
            df["订单金额"] = df["订单实付金额"]
        elif "应收订单金额" in df.columns:
            df["订单金额"] = df["应收订单金额"]
        else:
            df["订单金额"] = df["销售金额"]
        if "店铺名称" in df.columns:
            pass
        else:
            if filename.find("2019") >= 0:
                df["店铺名称"] = "".join(filename.split(os.sep)[-2:-1])
                df["店铺名称"] = df["店铺名称"].replace("（有赞）订单原表","")
            elif filename.find("2020") >= 0:
                if "归属店铺" in df.columns:
                    df["店铺名称"] = df["归属店铺"]
                else:
                    shop = "".join(filename.split(os.sep)[-1:])
                    df["店铺名称"] = "".join(shop.split("（")[:1])
        df["开始时间"] = df["订单创建时间"]
        print("有赞")
    elif filename.find("唯品会") > 0:
        if "订单号" in df.columns:
            df = df[~df["订单号"].str.contains("-")]
        plat = "唯品会"
        if "客户实际支付金额" in df.columns:
            df["订单金额"] = df["客户实际支付金额"]
        elif "客户应付金额" in df.columns:
            df["订单金额"] = df["客户应付金额"]
        elif "客户实际支付金额（商品金额-商家优惠" in df.columns:
            df["订单金额"] = df["客户实际支付金额（商品金额-商家优惠"]
        if "下单时间" in df.columns:
            df["开始时间"] = df["下单时间"]
        if "品牌" in df.columns:
            df["店铺名称"] = df["品牌"]
        if "订单金额" in df.columns:
            pass
        else:
            print(f"无数据{filename}")
            dict = {"平台": "", "店铺名称": "", "年度": "", "月份": "", "订单数量": "", "订单金额": ""}
            df = pd.DataFrame(dict, index=[0])
            return df
        print("唯品会")
    elif filename.find("金牛") > 0:
        plat = "金牛"
        if "结算金额" in df.columns:
            df["订单金额"] = df["结算金额"]
        elif "订单实付" in df.columns:
            df["订单金额"] = df["订单实付"]
        else:
            pass
        df["开始时间"] = df["下单时间"]
        if "店铺名称" in df.columns:
            pass
        else:
            file = "".join(filename.split(os.sep)[-1:])
            shop = "".join(file.split("（")[:1])
            df["店铺名称"] = shop
        print("金牛电商")
    elif filename.find("百度") > 0:
        plat = "百度小店"
        if "结算金额" in df.columns:
            df["订单金额"] = df["结算金额"]
        else:
            df["订单金额"] = df["总价"]
        df["下单时间"] = df["下单时间"].apply(lambda x:x.replace("年","-").replace("月","-").replace("日","").replace("时",":").replace("分",":").replace("秒",""))
        df["开始时间"] = df["下单时间"]
        if "店铺名称" in df.columns:
            pass
        else:
            file = "".join(filename.split(os.sep)[-1:])
            shop = "".join(file.split("（")[:1])
            df["店铺名称"] = shop
        print("百度小店")
    else:
        dict = {"平台": "", "店铺名称": "", "年度": "", "月份": "", "订单数量": "", "订单金额": ""}
        df = pd.DataFrame(dict, index=[0])
        print("其他平台")
        return df
    for column_name in df.columns:
        df.rename(columns={column_name:column_name.replace(" ","").replace("\n","").strip()},inplace=True)
        if "店铺名称" in df.columns:
            pass
        else:
            if column_name == "店铺":
                df.rename(columns={"店铺名称": "店铺名称"}, inplace=True)
            elif column_name == "店铺号":
                df.rename(columns={"店铺号": "店铺名称"}, inplace=True)
            elif column_name == "卖家名称":
                df.rename(columns={"卖家名称": "店铺名称"}, inplace=True)
            elif column_name == "归属店铺":
                df.rename(columns={"归属店铺": "店铺名称"}, inplace=True)
            else:
                file = "".join(filename.split(os.sep)[-1:])
                shop = "".join(file.split("2019")[:1])
                df["店铺名称"] = shop
                # df["店铺名称"] = "".join(filename.split("/")[-1:])
                # df["店铺名称"] = "".join(df["店铺名称"][0].split("19")[:1])
    print(df.head(1).to_markdown())

    # shop = df["店铺名称"][0]
    # df["店铺名称"] = df["店铺名称"].astype(str)
    # df["店铺名称"] = df["店铺名称"].str.replace("买家未付款",shop).str.replace("订单未关闭",shop).str.replace("退款",shop).str.replace("nan",shop)
    df["平台"] = plat
    df.dropna(subset=["开始时间"],inplace=True)
    df["开始时间"] = df["开始时间"].astype(str).apply(lambda x: x.replace(".", "-"))
    df["开始时间"] = df["开始时间"].astype("datetime64[ns]")
    df["年度"]=df["开始时间"].apply(lambda x: x.year)
    df["月份"] = df["开始时间"].apply(lambda x: x.month)
    df["订单数量"] = 1
    df["订单金额"] = df["订单金额"].astype(float)

    temp_df = df.groupby(["平台", "店铺名称","年度", "月份"]).agg({"订单数量": "sum", "订单金额": "sum"})
    temp_df = pd.DataFrame(temp_df).reset_index()
    print(temp_df.head(5).to_markdown())
    return temp_df



def get_all_files(rootdir, filekey):
    filelist = list_all_files(rootdir, filekey)
    # print(filelist)
    mySeries = pd.Series(filelist)
    df = pd.DataFrame(mySeries)
    df.columns = ["filename"]
    # df["filename"] = df["filename"].apply(lambda x: x.strip().replace(" ", ""))

    df=df[~df["filename"].str.contains("快递")]

    df=df[(~df["filename"].str.contains("账单") & (~df["filename"].str.contains("小红书"))) | df["filename"].str.contains("小红书") ]

    df=df[~df["filename"].str.contains("无")]
    df=df[~df["filename"].str.contains("结算")]
    df=df[~df["filename"].str.contains("刷单")]
    df=df[~df["filename"].str.contains("差异")]
    df=df[~df["filename"].str.contains("支付宝")]
    df=df[~df["filename"].str.contains("商务")]
    df=df[~df["filename"].str.contains("商品")]
    df=df[~df["filename"].str.contains("回款")]
    df=df[~df["filename"].str.contains("总表")]
    df=df[~df["filename"].str.contains("Sku")]
    df=df[~df["filename"].str.contains("不存在")]
    df=df[~df["filename"].str.contains("报表")]
    df=df[~df["filename"].str.contains("账户")]
    df=df[~df["filename"].str.contains("钱包")]
    df=df[~df["filename"].str.contains("其他费用")]
    # df=df[~df["filename"].str.contains("账单")]
    df=df[~df["filename"].str.contains(".zip")]


    # print(df.to_markdown())
    # print("抽查是否还有快递！")
    # print(df[df.filename.str.contains("快递")].to_markdown())
    return df


# def read_all_excel(rootdir, filekey):
#     df_files = get_all_files(rootdir, filekey)
#     for index, file in df_files.iterrows():
#         if 'df' in locals().keys():  # 如果变量已经存在
#             dd = read_excel(file["filename"])
#             dd["filename"] = file["filename"]
#             print("进度表：{}/{}  文件{}，行数{}".format(index + 1, df_files.shape[0], file["filename"], dd.shape[0]))
#             df = df.append(dd)
#         else:
#             df = read_excel(file["filename"])
#             df["filename"] = file["filename"]
#             print("进度表：{}/{}  文件{}，行数{}".format(index + 1, df_files.shape[0], file["filename"], df.shape[0]))
#
#     return df

def get_amount(filename):
    df=pd.read_excel(r"/Users/maclove/PycharmProjects/pythonConda/data/文件分类1.xlsx",sheet_name="Sheet2")
    # df = pd.read_excel(filename)
    # print(df.to_markdown())
    for index ,row in df.iterrows():
        # print(filename)
        # print(index)
        # print(row)
        if filename.find(row["平台"]) >= 0:
            # print(row["平台"])
            try:
                amount_column=row["金额字段"]
                # print("文件名:", filename, " 金额字段为：", amount_column)
                if filename.find("小红书") > 0:
                    tempdb=pd.read_excel(filename,sheet_name="商品销售")
                else:
                    tempdb = pd.read_excel(filename)
                    if filename.find("快手") > 0:
                        tempdb = tempdb.apply(lambda x: x.astype(str).str.replace("¥", ""))
                        if "实付款" in tempdb.columns:
                            tempdb["实付款"] = tempdb["实付款"].astype(float)
                        elif "实付款(元)" in tempdb.columns:
                            tempdb["实付款(元)"] = tempdb["实付款(元)"].astype(float)
                    else:
                        pass
                # print(tempdb.head(1).to_markdown())
                if amount_column.find(",") > 0:
                    # print(amount_column," is in  ",tempdb.columns)
                    amount_columns=amount_column.split(",")
                    for acl in amount_columns:
                        if acl in tempdb.columns:
                            return tempdb[[acl]].sum()
                else:
                    return tempdb[[amount_column]].sum()
            except Exception as e:
                print("没有找到金额字段！", filename)
                return 0
    print("没有找到平台！",filename)
    return 0


def read_all_excel(rootdir, filekey):
    df_files = get_all_files(rootdir, filekey)
    for index, file in df_files.iterrows():
        if 'df' in locals().keys():  # 如果变量已经存在
            dd = read_excel(file["filename"])
            dd["filename"] = file["filename"]
            print("进度表：{}/{}  文件{}，行数{}".format(index + 1, df_files.shape[0], file["filename"], dd.shape[0]))

            df = df.append(dd)
            # hangshu= dd.shape[0]
            # df = df.append( file["filename"],hangshu]  )
            # print(file["filename"],hangshu)

            # print(file["filename"] )
            # amount=get_amount(file["filename"])
            # print(file["filename"], dd.shape[0], amount)

        else:
            df = read_excel(file["filename"])
            df["filename"] = file["filename"]
            # print(file["filename"],  df.shape[0])
            print("进度表：{}/{}  文件{}，行数{}".format(index + 1, df_files.shape[0], file["filename"], df.shape[0]))

    return df


def combine_excel():
    print('订单数据校对逻辑:')
    print('1.财务订单数据需要放在财务数据文件夹下，例如/校对数据/财务数据/...')
    print('2.导出订单数据需要放在导出数据文件夹下，例如/校对数据/导出数据/...')
    print("请输入财务订单和导出订单所在的文件夹：")
    # filedir=""
    # filedir = input()
    # myTuple = shell.SHBrowseForFolder(0, None, "", 64)
    try:
        # path = shell.SHGetPathFromIDList(myTuple[0])
        filedir = input()
    except:
        print("你没有输入任何目录 :(")
        sys.exit()
        return

    # filedir=path.decode('ansi')
    print("你选择的路径是：", filedir)


    global default_dir
    default_dir = filedir

    if len(filedir) == 0:
        print("你没有输入任何目录 :(")
        sys.exit()
        return

    print('请输入要合并的文件名称或者后缀匹配符:（比如都必须包含 序时账 三个字，那么就请输入  "序时账" ， 不输入就表示所有excel都要合并！）')
    filekey = input()

    if len(filedir) == 0:
        print("你没有输入任何关键词 :(")
        filekey = ''
        # sys.exit()
        # return

    print("你希望在'{}'目录下找到所有的包含“{}”文件，然后合并。".format(filedir, filekey))

    # table = read_all_excel(filedir, filekey)
    table = read_all_excel(filedir, filekey)
    # if default_dir.find("财务数据") >= 0:
    #     data = "财务订单"
    #     table["数据来源"] = data
    # else:
    #     data = "导出订单"
    #     table["数据来源"] = data
    table.loc[table["filename"].str.contains("财务数据"), "数据来源"] = "财务订单"
    table.loc[table["filename"].str.contains("导出数据"), "数据来源"] = "导出订单"
    # table["数据来源"] = data

    # table.to_excel("data/财务订单的数量和金额.xlsx",index=False)
    # table.to_excel(default_dir + "/全部订单的数量和金额.xlsx", index=False)
    table1 = table[table["数据来源"].str.contains("财务订单")]
    table2 = table[table["数据来源"].str.contains("导出订单")]
    table1.to_excel(default_dir + "/财务订单的数量和金额.xlsx", index=False)
    table2.to_excel(default_dir + "/导出订单的数量和金额.xlsx", index=False)

    return table1,table2


def groupby_amt():
    # df = pd.read_excel("/Users/maclove/Downloads/2020/2020-不含天猫/2020/财务订单的数量和金额.xlsx")
    # df1,df2 = combine_excel()
    # if default_dir.find("财务数据") >= 0:
    #     data = "财务订单"
    # else:
    #     data = "导出订单"
    # default_dir = r"/Users/maclove/Downloads/2020/2020-不含天猫/2020"
    data1 = "财务订单"
    data2 = "导出订单"
    filename1 = default_dir+"/财务订单的数量和金额.xlsx"
    df1 = pd.read_excel(filename1)
    del df1["filename"]
    df1.drop_duplicates(inplace=True)
    df1.dropna(subset=["订单数量"],inplace=True)
    df1["订单数量"] = df1["订单数量"].astype(int)
    df1["订单金额"] = df1["订单金额"].astype(float)
    df1["店铺名称"] = df1["店铺名称"].str.lower()
    df1["店铺名称"] = df1["店铺名称"].str.replace(" ", "").str.strip()
    df1["店铺名称"] = df1["店铺名称"].astype(str)
    df1["店铺名称"] = df1["店铺名称"].apply(
        lambda x: x.replace("loshi京东自营旗舰店", "loshi旗舰店").
            replace("manuka海外旗舰店", "manukascosmet海外旗舰店").replace("chiara京东旗舰店", "chiarabcaambra海外旗舰店")
            .replace("dentyl京东旗舰店", "dentylactive旗舰店").replace("京东美妆店","麦凯莱美妆专营店").replace(
            "unix优丽氏海外旗舰店", "unix优丽氏海外品牌店").replace("乐丝拼多多店", "乐丝旗舰店").replace("卖家联合海外专卖店", "卖家联合海外专营店")
            .replace("卖家联合全球购（有赞）订单原表", "卖家联合全球购").replace("时尚芭莎美妆店","时尚芭莎美妆店主").replace("莎莎美妆品牌馆","莎莎美妆品牌店")
            .replace("小红书samouaiwoman海外品牌店", "samouraiwoman海外品牌店")
            .replace("惠优乐购", "惠优购企业店").replace("芭莎美妆馆","时尚芭莎美妆店主").replace("unix优丽氏品牌店-","unix优丽氏品牌店"))
    df1["店铺名称"] = df1["店铺名称"].apply(
        lambda x: x.replace("lawrence_zhao1", "时尚芭莎美妆店主").replace("dentyl京东自营店", "dentylactive自营旗舰店")
            .replace("dentyl拼购旗舰店", "dentylactive拼购旗舰店").replace("dicora京东海外旗舰店", "dicoraurbanfit海外旗舰店")
            .replace("ambra旗舰店", "AMBRA京东旗舰店").replace("manuka京东海外店原表", "manukascosmet海外旗舰店").replace(
            "manuka京东海外店", "manukascosmet海外旗舰店").replace("ultardex京东旗舰店", "ultradex旗舰店").replace("博滴播地艾专卖店",
            "博滴品牌专卖店").replace("播地艾个护专营店", "睿旗个护店").replace("奈萃拉樱岚专卖店", "allnaturaladvice樱岚专卖店")
            .replace("奈萃拉旗舰店", "bodyaid个护旗舰店").replace("羊羊的小铺", "二十四小时七天个护专营店").replace("mades个护", "博滴个护旗舰店")
            .replace("精酿商贸", "博滴旗舰店").replace("可瘾家清专营店", "可瘾个护家清专营店").replace("morei多瑞专卖店", "来一泡多瑞专卖店")
            .replace("玫德丝配颜师专卖店", "植之璨配颜师专卖店").replace("睿旗口腔专营店", "睿旗个护家清专营店").replace("mega店铺", "航星个护专营店")
            .replace("魔湾小店", "魔湾个护家清专营店").replace("mega小店", "魔湾游戏个护专营店").replace("麦凯莱个护专营店", "麦凯莱个护家清专营店")
            .replace("若蘅美妆店", "若蘅旗舰店").replace("unix专卖店", "unix优丽氏品牌店").replace("mega精选小店", "mega精选小铺")
            .replace("可瘾个护家清专营店-202009-", "可瘾个护家清专营店").replace("magicsymbol美妆店", "magicsymbol旗舰店").replace("卖家联合海外旗舰店", "卖家联合海外专营店")
            .replace("baza", "时尚芭莎美妆店主"))
    df1["店铺名称"] = df1.apply(lambda x:"loshi旗舰店" if ((x["平台"] == "抖音") & (x["店铺名称"] == "若蘅旗舰店")) else x["店铺名称"],axis=1)

    df1["店铺名称"] = df1["店铺名称"].str.replace("京东", "").str.replace("拼多多", "").str.replace("网易考拉", "").str.replace("小红书", "").str.replace("原表", "")

    group_df1 = df1.groupby(["平台","店铺名称","年度","月份"]).agg({"订单数量":"sum","订单金额":"sum"})
    # group_df["数据来源"] = "导出数据"
    group_df1["数据来源"] = "财务订单"
    group_df1 = pd.DataFrame(group_df1).reset_index()
    # group_df1.to_excel("data/整理后的财务订单的数量和金额.xlsx")
    group_df1.to_excel(default_dir + "/整理后的财务订单的数量和金额.xlsx", index=False)

    filename2 = default_dir + "/导出订单的数量和金额.xlsx"
    df2 = pd.read_excel(filename2)
    del df2["filename"]
    df2.drop_duplicates(inplace=True)
    df2.dropna(subset=["订单数量"], inplace=True)
    df2["订单数量"] = df2["订单数量"].astype(int)
    df2["订单金额"] = df2["订单金额"].astype(float)
    df2["店铺名称"] = df2["店铺名称"].str.lower()
    df2["店铺名称"] = df2["店铺名称"].str.replace(" ", "").str.strip()
    df2["店铺名称"] = df2["店铺名称"].replace("loshi自营旗舰店", "loshi旗舰店").replace("loshi京东自营旗舰店", "loshi旗舰店")
    df2["店铺名称"] = df2["店铺名称"].str.replace("京东", "").str.replace("拼多多", "").str.replace("网易考拉", "").str.replace("小红书", "")
    group_df2 = df2.groupby(["平台", "店铺名称", "年度", "月份"]).agg({"订单数量": "sum", "订单金额": "sum"})
    # group_df2["数据来源"] = "导出数据"
    group_df2["数据来源"] = "导出数据"
    group_df2 = pd.DataFrame(group_df2).reset_index()
    # group_df2.to_excel("data/整理后的财务订单的数量和金额.xlsx")
    group_df2.to_excel(default_dir + "/整理后的导出订单的数量和金额.xlsx", index=False)

def math_file():
    # default_dir = r"/Users/maclove/Downloads/2020/2020-不含天猫/2020"
    df1 = pd.read_excel(default_dir + "/整理后的财务订单的数量和金额.xlsx",keep_default_na=False)
    df2 = pd.read_excel(default_dir + "/整理后的导出订单的数量和金额.xlsx",keep_default_na=False)

    # df1["店铺名称"] = df1["店铺名称"].str.lower()
    # df1["店铺名称"] = df1["店铺名称"].str.replace(" ", "").str.strip()
    # df2["店铺名称"] = df2["店铺名称"].str.lower()
    # df2["店铺名称"] = df2["店铺名称"].str.replace(" ", "").str.strip()
    # df1["店铺名称"] = df1["店铺名称"].apply(
    #     lambda x: x.replace("ambra京东海外旗舰店", "chiarabcaambra海外旗舰店").replace("loshi京东自营旗舰店","loshi旗舰店").
    #         replace("manuka海外旗舰店", "manukascosmet海外旗舰店").replace("chiara京东旗舰店", "chiarabcaambra海外旗舰店")
    #         .replace("dentyl京东旗舰店","dentylactive旗舰店").replace("smilelab京东旗舰店", "smilelab旗舰店").replace("京东美妆店",
    #         "麦凯莱美妆专营店").replace("unix优丽氏海外旗舰店","unix优丽氏海外品牌店").replace("乐丝拼多多店","乐丝旗舰店")
    #         .replace("卖家联合拼多多", "卖家联合海外专营店").replace("卖家联合全球购（有赞）", "卖家联合全球购").replace("时尚芭莎美妆店",
    #         "时尚芭莎美妆店主").replace("莎莎美妆品牌馆", "莎莎美妆品牌店"))
    # df2["店铺名称"] = df1["店铺名称"].apply(
    #     lambda x: x.replace("ambra京东海外旗舰店", "chiarabcaambra海外旗舰店").replace("loshi京东自营旗舰店", "loshi旗舰店").
    #         replace("manuka海外旗舰店", "manukascosmet海外旗舰店").replace("chiara京东旗舰店", "chiarabcaambra海外旗舰店")
    #         .replace("dentyl京东旗舰店", "dentylactive旗舰店").replace("smilelab京东旗舰店", "smilelab旗舰店").replace("京东美妆店",
    #                                                                                                    "麦凯莱美妆专营店").replace(
    #         "unix优丽氏海外旗舰店", "unix优丽氏海外品牌店").replace("乐丝拼多多店", "乐丝旗舰店")
    #         .replace("卖家联合拼多多", "卖家联合海外专营店").replace("卖家联合全球购（有赞）", "卖家联合全球购").replace("时尚芭莎美妆店",
    #                                                                                    "时尚芭莎美妆店主").replace("莎莎美妆品牌馆",
    #                                                                                                        "莎莎美妆品牌店"))
    # df1["店铺名称"] = df1["店铺名称"].str.replace("京东","").str.replace("拼多多","").str.replace("网易考拉","")
    # df2["店铺名称"] = df2["店铺名称"].str.replace("京东","").str.replace("拼多多","").str.replace("网易考拉","")
    df1.rename(columns={"订单数量":"财务-订单数量","订单金额":"财务-订单金额"},inplace=True)
    df2.rename(columns={"订单数量":"导出-订单数量","订单金额":"导出-订单金额"},inplace=True)
    del df1["数据来源"]
    del df2["数据来源"]
    df = pd.merge(df1,df2,how="outer",on=["平台","店铺名称","年度","月份"])
    # df["数量差异（财务/订单）"] = df.apply(lambda  x:  (1-round((x["财务-订单数量"] / x["导出-订单数量"]),2))*100  ,axis=1 )
    # df["数量差异（财务/订单）"] = df.apply(lambda  x: "{}".format((1-round((x["财务-订单数量"] / x["导出-订单数量"]),2))*100),axis=1 )
    # df["数量差异（财务/订单）"] = df.apply(lambda  x: "{:0>2d}%".format((1-round((x["财务-订单数量"] / x["导出-订单数量"]),2))*100) ,axis=1 )

    # df["金额差异（财务/订单）"] = df.apply(lambda  x: "{}".format((1-round((x["财务-订单金额"] / x["导出-订单金额"]),2))*100) ,axis=1 )
    df["数量差异（财务/订单）"] = round((1-(df["财务-订单数量"] / df["导出-订单数量"])), 6)
    df["金额差异（财务/订单）"] = round((1-(df["财务-订单金额"] / df["导出-订单金额"])), 6)

    # df["数量差异（财务/订单）"]=df["数量差异（财务/订单）"].apply(lambda x:  "" if x.find("nan")>=0  else x )
    # df["金额差异（财务/订单）"]=df["金额差异（财务/订单）"].apply(lambda x:  "" if x.find("nan")>=0  else x )


    print(df.to_markdown())
    df.to_excel(default_dir + "/财务和导出订单的数量和金额差距.xlsx")

if __name__ == "__main__":
    print("开始:", time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()))
    # print("abc".find("x"))

    combine_excel()
    groupby_amt()
    math_file()

    print("ok")