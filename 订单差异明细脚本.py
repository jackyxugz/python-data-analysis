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
        df = pd.read_csv(filename, dtype=str, encoding="gb18030")
    print(df.head(1).to_markdown())
    df.drop_duplicates(inplace=True)
    print("定位1")
    if ((filename.find("淘宝") > 0) or (filename.find("TAOBAO") > 0) or (filename.find("天猫") > 0) or (filename.find("TMALL") > 0)):
        plat = "淘宝/天猫"
        if ((filename.find("TAOBAO") > 0) or (filename.find("TMALL") > 0)):
            df = df[df["店铺名称"].str.contains("dentylactive旗舰店|loshi旗舰店|manukascosmet海外旗舰店|smilelab旗舰店|unix旗舰店|unix电器海外旗舰店|前男友美妆|惠优购官方旗舰店|时尚芭莎美妆店主|米兰站美妆|莎莎美妆品牌店|麦凯莱品牌自营店")]
        else:
            pass
        if "订单编号" in df.columns:
            pass
        elif "订单号" in df.columns:
            df["订单编号"] = df["订单号"]
        elif "主订单编号" in df.columns:
            df["订单编号"] = df["主订单编号"]
        if "买家实际支付金额" in df.columns:
            df["订单金额"] = df["买家实际支付金额"]
        elif "金额" in df.columns:
            df["订单金额"] = df["金额"]
        elif "销售金额" in df.columns:
            df["订单金额"] = df["销售金额"]
        if "订单创建时间" in df.columns:
            df["开始时间"] = df["订单创建时间"]
        else:
            dict = {"平台": "", "店铺名称": "", "开始时间": "", "订单编号": "", "订单金额": ""}
            df = pd.DataFrame(dict, index=[0])
            return df
        print("淘宝/天猫")

    elif ((filename.find("京东") > 0) or (filename.find("JD") > 0)):
        plat = "京东"
        if filename.find("JD") > 0:
            df = df[df["店铺名称"].str.contains("优丽氏旗舰店")]
        else:
            pass
        if "订单编号" in df.columns:
            pass
        elif "订单号" in df.columns:
            df["订单编号"] = df["订单号"]
        elif "主订单编号" in df.columns:
            df["订单编号"] = df["主订单编号"]
        df = df[~df["订单编号"].str.contains("nan")]
        if filename.find("财务数据") >= 0:
            df.drop_duplicates(subset="订单编号",inplace=True)
        else:
            pass
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
        if "店铺名称" in df.columns:
            pass
        else:
            df["店铺名称"] = "".join(filename.split(os.sep)[-2:-1])
            df["店铺名称"] = df["店铺名称"].replace("京东", "")
        if "订单金额" in df.columns:
            pass
        else:
            dict = {"平台": "", "店铺名称": "", "开始时间": "", "订单编号": "", "订单金额": ""}
            df = pd.DataFrame(dict, index=[0])
            return df
        print("京东")
    elif ((filename.find("拼多多") > 0) or (filename.find("PDD") > 0)):
        plat = "拼多多"
        if filename.find("PDD") > 0:
            df = df[df["店铺名称"].str.contains("乐丝旗舰店|卖家联合海外专营店|麦凯莱美妆专营店")]
        else:
            pass
        if "订单编号" in df.columns:
            pass
        elif "订单号" in df.columns:
            df["订单编号"] = df["订单号"]
        elif "主订单编号" in df.columns:
            df["订单编号"] = df["主订单编号"]
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
        if "店铺名称" in df.columns:
            pass
        else:
            df["店铺名称"] = "".join(filename.split(os.sep)[-2:-1])
            df["店铺名称"] = df["店铺名称"].replace("拼多多", "")
        if "拼多多店铺账务明细查询" in df.columns:
            dict = {"平台": "", "店铺名称": "", "开始时间": "", "订单编号": "", "订单金额": ""}
            df = pd.DataFrame(dict, index=[0])
            return df
        else:
            pass
        # df = df[df["开始时间"].str.contains("2019")]
        print("拼多多")
    # elif ((filename.find("抖音") > 0) or (filename.find("DY") > 0)):
    #     plat = "抖音"
    #     if "实收款（到付按此收费）" in df.columns:
    #         df["订单金额"] = df["实收款（到付按此收费）"]
    #     elif "应付金额（到付按此收费）" in df.columns:
    #         df["订单金额"] = df["应付金额（到付按此收费）"]
    #     else:
    #         df["订单金额"] = df["销售金额"]
    #     if "订单提交时间" in df.columns:
    #         df["开始时间"] = df["订单提交时间"]
    #     else:
    #         df["开始时间"] = df["订单创建时间"]
    #     print("抖音")
    # elif ((filename.find("快手") > 0) or (filename.find("KS") > 0)):
    #     plat = "快手"
    #     if "实付款(元)" in df.columns:
    #         df["订单金额"] = df["实付款(元)"]
    #     elif "实付款" in df.columns:
    #         df["订单金额"] = df["实付款"]
    #     else:
    #         df["订单金额"] = df["销售金额"]
    #     df["开始时间"] = df["订单创建时间"]
    #     print("快手")
    elif ((filename.find("小红书") > 0) or (filename.find("XHS") > 0)):
        if ((filename.find("小红书") >= 0) & (filename.find("订单") >= 0)):
            # print("小红书定位1")
            dict = {"平台": "", "店铺名称": "", "开始时间": "", "订单编号": "", "订单金额": ""}
            df = pd.DataFrame(dict, index=[0])
            # print("小红书定位3")
            return df
        elif ((filename.find("XHS") >= 0) & (filename.find("订单") >= 0)):
            df = df[df["主订单编号"].str.contains("P")]
            df["订单编号"] = df["主订单编号"]
        else:
            try:
            # filename = filename[:-7] + "账单.xlsx"
                df = pd.read_excel(filename, sheet_name="商品销售", dtype=str)
                print(df.head(1).to_markdown())
                df = df.apply(
                    lambda x: x.astype(str).str.replace("¥", "").str.replace(":", "").str.replace("'", "").str.replace('"', ''))
                # print("小红书定位2")
            except Exception as e:
                dict = {"平台": "", "店铺名称": "", "开始时间": "", "订单编号": "", "订单金额": ""}
                df = pd.DataFrame(dict, index=[0])
                # print("小红书定位3")
                return df
        plat = "小红书"
        # print("小红书定位4")
        if "订单编号" in df.columns:
            pass
        elif "订单号" in df.columns:
            df["订单编号"] = df["订单号"]
        elif "主订单编号" in df.columns:
            df["订单编号"] = df["主订单编号"]
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
        if "订单编号" in df.columns:
            pass
        elif "订单号" in df.columns:
            df["订单编号"] = df["订单号"]
        elif "主订单编号" in df.columns:
            df["订单编号"] = df["主订单编号"]
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
            df["店铺名称"] = "".join(filename.split(os.sep)[-2:-1])
            df["店铺名称"] = df["店铺名称"].replace("网易考拉", "")
        print("网易考拉")
    elif (filename.find("有赞") > 0 or filename.find("YZ") > 0):
        plat = "有赞"
        if "订单编号" in df.columns:
            pass
        elif "订单号" in df.columns:
            df["订单编号"] = df["订单号"]
        elif "主订单编号" in df.columns:
            df["订单编号"] = df["主订单编号"]
        if "订单实付金额" in df.columns:
            df["订单金额"] = df["订单实付金额"]
        elif "应收订单金额" in df.columns:
            df["订单金额"] = df["应收订单金额"]
        else:
            df["订单金额"] = df["销售金额"]
        if "店铺名称" in df.columns:
            pass
        else:
            df["店铺名称"] = "".join(filename.split(os.sep)[-2:-1])
            df["店铺名称"] = df["店铺名称"].replace("（有赞）订单原表","")
        df["开始时间"] = df["订单创建时间"]
        print("有赞")
    # elif filename.find("唯品会") > 0:
    #     if "订单号" in df.columns:
    #         df = df[~df["订单号"].str.contains("-")]
    #     plat = "唯品会"
    #     if "客户实际支付金额" in df.columns:
    #         df["订单金额"] = df["客户实际支付金额"]
    #     elif "客户应付金额" in df.columns:
    #         df["订单金额"] = df["客户应付金额"]
    #     elif "客户实际支付金额（商品金额-商家优惠" in df.columns:
    #         df["订单金额"] = df["客户实际支付金额（商品金额-商家优惠"]
    #     if "下单时间" in df.columns:
    #         df["开始时间"] = df["下单时间"]
    #     if "订单金额" in df.columns:
    #         pass
    #     else:
    #         print(f"无数据{filename}")
    #         dict = {"平台": "", "店铺名称": "", "年度": "", "月份": "", "订单数量": "", "订单金额": ""}
    #         df = pd.DataFrame(dict, index=[0])
    #         return df
    #     print("唯品会")
    # elif filename.find("金牛") > 0:
    #     plat = "金牛电商"
    #     if "结算金额" in df.columns:
    #         df["订单金额"] = df["结算金额"]
    #     elif "订单实付" in df.columns:
    #         df["订单金额"] = df["订单实付"]
    #     else:
    #         pass
    #     df["开始时间"] = df["下单时间"]
    #     print("金牛电商")
    # elif filename.find("百度") > 0:
    #     plat = "金牛电商"
    #     if "结算金额" in df.columns:
    #         df["订单金额"] = df["结算金额"]
    #     else:
    #         df["订单金额"] = df["总价"]
    #     df["开始时间"] = df["下单时间"]
    #     print("百度")
    else:
        dict = {"平台": "", "店铺名称": "", "开始时间": "", "订单编号": "", "订单金额": ""}
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
        if "店铺名称" in df.columns:
            pass
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
    # df.dropna(subset=["开始时间"],inplace=True)
    df["开始时间"] = df["开始时间"].astype(str).apply(lambda x: x.replace(".", "-"))
    df = df[df["开始时间"].str.contains("2019")]
    df["开始时间"] = df["开始时间"].astype("datetime64[ns]")
    # df["年度"]=df["开始时间"].apply(lambda x: x.year)
    # df["月份"] = df["开始时间"].apply(lambda x: x.month)
    # df["订单数量"] = 1
    df["订单金额"] = df["订单金额"].astype(float)
    df = df[["平台","店铺名称","订单编号","开始时间","订单金额"]]
    print(df.head(5).to_markdown())

    # temp_df = df.groupby(["平台", "店铺名称","年度", "月份"]).agg({"订单数量": "sum", "订单金额": "sum"})
    # temp_df = pd.DataFrame(temp_df).reset_index()
    # print(temp_df.head(5).to_markdown())
    return df


def get_all_files(rootdir, filekey):
    filelist = list_all_files(rootdir, filekey)
    # print(filelist)
    mySeries = pd.Series(filelist)
    df = pd.DataFrame(mySeries)
    df.columns = ["filename"]
    # df["filename"] = df["filename"].apply(lambda x: x.strip().replace(" ", ""))

    df=df[~df["filename"].str.contains("快递")]

    df=df[ (~df["filename"].str.contains("账单") & (~df["filename"].str.contains("小红书"))) | df["filename"].str.contains("小红书") ]

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
    table1.to_excel(default_dir + "/财务订单的明细.xlsx", index=False)
    table2.to_excel(default_dir + "/导出订单的明细.xlsx", index=False)

    return table1,table2


def groupby_amt():
    default_dir = r"/Users/maclove/Downloads/2019"
    filename1 = default_dir+"/财务订单的明细.xlsx"
    df1 = pd.read_excel(filename1)
    del df1["filename"]
    df1.drop_duplicates(inplace=True)
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
            .replace("小红书samouaiwoman海外品牌店", "samouraiwoman海外品牌店").replace("dentylactive官方旗舰店","dentylactive旗舰店")
            .replace("惠优乐购", "惠优购企业店").replace("芭莎美妆馆","时尚芭莎美妆店主"))
    df1["店铺名称"] = df1["店铺名称"].str.replace("京东", "").str.replace("拼多多", "").str.replace("网易考拉", "").str.replace("小红书", "")
    df1.drop_duplicates(inplace=True)
    df1.to_excel(default_dir + "/整理后的财务订单的明细.xlsx", index=False)

    filename2 = default_dir + "/导出订单的明细.xlsx"
    df2 = pd.read_excel(filename2)
    del df2["filename"]
    # df2.drop_duplicates(inplace=True)
    df2["订单金额"] = df2["订单金额"].astype(float)
    df2["店铺名称"] = df2["店铺名称"].str.lower()
    df2["店铺名称"] = df2["店铺名称"].str.replace(" ", "").str.strip()
    df2["店铺名称"] = df2["店铺名称"].replace("loshi自营旗舰店", "loshi旗舰店").replace("loshi京东自营旗舰店", "loshi旗舰店")
    df2["店铺名称"] = df2["店铺名称"].str.replace("京东", "").str.replace("拼多多", "").str.replace("网易考拉", "").str.replace("小红书", "")
    print(df2.head(1).to_markdown())
    # print(df2[df2["订单编号"].str.contains("100378755122")])
    df3 = df2[["平台","店铺名称","订单编号","开始时间","数据来源"]]
    df3.drop_duplicates(["订单编号"],inplace=True)
    group_df2 = df2.groupby(["订单编号"]).agg({"订单金额":"sum"})
    print(group_df2.head(1).to_markdown())
    group_df2 = pd.merge(df3,group_df2,how="left",on="订单编号")
    print(group_df2.head(1).to_markdown())
    # group_df2["数据来源"] = "导出订单"
    group_df2 = group_df2[["平台","店铺名称","订单编号","开始时间","订单金额","数据来源"]]
    print(group_df2.head(1).to_markdown())
    # print(group_df2[group_df2["订单编号"].str.contains("100378755122")])
    group_df2.to_excel(default_dir + "/整理后的导出订单的明细.xlsx", index=False)

def math_file():
    default_dir = r"/Users/maclove/Downloads/2019"
    df1 = pd.read_excel(default_dir + "/整理后的财务订单的明细.xlsx",keep_default_na=False)
    df2 = pd.read_excel(default_dir + "/整理后的导出订单的明细.xlsx",keep_default_na=False)

    df1.rename(columns={"开始时间":"财务-订单时间","订单金额":"财务-订单金额"},inplace=True)
    df2.rename(columns={"开始时间":"导出-订单时间","订单金额":"导出-订单金额"},inplace=True)
    del df1["数据来源"]
    del df2["数据来源"]
    df = pd.merge(df1,df2,how="outer",on=["平台","店铺名称","订单编号"])
    # df["数量差异（财务/订单）"] = df.apply(lambda  x:  (1-round((x["财务-订单数量"] / x["导出-订单数量"]),2))*100  ,axis=1 )
    # df["数量差异（财务/订单）"] = df.apply(lambda  x: "{}".format((1-round((x["财务-订单数量"] / x["导出-订单数量"]),2))*100),axis=1 )
    # df["数量差异（财务/订单）"] = df.apply(lambda  x: "{:0>2d}%".format((1-round((x["财务-订单数量"] / x["导出-订单数量"]),2))*100) ,axis=1 )

    # df["金额差异（财务/订单）"] = df.apply(lambda  x: "{}".format((1-round((x["财务-订单金额"] / x["导出-订单金额"]),2))*100) ,axis=1 )
    # df["数量差异（财务/订单）"] = round((1-(df["财务-订单数量"] / df["导出-订单数量"])), 6)
    # df["金额差异（财务/订单）"] = round((1-(df["财务-订单金额"] / df["导出-订单金额"])), 6)

    # df["数量差异（财务/订单）"]=df["数量差异（财务/订单）"].apply(lambda x:  "" if x.find("nan")>=0  else x )
    # df["金额差异（财务/订单）"]=df["金额差异（财务/订单）"].apply(lambda x:  "" if x.find("nan")>=0  else x )



    print(df.tail(10).to_markdown())
    df.to_excel(default_dir + "/财务和导出订单的明细合并.xlsx")

if __name__ == "__main__":
    print("开始:", time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()))
    # print("abc".find("x"))

    combine_excel()
    groupby_amt()
    math_file()

    print("ok")