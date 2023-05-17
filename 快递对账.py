# __coding=utf8__
# /** 作者：zengyanghui **/

import pandas as pd
import numpy as np
from collections import Counter
import os

# import Tkinter
# import win32api
# import win32ui
# import win32con
# from win32com.shell import shell
import tabulate
import math
import xlsxwriter
# import win32com
# from win32com.shell import shell


# def get_file():
#     # print("请确认是否有需要合并的电商数据，输入“Y”或者“N”,输入其他则跳过处理！")
#     ms=win32api.MessageBox(0,"请选择，输入Y合并多个电商文件，输入N选择单个电商文件！","提醒",win32con.MB_YESNO)
#     # ms = input()
#     if ms == 6:
#         print("请输入需要合并的文件夹路径")
#         # dir_path = "data/快递对账/7月份电商数据汇总-包装部"
#         dir_path = input()
#         # print("请输入处理后的文件保存路径：")
#         # result_path = input()
#         filenames = os.listdir(dir_path)
#         index = 0
#         dfs = []
#         for name in filenames:
#             print(index)
#             print("filename:", name)
#             dfs.append(pd.read_excel(os.path.join(dir_path, name),dtype=str))
#             index += 1
#         df1 = pd.concat(dfs)
#         df1 = df1[["快递单号", "店铺名称", "收件人", "订单编号"]].copy()
#         df1.rename(columns={"快递单号": "express_id", "店铺名称": "shop", "收件人": "name", "订单编号": "order_id"}, inplace=True)
#         df1["express_id"] = df1["express_id"].astype(str)
#         # df1 = df1.drop_duplicates(["express_id"], keep='first')
#         print(len(df1))
#         print(df1.head(5).to_markdown())
#         # df1.to_csv(result_path + "/电商数据汇总.csv",index=False)
#         df1.to_pickle("data/快递对账/dianshang.pkl")
#         return df1
#
#     else:
#         # print("请确认是否有需要处理的数据：1.选择文件；2.跳过处理")
#         # choice = input()
#         choice = win32api.MessageBox(0, "输入Y选择单个电商文件，输入N跳过电商文件处理！", "提醒", win32con.MB_YESNO)
#         if int(choice) == 6:
#             print("请选择需要处理的电商数据文件（包含完整路径）：")
#             file_path = open_file()
#             # open_file()
#             # print("请输入处理后的文件保存路径：")
#             # result_path = input()
#             # print("请选择你的文件类型：1.xlsx(xls)；2.csv")
#             # type = input()
#             type = win32api.MessageBox(0, "请确认你的文件类型,xlsx(xls)选是，csv选否！", "提醒", win32con.MB_YESNO)
#             if int(type) == 6:
#                 df = pd.read_excel(file_path)
#             else:
#                 df = pd.read_csv(file_path, encoding="gbk")
#             df1 = df[["快递单号", "店铺名称", "收件人", "订单编号"]].copy()
#             df1.rename(columns={"快递单号": "express_id", "店铺名称": "shop", "收件人": "name", "订单编号": "order_id"}, inplace=True)
#             # df1 = df1.drop_duplicates(["express_id"], keep='first')
#             print(len(df1))
#             print(df1.head(5).to_markdown())
#             # df1.to_csv(result_path + "/电商数据汇总.csv", index=False)
#             return df1
#         else:
#             print("跳过文件处理步骤！")
#             pass
#     else:
#         print("跳过文件合并步骤！")
#         pass


def get_file1():
    # ms=win32api.MessageBox(0,"请选择，输入Y合并多个erp聚水潭文件，输入N选择单个erp聚水潭文件！","提醒",win32con.MB_YESNO)
    # if ms == 6:
    print("请把销售主题分析的解压文件均放到一个目录下，并输入此目录路径：")
    # dir_path = input()
    dir_path = r"D:\沙井\快递单校对\202201\匹配数据\聚水潭"
    filenames = os.listdir(dir_path)
    index = 0
    dfs = []
    print(index)
    for name in filenames:
        print("filename:", name)
        dfs.append(pd.read_csv(os.path.join(dir_path, name), dtype=str, encoding="gb18030"))
        index += 1
    df = pd.concat(dfs)
    # df2 = df[["店铺", "收件人姓名", "快递单号", "线上订单号"]].copy()
    # df2.rename(columns={"店铺": "shop", "收件人姓名": "name", "快递单号": "express_id", "线上订单号": "order_id"}, inplace=True)
    df2 = df[["店铺", "快递单号", "线上订单号"]].copy()
    df2.rename(columns={"店铺": "shop", "快递单号": "express_id", "线上订单号": "order_id"}, inplace=True)
    df2["express_id"] = df2["express_id"].astype(str)
    df2["express_id"] = df2["express_id"].str.replace(" ", "").str.replace(",", "").str.replace("\n", "").str.replace("@", "").str.strip()
    # df2["express_id"] = df2["express_id"].astype(str)
    return df2

def get_file2():
    # choice = win32api.MessageBox(0, "输入Y选择单个打印信息文件，输入N跳过打印信息文件处理！", "提醒", win32con.MB_YESNO)
    # if int(choice) == 6:
    print("请输入打印信息的文件路径，包含文件名，例如C:/快递单校对/打印记录.xls")
    # file_path = open_file()
    # file_path = input()
    file_path = r"D:\沙井\快递单校对\202201\匹配数据\打印记录.xls"
    df = pd.read_excel(file_path,dtype=str)
    if "店铺名" in df.columns:
        pass
    elif "店铺名称" in df.columns:
        df.rename(columns={"店铺名称":"店铺名"},inplace=True)
    df3 = df[["快递单号", "发件人", "收件人", "订单编号"]].copy()
    df3.rename(columns={"快递单号": "express_id", "发件人": "shop", "收件人": "name", "订单编号": "order_id"}, inplace=True)
    df3["express_id"] = df3["express_id"].astype(str)
    print(len(df3))
    print(df3.head(5).to_markdown())
    return df3
    # else:
    #     print("跳过文件处理步骤！")
    #     pass

def get_file3():
    # ms=win32api.MessageBox(0,"请选择，输入Y合并多个快递信息文件，输入N选择单个快递文件！）","提醒",win32con.MB_YESNO)
    # if ms == 6:
    print("请把快递信息的文件均放到一个目录下，并输入此目录路径：")
    # dir_path = input()
    dir_path = r"D:\沙井\快递单校对\202201\1月账单"
    filenames = os.listdir(dir_path)
    index = 0
    dfs = []
    print(index)
    for name in filenames:
        print("filename:", name)
        if name.find("~") >= 0:
            pass
        elif name.find("DS") >= 0:
            pass
        elif name.find("中通快递对账单") >= 0:
            temp_df = pd.read_excel(os.path.join(dir_path, name), sheet_name=None, dtype=str)
            sheet_list = list(temp_df)
            print(sheet_list)
            for sheet in sheet_list:
                if sheet.find("对账单") >= 0:
                    print(sheet)
                    sheet_name = sheet
                    break
                elif sheet.find("日明细") >= 0:
                    # print("定位1")
                    print(sheet)
                    sheet_name = sheet
                    break
            print(sheet_name)
            temp_df = pd.read_excel(os.path.join(dir_path, name), sheet_name=sheet_name, dtype=str, usecols=["运单号"])
        elif name.find("圆通") >= 0:
            temp_df = pd.read_excel(os.path.join(dir_path, name), sheet_name="快递", skiprows=1, dtype=str, usecols=["运单号"])
        # elif name.find("圆通") >= 0:
        #     temp_df = pd.read_excel(os.path.join(dir_path, name), sheet_name="快递", skiprows=1, dtype=str, usecols=["运单号"])
        # elif name.find("韵达") >= 0:
        #     temp_df = pd.read_excel(os.path.join(dir_path, name), sheet_name="对账单", skiprows=1, dtype=str, usecols=["运单号码"])
        #     temp_df.rename(columns={"运单号码":"运单号"},inplace=True)
        elif name.find("韵达") >= 0:
            temp_df = pd.read_excel(os.path.join(dir_path, name), sheet_name=None, dtype=str)
            sheet_list = list(temp_df)
            print(sheet_list)
            for sheet in sheet_list:
                if sheet.find("对账单") >= 0:
                    print(sheet)
                    sheet_name = sheet
                    break
                elif sheet.find("日") >= 0:
                    # print("定位1")
                    print(sheet)
                    sheet_name = sheet
                    break
            print(sheet_name)
            temp_df = pd.read_excel(os.path.join(dir_path, name), sheet_name=sheet_name, skiprows=1, dtype=str, usecols=["运单号码"])
            temp_df.rename(columns={"运单号码":"运单号"},inplace=True)
        elif name.find("京东") >= 0:
            temp_df = pd.read_excel(os.path.join(dir_path, name), sheet_name="对账单", dtype=str, usecols=["运单号"])
        else:
            temp_df=pd.read_excel(os.path.join(dir_path, name), sheet_name="对账单", skiprows=1, dtype=str, usecols=["运单号"])
        temp_df["company"]=name
        if name.find("中通")>0:
            temp_df["company"] = "中通"
        elif name.find("申通")>0:
            temp_df["company"] = "申通"
        elif name.find("邮政") > 0:
            temp_df["company"] = "邮政"
        elif name.find("圆通") > 0:
            temp_df["company"] = "圆通"
        elif name.find("京东") > 0:
            temp_df["company"] = "京东"
        elif name.find("韵达") > 0:
            temp_df["company"] = "韵达"
        dfs.append(temp_df)
        index += 1
    df = pd.concat(dfs)
    df4 = df[["company","运单号"]].copy()
    df4.rename(columns={"运单号": "express_id"}, inplace=True)
    # ms = win32api.MessageBox(0, "输入Y选择单个云仓快递文件，输入N跳过云仓快递文件处理！", "提醒", win32con.MB_YESNO)
    # print("请问是否有云仓的快递信息，如果有请输入1，输入其他则跳过")
    # ms = input()
    # if ms == 1:
    #     print("请输入云仓快递信息的文件路径，包含文件名，例如C:/快递单校对/云仓快递信息.xls")
    #     file_path = input()
    #     df = pd.read_excel(file_path,sheet_name="运费明细",dtype=str)
    #     df1 = pd.read_excel(file_path, sheet_name="拆单发货", dtype=str)
    #     df2 = df[["物流单号"]].copy()
    #     df2["company"] = "云仓"
    #     df2.rename(columns={"物流单号": "express_id"},inplace=True)
    #     df3 = df1[["物流单号"]].copy()
    #     df3["company"] = "云仓"
    #     df3.rename(columns={"物流单号": "express_id"}, inplace=True)
    #     dfs = [df2,df3,df4]
    #     df4 = pd.concat(dfs)
    #     print(f"快递信息合并后的总行数：{len(df4)}")
    #     df4["express_id"] = df4["express_id"].astype(str)
    #     df4["express_id"].dropna(axis=0,inplace=True)
    #     df4=df4[~df4.express_id.str.contains("nan")]
    #     print(len(df4))
    #     print(df4.head(5).to_markdown())
    #     df4.to_pickle("data/快递对账/express.pkl")
    #     return df4
    # else:
    #     print("无云仓数据，不处理！")

    return df4


def get_erp_express():
    # dfs = []
    # ms = win32api.MessageBox(0, "有电商文件吗？，输入“Y”或者“N”,输入其他则跳过处理！", "提醒", win32con.MB_YESNO)
    # if ms==6:
    #     df1 = get_file()
    #     dfs.append(df1)
    # else:
    #     pass

    # ms = win32api.MessageBox(0, "有erp文件吗？，输入“Y”或者“N”,输入其他则跳过处理！", "提醒", win32con.MB_YESNO)
    # if ms == 6:
    #     df2 = get_file1()
    #     dfs.append(df2)
    # else:
    #     pass
    #
    # ms = win32api.MessageBox(0, "有打印信息文件吗？，输入“Y”或者“N”,输入其他则跳过处理！", "提醒", win32con.MB_YESNO)
    # if ms == 6:
    #     df3 = get_file2()
    #     dfs.append(df3)
    # else:
    #     pass
    #
    # ms = win32api.MessageBox(0, "有快递信息文件吗？，输入“Y”或者“N”,输入其他则跳过处理！", "提醒", win32con.MB_YESNO)
    # if ms == 6:
    #     df4 = get_file3()
    # else:
    #     pass

    dfs = []
    df2 = get_file1()
    dfs.append(df2)
    df3 = get_file2()
    dfs.append(df3)
    df4 = get_file3()

    dfs = [df2,df3]
    df5 = pd.concat(dfs)
    print("正在汇总电商、erp、打印信息的数据：")
    print(f"汇总后的文件总行数：{len(df5)}")
    df5["express_id"] = df5["express_id"].apply(lambda x: x.replace("@", "").replace("=", "").replace('"', '').replace("'", "").strip())
    # df5["order_id"] = df5["order_id"].astype(str)
    df5["order_id"] = df5["order_id"].apply(lambda x: x.replace("@", "").replace("=", "").replace('"', '').replace("'", "").strip())
    df5 = df5.drop_duplicates(["express_id"], keep='first')
    df5.dropna(subset=["express_id"], inplace=True)
    df5 = df5[~df5.express_id.str.contains("nan")]

    df4["express_id"] = df4["express_id"].astype(str)
    df4["express_id"] = df4["express_id"].apply(lambda x: x.replace("@", "").replace("=", "").replace('"', '').replace("'", "").strip())
    df4 = df4[~df4.express_id.str.contains("nan")]

    print(f"汇总后去重的文件总行数：{len(df5)}")
    print(df5.head(5).to_markdown())

    print("请输入校对后的文件保存路径：")
    # result_path = input()
    result_path = r"D:\沙井\快递单校对\202201"
    print("正在匹配汇总后的数据与快递信息的校对...")

    df5["express_id"] = df5["express_id"].astype(str)
    df4["express_id"] = df4["express_id"].astype(str)
    print("df5:")
    print(df5.head(5).to_markdown())
    print("df4:")
    print(df4.head(5).to_markdown())
    df6 = df4.merge(df5, how="left", on="express_id")
    company_list = ["中通", "申通", "邮政", "圆通", "韵达", "京东"]
    for company in company_list:
        df_company = df6[df6.company.str.contains(company)]
        # df_company.to_excel(result_path + "/{}_快递单号校对结果-{}.xlsx".format(company,x + 1), engine='xlsxwriter')
        pagecount = math.ceil(df_company.shape[0] / 1000000.00)
        pagecount = "{:d}".format(pagecount)
        print("总共需要拆分{}个文件".format(pagecount))
        writer = pd.ExcelWriter(result_path + "/{}_快递单号校对结果.xlsx".format(company))
        for x in range(int(pagecount)):
            from_line = x * 1000000
            to_line = (x + 1) * 1000000
            df7 = df_company[from_line:to_line]
            print(df7.head(5).to_markdown())
            print("输出文件总行数：{}".format(df7.shape[0]))
            sheetname = "Sheet{}".format(x + 1)
            # df7.to_excel(result_path + "/{}_快递单号校对结果.xlsx".format(company), sheet_name=sheetname,engine='xlsxwriter')
            df7.to_excel(writer,sheetname)
        writer.save()

    print("全部数据处理完成！！！")

def open_file():
    dlg = win32ui.CreateFileDialog(1)  # 1表示打开文件对话框
    dlg.SetOFNInitialDir('E:/Python')  # 设置打开文件对话框中的初始显示目录
    dlg.DoModal()
    filename = dlg.GetPathName()  # 获取选择的文件名称
    print("filename=",filename)
    print("read ok")
    return filename

def test():
    df1 = pd.read_pickle("data/快递对账/dianshang.pkl")
    df1 = df1[~df1.express_id.str.contains("nan")]
    df1["express_id"] = df1["express_id"].apply(lambda x: x.replace("@", "").replace("=", "").replace('"', '').strip())
    df1["order_id"] = df1["order_id"].astype(str)
    df1["order_id"] = df1["order_id"].apply(lambda x: x.replace("@", "").replace("=", "").replace('"', '').strip())
    print(df1[df1["express_id"].str.contains("776300234845914")].to_markdown())
    df2 = pd.read_pickle("data/快递对账/erp.pkl")
    df2["express_id"] = df2["express_id"].apply(lambda x: x.replace("@", "").replace("=", "").replace('"', '').strip())
    df2["order_id"] = df2["order_id"].apply(lambda x: x.replace("@", "").replace("=", "").replace('"', '').strip())
    print(df2[df2["express_id"].str.contains("776300234845914")].to_markdown())
    df3 = pd.read_excel("data/快递对账/打印记录.xls",dtype=str)
    df3 = df3[["运单号", "店铺名", "收件人", "订单编号"]].copy()
    df3.rename(columns={"运单号": "express_id", "店铺名": "shop", "收件人": "name", "订单编号": "order_id"}, inplace=True)
    df3["express_id"] = df3["express_id"].astype(str)
    print(df3[df3["express_id"].str.contains("776300234845914")].to_markdown())
    df4 = pd.read_pickle("data/快递对账/express.pkl")
    df4["express_id"] = df4["express_id"].astype(str)
    df4["express_id"] = df4["express_id"].apply(lambda x: x.replace("@", "").strip())
    print(df4[df4["express_id"].str.contains("776300234845914")].to_markdown())
    print("打印文件行数：")
    print(f"{len(df1)}\n{len(df2)}\n{len(df3)}\n{len(df4)}")
    print(df1.head(5).to_markdown())
    print(df2.head(5).to_markdown())
    print(df3.head(5).to_markdown())
    print(df4.head(5).to_markdown())

    dfs = [df1, df2, df3]
    df5 = pd.concat(dfs)
    print("正在汇总电商、erp的数据：")
    print(f"汇总后的文件总行数：{len(df5)}")
    df5["express_id"] = df5["express_id"].str.strip()
    df5 = df5.drop_duplicates(["express_id"], keep='last')
    # df5["express_id"].dropna(axis=0, inplace=True)
    # df5 = df5[~df5.express_id.str.contains("nan")]
    print(f"汇总后去重的文件总行数：{len(df5)}")
    print(df5.head(5).to_markdown())
    # # df5.to_csv("data/erp-order.xlsx")

    print("请输入匹配后的文件保存路径：")
    result_path = r"C:\Users\mega\PycharmProjects\Shangwu\data\快递校对结果"
    print("正在匹配汇总后的数据与打印信息的校对...")

    df5["express_id"] = df5["express_id"].astype(str)
    df4["express_id"] = df4["express_id"].astype(str)
    print("df5:")
    print(df5.head(5).to_markdown())
    print("df4:")
    print(df4.head(5).to_markdown())
    df6 = df4.merge(df5, how="left", on="express_id")
    company_list = ["中通", "申通", "邮政", "云仓"]
    for company in company_list:
        df_company = df6[df6.company.str.contains(company)]
        # df_company.to_excel(result_path + "/{}_快递单号校对结果-{}.xlsx".format(company,x + 1), engine='xlsxwriter')
        pagecount = math.ceil(df_company.shape[0] / 1000000.00)
        pagecount = "{:d}".format(pagecount)
        print("总共需要拆分{}个文件".format(pagecount))
        writer = pd.ExcelWriter(result_path + "/{}_快递单号校对结果.xlsx".format(company))
        for x in range(int(pagecount)):
            from_line = x * 1000000
            to_line = (x + 1) * 1000000
            df7 = df_company[from_line:to_line]
            print(df7.head(5).to_markdown())
            print("输出文件总行数：{}".format(df7.shape[0]))
            sheetname = "Sheet{}".format(x + 1)
            # df7.to_excel(result_path + "/{}_快递单号校对结果.xlsx".format(company), sheet_name=sheetname,engine='xlsxwriter')
            df7.to_excel(writer,sheetname)
        writer.save()

def test1():
    dir_path = r"C:\Users\mega\Downloads\商务\快递信息"
    filenames = os.listdir(dir_path)
    index = 0
    dfs = []
    for name in filenames:
        print("filename:", name)
        temp_df=pd.read_excel(os.path.join(dir_path, name),sheet_name="对账单",dtype=str)
        temp_df["company"]=name
        if name.find("中通")>0:
            temp_df["company"] = "中通"
        elif name.find("申通")>0:
            temp_df["company"] = "申通"
        elif name.find("邮政") > 0:
            temp_df["company"] = "邮政"
        elif name.find("中通") > 0:
            temp_df["company"] = "中通"
        print(f"{name}行数{len(temp_df)}")
        temp_df = temp_df[["company", "运单号"]].copy()
        temp_df.rename(columns={"运单号": "express_id"}, inplace=True)
        temp_df.to_pickle("data/快递对账/express_file{}.pkl".format(index))
        dfs.append(temp_df)
        index += 1
    df = pd.read_excel(r"C:\Users\mega\Downloads\商务\云仓8月账单.xlsx", sheet_name="运费明细", dtype=str)
    df1 = pd.read_excel(r"C:\Users\mega\Downloads\商务\云仓8月账单.xlsx", sheet_name="拆单发货", dtype=str)
    df2 = df[["物流单号"]].copy()
    df2["company"] = "云仓"
    df2.rename(columns={"物流单号": "express_id"}, inplace=True)
    df3 = df1[["物流单号"]].copy()
    df3["company"] = "云仓"
    df3.rename(columns={"物流单号": "express_id"}, inplace=True)
    df2.to_pickle("data/快递对账/express_file_yun1.pkl")
    df3.to_pickle("data/快递对账/express_file_yun2.pkl")
    # df = pd.concat(dfs)

def test2():
    df1 = pd.read_pickle("data/快递对账/express_file0.pkl")
    df2 = pd.read_pickle("data/快递对账/express_file1.pkl")
    df3 = pd.read_pickle("data/快递对账/express_file2.pkl")
    df4 = pd.read_pickle("data/快递对账/express_file3.pkl")
    df5 = pd.read_pickle("data/快递对账/express_file4.pkl")
    df6 = pd.read_pickle("data/快递对账/express_file5.pkl")
    df7 = pd.read_pickle("data/快递对账/express_file_yun1.pkl")
    df8 = pd.read_pickle("data/快递对账/express_file_yun2.pkl")
    print("打印文件行数：")
    print(f"{len(df1)}\n{len(df2)}\n{len(df3)}\n{len(df4)}\n{len(df5)}\n{len(df6)}\n{len(df7)}\n{len(df8)}")
    dfs = [df1,df2,df3,df4,df5,df6,df7,df8]
    dfs = pd.concat(dfs)
    print("打印合并文件行数：")
    print(len(dfs))
    print(df1.head(5).to_markdown())
    print(df2.head(5).to_markdown())
    print(df3.head(5).to_markdown())
    print(df4.head(5).to_markdown())
    print(df5.head(5).to_markdown())
    print(df6.head(5).to_markdown())
    print(df7.head(5).to_markdown())
    print(df8.head(5).to_markdown())
    print(dfs.head(10).to_markdown())
    dfs.to_pickle("data/快递对账/express.pkl")


if __name__ == "__main__" :

    get_erp_express()

    # pyinstaller -p c:\pyexe\  -F .\快递对账.py

    # test()


