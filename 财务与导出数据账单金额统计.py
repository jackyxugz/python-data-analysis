# __coding=utf8__
# /** 作者：zengyanghui **/

import sys
import os
import pandas as pd
import numpy as np
import time
import os.path
import xlrd
import xlwt
import re
import tabulate
import traceback

# import win32api
# import win32ui
# import win32con
#
# import win32com
# from win32com.shell import shell


error_files=[]

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
        filekey_list = filekey_list.replace(",", " ").replace("，", " ")
        filekey = filekey_list.split(" ")
        pass
    else:
        filekey = ''

    # print("filekey:",filekey)
    _files = []
    list = os.listdir(rootdir)  # 列出文件夹下所有的目录与文件
    for i in range(0, len(list)):
        path = os.path.join(rootdir, list[i])
        # print("path:",path)
        if os.path.isdir(path):
            _files.extend(list_all_files(path, filekey_list))
        if os.path.isfile(path):
            # if path.find("~") < 0:  # 带~符号表示临时文件，不读取
            if ((path.find("~") < 0) and (path.find(".DS_Store") < 0) and (path.find(".csv.zip") < 0) and (path.find(".zip") < 0)):  # 带~符号表示临时文件，不读取
                # if len(filekey_list) > 0:
                #     t=re.search(filekey_list,path)
                #     if t:  # 如果匹配成功
                #         _files.append(path)

                if len(filekey) > 0:
                    break_flag=False
                    for key in filekey:
                        if not break_flag:
                            # print(path)
                            # print("key:",key)

                            # 简化版的不包含(类似正则表达式)  !京东 = 不包含京东
                            if ((len(key.replace("!",""))+1==len(key))  and (key.find("?")<0)  and (key.find(".")<0)  and (key.find("(")<0)  and (key.find(")")<0)):
                                if path.find(key.replace("!",""))>=0:
                                    # 要求不包含，结果找到了！
                                    # print("要求不包含{}，结果找到了！".format(key.replace("!","")))
                                    break_flag = True
                            # else:
                            #     t=re.search(key,path)
                            #     if t:  # 如果匹配成功
                            #         print("匹配成功")
                            #         pass
                            #     else:
                            #         print("匹配失败！")
                            #         # 只要有一项匹配不成功，则自动退出，认为不符合条件
                            #         break_flag=True

                    if not break_flag:
                        # print("math ok ",path)
                        _files.append(path)
                    else:
                        # print("math error ",path)
                        pass

                else:
                    _files.append(path)

    # print(_files)
    return _files

def get_zhangjiali_shopname(filename):
    # /Volumes/IT审计处理需求/neworder-2019/天猫、淘宝/张佳丽/账单/Ceci茜茜美妆店支付宝201912_zip\20885314795456850156_201912_账务明细_1.csv
    if filename.find("张佳丽")>=0:
        # print("张佳丽店铺名：", filename)
        if filename.find("支付宝") >= 0:
            filename = "".join(filename.split("支付宝")[0])
        elif filename.find("结算单") >= 0:
            filename = "".join(filename.split("结算单")[0])
        # print(filename)
        filename = filename.split("/")[len(filename.split("/"))-1]
        # print("解析后店铺名：",filename)
        return filename

def save_error_log(error_files):
    error_log=pd.DataFrame(error_files,columns=["filename"])
    print(error_files)
    error_log.to_excel("data/error_files.xlsx")

def read_excel(filename):
    print(filename)
    if ((filename.find("淘宝") > 0 or filename.find("TAOBAO") > 0) and (filename.count("淘宝") > filename.count("天猫"))):
        plat = "淘宝"
        # table.to_excel(default_dir+"/"+data+"的收入和支出.xlsx",index=False)
        if filename.find("TAOBAO") >= 0:
            df = pd.read_excel(filename, skiprows=1, dtype=str)
            df = df.apply(
                lambda x: x.astype(str).str.replace("¥", "").str.replace(":", "").str.replace("'", "").str.replace('"',''))
            print("导出文件不需要额外处理")
        else:
            if filename.find("海外") >= 0:
                if filename.find("支付宝") < 0:
                    if ((filename.find("结算单") >= 0) and (filename.find("settle") >= 0)):
                        if os.path.exists(default_dir + "/海外-淘宝settle账单.xlsx"):
                            print(f"海外-淘宝settle账单文件已合并")
                            df = pd.read_excel(default_dir + "/海外-淘宝settle账单.xlsx",dtype=str)
                            return df
                        else:
                            if filename.find("xls") >= 0:
                                df = pd.read_excel(filename, dtype=str)
                            else:
                                df = pd.read_csv(filename, dtype=str, encoding="gb18030")
                            df["平台"] = plat
                            df["店铺"] = filename
                            if filename.find("张佳丽") >= 0:
                                df["店铺"] = df["店铺"].apply(lambda x: get_zhangjiali_shopname(x))
                                # df["店铺"] = "".join(filename.split("支付宝")[0])
                                # df["店铺"] = "".join(df["店铺"][0].split("/")[-1:])
                            else:
                                df["店铺"] = "".join(filename.split("/")[-3:-2])
                            df["收入"] = df["Rmb_amount"].astype(float)
                            # df["Fee"] = df["Fee"].astype(float)
                            # df["Rate"] = df["Rate"].astype(float)
                            # df["支出"] = df["Fee"] * df["Rate"]
                            df["支出"] = (df["Fee"].astype(float)) * (df["Rate"].astype(float))
                            df["发生时间"] = df["Settlement_time"].astype(str).apply(lambda x: x.replace(".", "-"))
                            df["发生时间"] = df["发生时间"].astype("datetime64[ns]")
                            df["年度"] = df["发生时间"].apply(lambda x: x.year)
                            df["月份"] = df["发生时间"].apply(lambda x: x.month)
                            print("海外-淘宝settle账单")
                            file = "/海外-淘宝settle账单.xlsx"
                            # print(df.tail(1).to_markdown())
                else:
                    if filename.find("账务明细") >= 0:
                        if os.path.exists(default_dir + "/海外-淘宝账务明细.xlsx"):
                            print(f"海外-淘宝账务明细文件已合并")
                            df = pd.read_excel(default_dir + "/海外-淘宝账务明细.xlsx", dtype=str)
                            return df
                        else:
                            if filename.find("xls") >= 0:
                                df = pd.read_excel(filename, dtype=str)
                            else:
                                df = pd.read_csv(filename, dtype=str, skiprows=4, encoding="gb18030")
                            df = df[~df["账务流水号"].str.contains("#")]
                            df["平台"] = plat
                            df["店铺"] = filename
                            if filename.find("张佳丽") >= 0:
                                df["店铺"] = df["店铺"].apply(lambda x: get_zhangjiali_shopname(x))
                            else:
                                df["店铺"] = "".join(filename.split("/")[-3:-2])
                            df["收入"] = df["收入金额（+元）"].astype(float)
                            df["支出"] = df["支出金额（-元）"].astype(float)
                            df["发生时间"] = df["发生时间"].astype(str).apply(lambda x: x.replace(".", "-"))
                            df["发生时间"] = df["发生时间"].astype("datetime64[ns]")
                            df["年度"] = df["发生时间"].apply(lambda x: x.year)
                            df["月份"] = df["发生时间"].apply(lambda x: x.month)
                            print("海外-淘宝账务明细")
                            file = "/海外-淘宝账务明细.xlsx"
                    else:
                        if os.path.exists(default_dir + "/海外-淘宝支付宝账单.xlsx"):
                            print(f"海外-淘宝支付宝账单文件已合并")
                            df = pd.read_excel(default_dir + "/海外-淘宝支付宝账单.xlsx", dtype=str)
                            return df
                        else:
                            if filename.find("xls") >= 0:
                                try:
                                    print("尝试读取xls")
                                    df = pd.read_excel(filename, dtype=str)
                                except Exception as e:
                                    dict = {"平台": "", "店铺": "", "月份": "", "收入": "", "支出": "", "收入-支出": "", "数据来源": ""}
                                    df = pd.DataFrame(dict, index=[0])
                                    # 这个是输出错误类别的，如果捕捉的是通用错误，其实这个看不出来什么
                                    # print('str(Exception):\t', str(Exception))  # 输出  str(Exception):	<type 'exceptions.Exception'>
                                    # 这个是输出错误的具体原因，这步可以不用加str，输出
                                    # print('str(e):\t\t', str(e))  # 输出 str(e):		integer division or modulo by zero
                                    # print('repr(e):\t', repr(e))  # 输出 repr(e):	ZeroDivisionError('integer division or modulo by zero',)
                                    # print('traceback.print_exc():')
                                    # 以下两步都是输出错误的具体位置的
                                    # traceback.print_exc()
                                    # print('traceback.format_exc():\n%s' % traceback.format_exc())
                                    print("异常文件！")
                                    error_files.append(filename)
                                    save_error_log(error_files)
                                    return df
                            else:
                                df = pd.read_csv(filename, dtype=str, encoding="gb18030")
                            df["平台"] = plat
                            df["店铺"] = filename
                            if filename.find("张佳丽") >= 0:
                                df["店铺"] = df["店铺"].apply(lambda x: get_zhangjiali_shopname(x))
                            else:
                                df["店铺"] = "".join(filename.split("/")[-3:-2])
                            df["收入"] = df["Rmb_amount"].astype(float)

                            df["支出"] = (df["Fee"].astype(float)) * (df["Rate"].astype(float))
                            df["发生时间"] = df["Settlement_time"].astype(str).apply(lambda x: x.replace(".", "-"))
                            df["发生时间"] = df["发生时间"].astype("datetime64[ns]")
                            df["年度"] = df["发生时间"].apply(lambda x: x.year)
                            df["月份"] = df["发生时间"].apply(lambda x: x.month)
                            print("海外-淘宝支付宝")
                            file = "/海外-淘宝支付宝账单.xlsx"
            else:
                if filename.find("账务明细") >= 0:
                    if os.path.exists(default_dir + "/国内-淘宝账务明细.xlsx"):
                        print(f"国内-淘宝账务明细文件已合并")
                        df = pd.read_excel(default_dir + "/国内-淘宝账务明细.xlsx", dtype=str)
                        return df
                    else:
                        if filename.find("csv") >= 0:
                            df = pd.read_csv(filename, skiprows=4, dtype=str, encoding="gb18030")
                            df = df[~df["账务流水号"].str.contains("#")]
                            df["平台"] = plat
                            df["店铺"] = filename
                            if filename.find("张佳丽") >= 0:
                                df["店铺"] = df["店铺"].apply(lambda x: get_zhangjiali_shopname(x))
                            else:
                                df["店铺"] = "".join(filename.split("/")[-3:-2])
                            df["收入"] = df["收入金额（+元）"].astype(float)
                            df["支出"] = df["支出金额（-元）"].astype(float)
                            df["发生时间"] = df["发生时间"].astype(str).apply(lambda x: x.replace(".", "-"))
                            df["发生时间"] = df["发生时间"].astype("datetime64[ns]")
                            df["年度"] = df["发生时间"].apply(lambda x: x.year)
                            df["月份"] = df["发生时间"].apply(lambda x: x.month)
                            print("国内-淘宝账务明细csv")
                            file = "/国内-淘宝账务明细.xlsx"
                            # print(df.tail(1).to_markdown())
                        else:
                            print("国内-淘宝账务明细xlsx")
                            print(filename)
                else:
                    if filename.find("支付宝") >= 0:
                        if os.path.exists(default_dir + "/国内-淘宝支付宝账单.xlsx"):
                            print(f"国内-淘宝支付宝账单文件已合并")
                            df = pd.read_excel(default_dir + "/国内-淘宝支付宝账单.xlsx", dtype=str)
                            return df
                        else:
                            df = pd.read_excel(filename, sheet_name=None, dtype=str)
                            sheet_list = list(df)
                            print(sheet_list)
                            for sheet in sheet_list:
                                if sheet.find("汇总") >= 0:
                                    # print("定位1")
                                    print(sheet)
                                    continue
                                elif sheet.find("账单") >= 0:
                                    # print("定位2")
                                    print(sheet)
                                    sheet_name = sheet
                                    break
                                elif sheet.find("账务明") >= 0:
                                    # print("定位3")
                                    print(sheet)
                                    sheet_name = sheet
                                    break
                                else:
                                    # print("定位4")
                                    print(sheet)
                                    sheet_name = "Sheet1"
                            print(sheet_name)
                            df = pd.read_excel(filename, sheet_name=sheet_name, dtype=str)
                            # print(df.tail(1).to_markdown())
                            df["平台"] = plat
                            df["店铺"] = "".join(filename.split("/")[-3:-2])
                            if "收入金额（+元）" in df.columns:
                                df["收入"] = df["收入金额（+元）"].astype(float)
                            elif "收入（+元）" in df.columns:
                                df["收入（+元）"] = df["收入（+元）"].replace(" ", "0")
                                df["收入"] = df["收入（+元）"].astype(float)
                            if "支出金额（-元）" in df.columns:
                                df["支出"] = df["支出金额（-元）"].astype(float)
                            elif "支出（-元）" in df.columns:
                                df["支出（-元）"] = df["支出（-元）"].replace(" ", "0")
                                df["支出"] = df["支出（-元）"].astype(float)
                            if "发生时间" in df.columns:
                                df["发生时间"] = df["发生时间"].astype(str).apply(lambda x: x.replace(".", "-"))
                                df["发生时间"] = df["发生时间"].astype("datetime64[ns]")
                                df["年度"] = df["发生时间"].apply(lambda x: x.year)
                                df["月份"] = df["发生时间"].apply(lambda x: x.month)
                            elif "入账时间" in df.columns:
                                df["入账时间"] = df["入账时间"].astype(str).apply(lambda x: x.replace(".", "-"))
                                df["入账时间"] = df["入账时间"].astype("datetime64[ns]")
                                df["年度"] = df["入账时间"].apply(lambda x: x.year)
                                df["月份"] = df["入账时间"].apply(lambda x: x.month)
                            print("国内-支付宝")
                            file = "/国内-淘宝支付宝账单.xlsx"
                            # print(df.tail(1).to_markdown())
                    else:
                        print(f"非淘宝账单文件{filename}")
                        dict = {"平台": "", "店铺": "", "月份": "", "收入": "", "支出": "", "收入-支出": "", "数据来源": ""}
                        df = pd.DataFrame(dict, index=[0])
                        return df
        df.drop_duplicates(inplace=True)
        print("淘宝")
    elif ((filename.find("天猫") > 0 or filename.find("TMALL") > 0) and (filename.count("淘宝") <= filename.count("天猫"))):
        plat = "天猫"
        if filename.find("TMALL") >= 0:
            df = pd.read_excel(filename, skiprows=1, dtype=str)
            df = df.apply(
                lambda x: x.astype(str).str.replace("¥", "").str.replace(":", "").str.replace("'", "").str.replace('"',''))
            print("导出文件不需要额外处理")
        else:
            if filename.find("海外") >= 0:
                if filename.find("支付宝") < 0:
                    if ((filename.find("结算单") >= 0) and (filename.find("settle") >= 0)):
                        # df = df[~df["账务流水号"].str.contains("#")]
                        # print(filename)
                        if filename.find("xls") >= 0:
                            df = pd.read_excel(filename,dtype=str)
                        else:
                            df = pd.read_csv(filename,dtype=str,encoding="gb18030")
                        df["平台"] = plat
                        df["店铺"] = filename
                        if filename.find("张佳丽")>=0:
                            df["店铺"]=df["店铺"].apply(lambda x: get_zhangjiali_shopname(x))
                            # df["店铺"] = "".join(filename.split("支付宝")[0])
                            # df["店铺"] = "".join(df["店铺"][0].split("/")[-1:])
                        else:
                            df["店铺"] = "".join(filename.split("/")[-3:-2])
                        df["收入"] = df["Rmb_amount"].astype(float)
                        # df["Fee"] = df["Fee"].astype(float)
                        # df["Rate"] = df["Rate"].astype(float)
                        # df["支出"] = df["Fee"] * df["Rate"]
                        df["支出"] = (df["Fee"].astype(float)) * (df["Rate"].astype(float))
                        df["发生时间"] = df["Settlement_time"].astype(str).apply(lambda x: x.replace(".", "-"))
                        df["发生时间"] = df["发生时间"].astype("datetime64[ns]")
                        df["年度"] = df["发生时间"].apply(lambda x: x.year)
                        df["月份"] = df["发生时间"].apply(lambda x: x.month)
                        print("海外-天猫账单1")
                        # print(df.tail(1).to_markdown())
                    else:
                        print(f"海外-非天猫账单文件{filename}")
                        dict = {"平台": "", "店铺": "", "月份": "", "收入": "", "支出": "", "收入-支出": "", "数据来源": ""}
                        df = pd.DataFrame(dict, index=[0])
                        return df
                else:
                    if filename.find("账务明细") >= 0:
                        if filename.find("xls") >= 0:
                            df = pd.read_excel(filename, dtype=str)
                        else:
                            df = pd.read_csv(filename, dtype=str, skiprows=4, encoding="gb18030")
                        df = df[~df["账务流水号"].str.contains("#")]
                        df["平台"] = plat
                        df["店铺"] = filename
                        if filename.find("张佳丽") >= 0:
                            df["店铺"] = df["店铺"].apply(lambda x: get_zhangjiali_shopname(x))
                        else:
                            df["店铺"] = "".join(filename.split("/")[-3:-2])
                        df["收入"] = df["收入金额（+元）"].astype(float)
                        df["支出"] = df["支出金额（-元）"].astype(float)
                        df["发生时间"] = df["发生时间"].astype(str).apply(lambda x: x.replace(".", "-"))
                        df["发生时间"] = df["发生时间"].astype("datetime64[ns]")
                        df["年度"] = df["发生时间"].apply(lambda x: x.year)
                        df["月份"] = df["发生时间"].apply(lambda x: x.month)
                        print("海外-天猫账务明细")
                    else:
                        if filename.find("xls") >= 0:
                            try:
                                print("尝试读取xls")
                                df = pd.read_excel(filename, dtype=str)
                            except Exception as e:
                                dict = {"平台": "", "店铺": "", "月份": "", "收入": "", "支出": "", "收入-支出": "", "数据来源": ""}
                                df = pd.DataFrame(dict, index=[0])
                                # 这个是输出错误类别的，如果捕捉的是通用错误，其实这个看不出来什么
                                # print('str(Exception):\t', str(Exception))  # 输出  str(Exception):	<type 'exceptions.Exception'>
                                # 这个是输出错误的具体原因，这步可以不用加str，输出
                                # print('str(e):\t\t', str(e))  # 输出 str(e):		integer division or modulo by zero
                                # print('repr(e):\t', repr(e))  # 输出 repr(e):	ZeroDivisionError('integer division or modulo by zero',)
                                # print('traceback.print_exc():')
                                # 以下两步都是输出错误的具体位置的
                                # traceback.print_exc()
                                # print('traceback.format_exc():\n%s' % traceback.format_exc())
                                print("异常文件！")
                                error_files.append(filename)
                                save_error_log(error_files)
                                return df
                        else:
                            df = pd.read_csv(filename, dtype=str, encoding="gb18030")
                        df["平台"] = plat
                        df["店铺"] = filename
                        if filename.find("张佳丽") >= 0:
                            df["店铺"] = df["店铺"].apply(lambda x: get_zhangjiali_shopname(x))
                        else:
                            df["店铺"] = "".join(filename.split("/")[-4:-3])
                        df["收入"] = df["Rmb_amount"].astype(float)

                        df["支出"] = (df["Fee"].astype(float)) * (df["Rate"].astype(float))
                        df["发生时间"] = df["Settlement_time"].astype(str).apply(lambda x: x.replace(".", "-"))
                        df["发生时间"] = df["发生时间"].astype("datetime64[ns]")
                        df["年度"] = df["发生时间"].apply(lambda x: x.year)
                        df["月份"] = df["发生时间"].apply(lambda x: x.month)
                        print("海外-天猫支付宝")
            else:
                if filename.find("账务明细") >= 0:
                    if filename.find("csv") >= 0:
                        df = pd.read_csv(filename, skiprows=4, dtype=str,encoding="gb18030")
                        df = df[~df["账务流水号"].str.contains("#")]
                        df["平台"] = plat
                        df["店铺"] = filename
                        if filename.find("张佳丽") >= 0:
                            df["店铺"] = df["店铺"].apply(lambda x: get_zhangjiali_shopname(x))
                        else:
                            df["店铺"] = "".join(filename.split("/")[-3:-2])
                        df["收入"] = df["收入金额（+元）"].astype(float)
                        df["支出"] = df["支出金额（-元）"].astype(float)
                        df["发生时间"] = df["发生时间"].astype(str).apply(lambda x: x.replace(".", "-"))
                        df["发生时间"] = df["发生时间"].astype("datetime64[ns]")
                        df["年度"] = df["发生时间"].apply(lambda x: x.year)
                        df["月份"] = df["发生时间"].apply(lambda x: x.month)
                        print("国内-天猫账务明细csv")
                        # print(df.tail(1).to_markdown())
                    else:
                        print("国内-天猫账务明细xlsx")
                        pass
                else:
                    if filename.find("支付宝") >= 0:
                        df = pd.read_excel(filename, sheet_name=None, dtype=str)
                        sheet_list = list(df)
                        print(sheet_list)
                        for sheet in sheet_list:
                            if sheet.find("汇总") >= 0:
                                # print("定位1")
                                print(sheet)
                                continue
                            elif sheet.find("账单") >= 0:
                                # print("定位2")
                                print(sheet)
                                sheet_name = sheet
                                break
                            elif sheet.find("账务明") >= 0:
                                # print("定位3")
                                print(sheet)
                                sheet_name = sheet
                                break
                            else:
                                # print("定位4")
                                print(sheet)
                                sheet_name = "Sheet1"
                        print(sheet_name)
                        df = pd.read_excel(filename,sheet_name=sheet_name,dtype=str)
                        # print(df.tail(1).to_markdown())
                        df["平台"] = plat
                        df["店铺"] = "".join(filename.split("/")[-3:-2])
                        if "收入金额（+元）" in df.columns:
                            df["收入"] = df["收入金额（+元）"].astype(float)
                        elif "收入（+元）" in df.columns:
                            df["收入（+元）"] = df["收入（+元）"].replace(" ","0")
                            df["收入"] = df["收入（+元）"].astype(float)
                        if "支出金额（-元）" in df.columns:
                            df["支出"] = df["支出金额（-元）"].astype(float)
                        elif "支出（-元）" in df.columns:
                            df["支出（-元）"] = df["支出（-元）"].replace(" ","0")
                            df["支出"] = df["支出（-元）"].astype(float)
                        if "发生时间" in df.columns:
                            df["发生时间"] = df["发生时间"].astype(str).apply(lambda x: x.replace(".", "-"))
                            df["发生时间"] = df["发生时间"].astype("datetime64[ns]")
                            df["年度"] = df["发生时间"].apply(lambda x: x.year)
                            df["月份"] = df["发生时间"].apply(lambda x: x.month)
                        elif "入账时间" in df.columns:
                            df["入账时间"] = df["入账时间"].astype(str).apply(lambda x: x.replace(".", "-"))
                            df["入账时间"] = df["入账时间"].astype("datetime64[ns]")
                            df["年度"] = df["入账时间"].apply(lambda x: x.year)
                            df["月份"] = df["入账时间"].apply(lambda x: x.month)
                        print("国内-支付宝")
                        # print(df.tail(1).to_markdown())
                    else:
                        print(f"非天猫账单文件{filename}")
                        dict = {"平台": "", "店铺": "", "月份": "", "收入": "", "支出": "", "收入-支出": "", "数据来源": ""}
                        df = pd.DataFrame(dict, index=[0])
                        return df
        df.drop_duplicates(inplace=True)
        print("天猫")
    elif (filename.find("京东") > 0 or filename.find("JD") > 0):
        plat = "京东"
        if filename.find("JD") >= 0:
            df = pd.read_excel(filename, skiprows=1, dtype=str)
            df = df.apply(
                lambda x: x.astype(str).str.replace("¥", "").str.replace(":", "").str.replace("'", "").str.replace('"',''))
            print("导出文件不需要额外处理")
        else:
            if filename.find("xls") >= 0:
                df = pd.read_excel(filename, dtype=str)
            else:
                df = pd.read_csv(filename, dtype=str, encoding="gb18030")
                if filename.find("妥投销货清单明细") >= 0:
                    return_file = filename[:-12] + "退货结算数据.csv"  # 对文件名进行修改
                    df1 = pd.read_csv(return_file, skiprows=1, keep_default_na=False, dtype=str, encoding="gb18030")
                    df = df[:-1]
                    df1 = df1[:-1]
                    if df1.shape[0] > 0:
                        df = pd.merge(df, df1[["订单编号", "退款金额"]], how="outer", on="订单编号")
                    else:
                        pass
                else:
                    pass
            df = df.apply(lambda x:x.astype(str).str.replace("=","").str.replace('"','').str.replace("'",""))
            df.dropna(inplace=True)
            df.drop_duplicates(inplace=True)
            df["平台"] = plat
            if filename.find("月账单") >= 0:
                file = "".join(filename.split("/")[-1:])
                df["店铺"] = "".join(file.split("月账单")[0])
            elif filename.find("结算单") >= 0:
                df["店铺"] = "".join(filename.split("/")[-3:-2])
            else:
                df["店铺"] = "".join(filename.split("/")[-2:-1])
            df["店铺"] = df["店铺"].str.replace("京东", "")
            if filename.find("海外") >= 0:
                df["收入"] = df["商品应结金额"].astype(float) * df["汇率(美元/人民币)"].astype(float)
                if "退款金额" in df.columns:
                    df["支出"] = (df["商品佣金"].astype(float) * df["汇率(美元/人民币)"].astype(float)) + (df["退款金额"].astype(float) * df["汇率(美元/人民币)"].astype(float))
                else:
                    df["支出"] = df["商品佣金"].astype(float) * df["汇率(美元/人民币)"].astype(float)
                # df["发生时间"] = df["费用发生时间"].astype(str).apply(lambda x: x.replace(".", "-"))
                file = "".join(filename.split("/")[-2:-1])
                time = "".join(file[-9:-3])
                # df["发生时间"] = df["发生时间"].astype("datetime64[ns]")
                df["年度"] = time[:4]
                df["月份"] = time[-2:]
                print("京东海外")
            else:
                df["收入"] = df["金额"][df["收支方向"].str.contains("收入")].astype(float)
                df["支出"] = df["金额"][df["收支方向"].str.contains("支出")].astype(float)
                df["发生时间"] = df["费用发生时间"].astype(str).apply(lambda x: x.replace(".", "-"))
                df["发生时间"] = df["发生时间"].astype("datetime64[ns]")
                df["年度"] = df["发生时间"].apply(lambda x: x.year)
                df["月份"] = df["发生时间"].apply(lambda x: x.month)
                print("京东国内")
        df.drop_duplicates(inplace=True)
    elif (filename.find("拼多多") > 0 or filename.find("PDD") > 0):
        plat = "拼多多"
        if filename.find("PDD") >= 0:
            df = pd.read_excel(filename, skiprows=1, dtype=str)
            df = df.apply(
                lambda x: x.astype(str).str.replace("¥", "").str.replace(":", "").str.replace("'", "").str.replace('"',''))
            print("导出文件不需要额外处理")
        else:
            if filename.find("账单") >= 0:
                if filename.find("xls") >= 0:
                    df = pd.read_excel(filename,dtype=str)
                else:
                    df = pd.read_csv(filename,skiprows=4,dtype=str,encoding="gb18030")
                df.dropna(inplace=True)
                df = df[~df["商户订单号"].str.contains("#")]
                df["平台"] = plat
                df["店铺"] = "".join(filename.split("/")[-2:-1])
                df["店铺"] = df["店铺"].replace("拼多多","")
                df["收入"] = df["收入金额（+元）"].astype(float)
                df["支出"] = df["支出金额（-元）"].astype(float)
                df["发生时间"] = df["发生时间"].astype(str).apply(lambda x: x.replace(".", "-"))
                df["发生时间"] = df["发生时间"].astype("datetime64[ns]")
                df["年度"] = df["发生时间"].apply(lambda x: x.year)
                df["月份"] = df["发生时间"].apply(lambda x: x.month)
                print("拼多多-账务明细")
            else:
                print("拼多多-无账单")
        df.drop_duplicates(inplace=True)
    elif ((0 < filename.find("抖音") < 45) or (0 < filename.find("DY") < 45)):
        plat = "抖音"
        print("抖音")
    elif ((0 < filename.find("快手") < 45) or (0 < filename.find("KS") < 45)):
        plat = "快手"
        print("快手")
    elif (filename.find("小红书") > 0 or filename.find("XHS") > 0):
        plat = "小红书"
        if filename.find("XHS") >= 0:
            df = pd.read_excel(filename, skiprows=1, dtype=str)
            df = df.apply(
                lambda x: x.astype(str).str.replace("¥", "").str.replace(":", "").str.replace("'", "").str.replace('"',''))
            print("导出文件不需要额外处理")
        else:
            if filename.find("账单") >= 0:
                if filename.find("xls") >= 0:
                    df = pd.read_excel(filename,skiprows=10,dtype=str)
                    # df2 = pd.read_excel(filename,skiprows=26,dtype=str)
                else:
                    df = pd.read_csv(filename,skiprows=4,dtype=str,encoding="gb18030")
                if filename.find("XHS") >= 0:
                    print("导出文件不用处理")
                else:
                    df.dropna(subset=["项目","金额（人民币）"],inplace=True)
                    print(df.to_markdown())
                    # df.columns=["项目","列","收支方向","金额"]
                    del df["Unnamed: 1"]
                    df = df.reset_index(drop=True)
                    df = df[~df["项目"].str.contains("项目")]
                    df["金额（人民币）"] = df["金额（人民币）"].astype(float)
                    num1 = df["金额（人民币）"][df["项目"].str.contains("商品销售|订单运费|销售佣金")].sum()
                    num2 = df["金额（人民币）"][~df["项目"].str.contains("商品销售|订单运费")].sum()
                    print(f"num1:{num1}")
                    print(f"num2:{num2}")
                    print(df.to_markdown())
                    df["平台"] = plat
                    "/Volumes/IT审计处理需求/neworder-2019/小红书/小红书Chiara Bca Ambra海外品牌店/小红书Chiara Bca Ambra海外品牌店201907-账单.xlsx"
                    df["店铺"] = "".join(filename.split("/")[-2:-1])
                    df["店铺"] = df["店铺"].str.replace("小红书", "")
                    df["收入"] = num1
                    df["支出"] = num2
                    month = "".join(filename.split("-")[:-1])
                    df["月份"] = "".join(month.split("2019")[-1:])
                    print(df.to_markdown())
                    print("小红书-账务明细")
            else:
                print("小红书-无账单")
        df.drop_duplicates(inplace=True)
    elif (filename.find("考拉") > 0 or filename.find("KAOLA") > 0):
        plat = "网易考拉"
        if filename.find("KAOLA") >= 0:
            df = pd.read_excel(filename, skiprows=1, dtype=str)
            df = df.apply(
                lambda x: x.astype(str).str.replace("¥", "").str.replace(":", "").str.replace("'", "").str.replace('"',''))
            print("导出文件不需要额外处理")
        else:
            if filename.find("账单") >= 0:
                if filename.find("xls") >= 0:
                    df = pd.read_excel(filename,sheet_name="总账单",dtype=str)
                else:
                    df = pd.read_csv(filename,skiprows=4,dtype=str,encoding="gb18030")
                df.dropna(inplace=True)
                df["平台"] = plat
                df["店铺"] = "".join(filename.split("/")[-2:-1])
                df["店铺"] = df["店铺"].str.replace("网易考拉","")
                df["收入"] = df["商品销售总金额"].astype(float)
                df["支出"] = (df["应扣商家优惠总金额"].astype(float)) + (df["退款总金额"].astype(float)) + (df["平台技术服务费"].astype(float)) + (
                    df["商品税费总金额"].astype(float)) + (df["运费总金额"].astype(float)) + (df["本期赔付金额"].astype(float)) + (
                               df["其他金额"].astype(float)) + (df["调整金额"].astype(float))
                df["发生时间"] = df["日期"].astype(str).apply(lambda x: x.replace(".", "-"))
                df["发生时间"] = df["发生时间"].astype("datetime64[ns]")
                df["年度"] = df["发生时间"].apply(lambda x: x.year)
                df["月份"] = df["发生时间"].apply(lambda x: x.month)
                print("网易考拉-账务明细")
            else:
                print("网易考拉-无账单")
        df.drop_duplicates(inplace=True)
    elif (filename.find("有赞") > 0 or filename.find("YZ") > 0):
        plat = "有赞"
        if filename.find("YZ") >= 0:
            df = pd.read_excel(filename, skiprows=1, dtype=str)
            df = df.apply(
                lambda x: x.astype(str).str.replace("¥", "").str.replace(":", "").str.replace("'", "").str.replace('"',''))
            print("导出文件不需要额外处理")
        else:
            if filename.find("账单") >= 0:
                if filename.find("xls") >= 0:
                    df = pd.read_excel(filename,dtype=str)
                else:
                    df = pd.read_csv(filename,skiprows=4,dtype=str,encoding="gb18030")
                df.dropna(inplace=True)
                df["平台"] = plat
                df["店铺"] = df["账务主体"]
                df["收入"] = df["收入(元)"].astype(float)
                df["支出"] = df["支出(元)"].astype(float)
                df["发生时间"] = df["入账时间"].astype(str).apply(lambda x: x.replace(".", "-"))
                df["发生时间"] = df["发生时间"].astype("datetime64[ns]")
                df["年度"] = df["发生时间"].apply(lambda x: x.year)
                df["月份"] = df["发生时间"].apply(lambda x: x.month)
                print("有赞-账务明细")
            else:
                print("有赞-无账单")
        df.drop_duplicates(inplace=True)
    elif filename.find("金牛") > 0:
        plat = "金牛电商"
        print("金牛电商")
    elif filename.find("百度") > 0:
        plat = "金牛电商"
        print("百度")
    else:
        dict = {"平台": "", "店铺名称": "", "年度": "", "月份": "", "订单数量": "", "订单金额": ""}
        df = pd.DataFrame(dict, index=[0])
        print("其他平台")
        return df
    # for column_name in df.columns:
    #     df.rename(columns={column_name:column_name.replace(" ","").replace("\n","").strip()},inplace=True)
    # print(df.head(1).to_markdown())
    if filename.find("导出") >= 0:
        month = "".join(filename.split("/")[-2:-1])
        if len(month) > 2:
            month = "".join(filename.split("/")[-3:-2])
        else:
            pass
        df["月份"] = month
        df["收入"] = df["收入"].astype(float)
        df["支出"] = df["支出"].astype(float)
    elif filename.find("财务") >= 0:
        pass

    temp_df = df.groupby(["平台", "店铺", "月份"]).agg({"收入": "sum", "支出": "sum"})
    temp_df = pd.DataFrame(temp_df).reset_index()
    temp_df["收入-支出"] = (temp_df["收入"] - temp_df["支出"].abs())
    print(temp_df.head(5).to_markdown())


    temp_df.to_excel(default_dir + file,index=False)
    return temp_df



def get_all_files(rootdir, filekey):
    filelist = list_all_files(rootdir, filekey)
    # print(filelist)
    mySeries = pd.Series(filelist)
    df = pd.DataFrame(mySeries)
    df.columns = ["filename"]
    # df["filename"] = df["filename"].apply(lambda x: x.strip().replace(" ", ""))
    # print()

    df=df[~df["filename"].str.contains("总表")]
    df=df[~df["filename"].str.contains("汇总")]
    df=df[~df["filename"].str.contains("订单")]
    # df=df[~df["filename"].str.contains("csv.zip")]
    df=df[~df["filename"].str.contains("业务明细")]
    df=df[~df["filename"].str.contains("其他费用")]
    df=df[~df["filename"].str.contains("无账单")]
    df=df[~df["filename"].str.contains("未下载")]
    df=df[~df["filename"].str.contains("推广费")]
    df=df[~df["filename"].str.contains("结算数据")]
    df=df[~df["filename"].str.contains("fee")]
    df=df[~df["filename"].str.contains("settlebatch")]
    df=df[~df["filename"].str.contains("strade")]
    df=df[~df["filename"].str.contains("myaccount")]

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
    print('请选择要合并的文件目录（请确定所有excel的表头都是相同的哦！）:')
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
    # filekey = input()
    filekey = "xlsx|csv !.csv.zip"

    if len(filedir) == 0:
        print("你没有输入任何关键词 :(")
        filekey = ''
        # sys.exit()
        # return

    print("你希望在'{}'目录下找到所有的包含“{}”文件，然后合并。".format(filedir, filekey))

    table = read_all_excel(filedir, filekey)
    if default_dir.find("财务数据") >= 0:
        data = "财务账单"
        table["数据来源"] = data
    else:
        data = "导出账单"
        table["数据来源"] = data
    # if len(table) > 500000:
    #     table.to_csv("data/导出账单的收入和支出.csv",index=False)
    # else:
    table.to_excel(default_dir+"/"+data+"的收入和支出.xlsx",index=False)

    table1 = table[table["数据来源"].str.contains("财务账单")]
    table2 = table[table["数据来源"].str.contains("导出账单")]
    table1.to_excel(default_dir + "/财务账单的收入和支出.xlsx", index=False)
    table2.to_excel(default_dir + "/导出账单的收入和支出.xlsx", index=False)

    return table1,table2


def groupby_amt():
    # if default_dir.find("财务数据") >= 0:
    #     data = "财务数据"
    # else:
    #     data = "导出数据"
    filename = default_dir+"/财务账单的收入和支出.xlsx"
    df = pd.read_excel(filename)
    del df["filename"]
    df.drop_duplicates(inplace=True)
    df.dropna(subset=["收入","支出"],inplace=True)
    df["店铺"] = df["店铺"].str.lower()
    df["店铺"] = df["店铺"].str.replace(" ", "").str.strip()
    df["店铺"] = df["店铺"].apply(
        lambda x: x.replace("ambra京东海外旗舰店", "chiarabcaambra海外旗舰店").replace("loshi京东自营旗舰店", "loshi旗舰店").
            replace("manuka’scosmet海外旗舰店", "manukascosmet海外旗舰店").replace("chiara京东旗舰店", "chiarabcaambra海外旗舰店")
            .replace("dentyl京东旗舰店","dentylactive旗舰店").replace("unix优丽氏海外旗舰店","unix优丽氏海外品牌店").replace("乐丝拼多多店","乐丝旗舰店")
            .replace("卖家联合拼多多", "卖家联合海外专营店").replace("卖家联合全球购（有赞）", "卖家联合全球购").replace("时尚芭莎美妆店",
            "时尚芭莎美妆店主").replace("莎莎美妆品牌馆", "莎莎美妆品牌店"))

    df["收入"] = df["收入"].astype(float)
    df["支出"] = df["支出"].astype(float)
    df["收入-支出"] = df["收入-支出"].astype(float)
    group_df = df.groupby(["平台","店铺","月份"]).agg({"收入":"sum","支出":"sum","收入-支出":"sum"})
    group_df["数据来源"] = "财务账单"
    group_df = pd.DataFrame(group_df).reset_index()
    group_df.to_excel(default_dir+"/整理后的财务账单的收入和支出.xlsx",index=False)
    
    filename2 = default_dir+"/导出账单的收入和支出.xlsx"
    df2 = pd.read_excel(filename2)
    del df2["filename"]
    df2.drop_duplicates(inplace=True)
    df2.dropna(subset=["收入","支出"],inplace=True)
    df2["店铺"] = df2["店铺"].str.lower()
    df2["店铺"] = df2["店铺"].apply(
        lambda x: x.replace("loshi京东自营旗舰店", "loshi旗舰店"))
    df2["收入"] = df2["收入"].astype(float)
    df2["支出"] = df2["支出"].astype(float)
    df2["收入-支出"] = df2["收入-支出"].astype(float)
    group_df2 = df2.groupby(["平台","店铺","月份"]).agg({"收入":"sum","支出":"sum","收入-支出":"sum"})
    group_df2["数据来源"] = "导出账单"
    group_df2 = pd.DataFrame(group_df2).reset_index()
    group_df2.to_excel(default_dir+"/整理后的导出账单的收入和支出.xlsx",index=False)

def math_file():
    df1 = pd.read_excel("data/整理后的财务账单的收入和支出.xlsx",keep_default_na=False)
    df2 = pd.read_excel("data/整理后的导出账单的收入和支出.xlsx",keep_default_na=False)
    df1["店铺"] = df1["店铺"].str.lower()
    df1["店铺"] = df1["店铺"].str.replace(" ", "").str.strip()
    df2["店铺"] = df2["店铺"].str.lower()
    df1["店铺"] = df1["店铺"].apply(
        lambda x: x.replace("ambra京东海外旗舰店", "chiarabcaambra海外旗舰店").replace("loshi京东自营旗舰店", "loshi旗舰店").
            replace("manuka’scosmet海外旗舰店", "manukascosmet海外旗舰店").replace("chiara京东旗舰店", "chiarabcaambra海外旗舰店")
            .replace("dentyl京东旗舰店","dentylactive旗舰店").replace("unix优丽氏海外旗舰店","unix优丽氏海外品牌店").replace("乐丝拼多多店","乐丝旗舰店")
            .replace("卖家联合拼多多", "卖家联合海外专营店").replace("卖家联合全球购（有赞）", "卖家联合全球购").replace("时尚芭莎美妆店",
            "时尚芭莎美妆店主").replace("莎莎美妆品牌馆", "莎莎美妆品牌店"))
    df2["店铺"] = df1["店铺"].apply(
        lambda x: x.replace("loshi京东自营旗舰店", "loshi旗舰店"))
    df1.rename(columns={"收入":"财务：收入","支出":"财务：支出","收入-支出":"财务：收入-支出"},inplace=True)
    df2.rename(columns={"收入":"导出：收入","支出":"导出：支出","收入-支出":"导出：收入-支出"},inplace=True)
    del df1["数据来源"]
    del df2["数据来源"]
    df = pd.merge(df1,df2,how="outer",on=["平台","店铺","月份"])
    # df["数量差异（财务/订单）"] = df.apply(lambda  x:  (1-round((x["财务-订单数量"] / x["导出-订单数量"]),2))*100  ,axis=1 )
    # df["数量差异（财务/订单）"] = df.apply(lambda  x: "{}".format((1-round((x["财务-订单数量"] / x["导出-订单数量"]),2))*100),axis=1 )
    # df["数量差异（财务/订单）"] = df.apply(lambda  x: "{:0>2d}%".format((1-round((x["财务-订单数量"] / x["导出-订单数量"]),2))*100) ,axis=1 )

    # df["金额差异（财务/订单）"] = df.apply(lambda  x: "{}".format((1-round((x["财务-订单金额"] / x["导出-订单金额"]),2))*100) ,axis=1 )
    df["收入差异（财务/订单）"] = round((1-(df["财务：收入"] / df["导出：收入"])), 6)
    df["支出差异（财务/订单）"] = round((1-(df["财务：支出"] / df["导出：支出"])), 6)
    df["收入-支出差异（财务/订单）"] = round((1-(df["财务：收入-支出"] / df["导出：收入-支出"])), 8)

    # df["数量差异（财务/订单）"]=df["数量差异（财务/订单）"].apply(lambda x:  "" if x.find("nan")>=0  else x )
    # df["金额差异（财务/订单）"]=df["金额差异（财务/订单）"].apply(lambda x:  "" if x.find("nan")>=0  else x )

    print(df.to_markdown())
    df.to_excel("data/财务和导出账单的收入和支出差距.xlsx",index=False)

if __name__ == "__main__":
    print("开始:", time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()))
    # print("abc".find("x"))

    # caiwu_xushizhang()



    # ss=get_zhangjiali_shopname(r"/Volumes/IT审计处理需求/neworder-2019/天猫、淘宝/张佳丽/账单/Ceci茜茜美妆店支付宝201912_zip\20885314795456850156_201912_账务明细_1.csv")
    # print(ss)

    # sys.exit()

    combine_excel()
    groupby_amt()
    math_file()

    print("ok")