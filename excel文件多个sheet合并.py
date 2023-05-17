# __coding=utf8__
# /** 作者：zengyanghui **/
import pandas as pd
import os
import time
import tabulate
import os.path
import xlrd
import xlwt
import xlsxwriter
import win32api
import win32ui
import win32con
import win32com


def read_all_sheet():
    file_path = open_file()
    print(file_path)
    default_path = os.sep.join(file_path.split(os.sep)[:-1])
    # out_file = "".join("".join(file_path.split(os.sep)[-1:]).split(".")[:1])
    if file_path.find(".xls")>=0:
        df = pd.read_excel(file_path,sheet_name=None,dtype=str)
    else:
        try:
            df = pd.read_csv(file_path, sheet_name=None, dtype=str)
        except Exception as e:
            df = pd.read_csv(file_path, sheet_name=None, dtype=str, encoding="gb18030")
    # print(df.head(5).to_markdown())
    print(list(df))
    sheet_list = list(df)

    df1 = None
    for sheet in sheet_list:
        if file_path.find(".xls") >= 0:
            df = pd.read_excel(file_path, sheet_name=sheet, dtype=str)
        else:
            try:
                df = pd.read_csv(file_path, sheet_name=sheet, dtype=str)
            except Exception as e:
                df = pd.read_csv(file_path, sheet_name=sheet, dtype=str, encoding="gb18030")
        # df = pd.read_excel(file_path,sheet_name=sheet,dtype=str)
        df["filename"] = file_path
        df["sheet"] = sheet
        # print(len(df))
        # print(df.head(5).to_markdown())
        if df1 is None:
            df1 = df
        else:
            print(f"正在读取“{sheet}”分表")
            print(f"\n{df.head().to_markdown()}")
            df1= pd.concat([df1,df])

    print("\n\n")
    print(f"合并的分表：{sheet_list}")
    print(f"合并后总行数：{len(df1)}")
    print(df1.head(5).to_markdown())
    print("\n\n合并结束！！！")

    df1.to_excel(default_path + os.sep + "合并所有sheet.xlsx")

    print("\n输出的文件路径：")
    print(default_path + os.sep + "合并所有sheet.xlsx")
    byebye = input()


def open_file():
    dlg = win32ui.CreateFileDialog(1)  # 1表示打开文件对话框
    dlg.SetOFNInitialDir('D:/')  # 设置打开文件对话框中的初始显示目录
    dlg.DoModal()
    filename = dlg.GetPathName()  # 获取选择的文件名称
    print("filename=",filename)
    print("read ok")
    return filename


if __name__ == "__main__":
    print("开始:", time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()))

    read_all_sheet()

    print("结束:", time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()))