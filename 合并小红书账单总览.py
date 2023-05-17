import pandas as pd
import numpy as np
import os
import tabulate
import xlrd
import xlwt
import xlsxwriter
import openpyxl
import time


def list_all_file(rootdir, filekey_list):
    if len(filekey_list) > 0:
        filekey_list = filekey_list.replace(","," ")
        filekey = filekey_list.split(" ")
    else:
        filekey = ""

    _files = []
    list = os.listdir(rootdir)  # 列出文件夹下所有的目录与文件
    for i in range(0,len(list)):
        path = os.path.join(rootdir,list[i])
        if os.path.isdir(path):
            _files.extend(list_all_file(path,filekey_list))
        if os.path.isfile(path):
            if ((path.find("~") < 0) and (path.find(".DS_Store") < 0)):
                if len(filekey) > 0:
                    for key in filekey:
                        filename = "".join(path.split("\\")[-1:])
                        if filename.find(key) > 0:
                            _files.append(path)
                else:
                    _files.append(path)

    return _files

def read_excel(filename,sheet):
    print(filename)
    if filename.find("xls")>=0:
        df = pd.read_excel(filename,sheet)
    elif filename.find("csv")>=0:
        try:
            df = pd.read_csv(filename,sheet)
        except Exception as e:
            df = pd.read_csv(filename,sheet,encoding="gb18030")
    else:
        print("非xls和csv文件，不读取！")
        dict = {"filename":filename}
        df = pd.DataFrame(dict,index=[0])

    return df

def get_all_file(rootdir,filekey):
    filelist = list_all_file(rootdir,filekey)
    mySeries = pd.Series(filelist)
    df = pd.DataFrame(mySeries)
    df.columns = ["filename"]

    count = len(df)
    # global default_count
    # default_count = count
    #
    # global default_filelist
    # default_filelist = filelist

    return filelist,count

# def read_all_excel(rootdir, filekey):
#     df_files = get_all_file(rootdir, filekey)
#     for index, file in df_files.iterrows():
#         if "df" in locals().keys():
#             dd = read_excel(file["filename"])
#             dd["filename"] = file["filename"]
#             print("进度表：{}/{}  文件{}，行数{}".format(index+1,df_files.shape[0],file["filename"],dd.shape[0]))
#             # df = df.append(dd)
#             df = pd.concat([df,dd],sort=True)
#         else:
#             df = read_excel(file["filename"])
#             df["filename"] = file["filename"]
#             print("进度表：{}/{}  文件{}，行数{}".format(index+1,df_files.shape[0],file["filename"],df.shape[0]))
#
#     return df
#
# def combine_excel():
#     print('请选择要合并的文件目录（请确定所有excel的表头都是相同的哦！）:')
#     filedir = input()
#     print("你输入的路径是：",filedir)
#
#     global default_dir
#     default_dir = filedir
#
#     if len(filedir) == 0:
#         print("你没有输入任何目录！")
#         exit()
#
#     print('请输入要合并的文件名称或者后缀匹配符:（比如都必须包含 序时账 三个字，那么就请输入  "序时账" ， 不输入就表示所有excel都要合并！）')
#     filekey = input()
#
#     if len(filekey) == 0:
#         print("你没有输入需要筛选的关键字！")
#         filekey = ""
#
#     print("你希望在'{}'目录下找到所有的包含“{}”文件，然后合并。".format(filedir, filekey))
#
#     table = read_all_excel(filedir, filekey)
#
#     return table

def result_out():
    print('请选择要合并的文件目录（请确定所有excel的表头都是相同的哦！）:')
    filedir = input()
    print("你输入的路径是：", filedir)

    global default_dir
    default_dir = filedir

    if len(filedir) == 0:
        print("你没有输入任何目录！")
        exit()

    print('请输入要合并的文件名称或者后缀匹配符:（比如都必须包含 序时账 三个字，那么就请输入  "序时账" ， 不输入就表示所有excel都要合并！）')
    filekey = input()

    if len(filekey) == 0:
        print("你没有输入需要筛选的关键字！")
        filekey = ""

    print("你希望在'{}'目录下找到所有的包含“{}”文件，然后合并。".format(filedir, filekey))

    # print('请输入要合并的分表！')
    # sheet_name = input()
    sheet_name = "总览"
    #
    # if len(sheet_name) == 0:
    #     print("你没有输入需要筛选的关键字！")
    #     sheet_name = ""

    filelist,count = get_all_file(filedir,filekey)

    print("总共合并{}文件".format(count))
    index = 0
    writer = pd.ExcelWriter(default_dir + os.sep + "合并文件.xlsx", engine="xlsxwriter")
    for i in filelist:
        print("正在读取{}文件".format(i))
        # if "df" in locals().keys():
        #     dd = read_excel(i,sheet_name)
        #     print("进度表：{}/{}  文件{}，行数{}".format(index + 1, count, i, dd.shape[0]))
        #     # df = df.append(dd)
        #     df = pd.concat([df, dd], sort=True)
        # else:
        #     df = read_excel(i,sheet_name)
        #     print("进度表：{}/{}  文件{}，行数{}".format(index + 1, count, i, df.shape[0]))
        df = read_excel(i, sheet_name)
        print("进度表：{}/{}  文件{}，行数{}".format(index + 1, count, i, df.shape[0]))
        index += 1
        sheet = "".join(("".join(i.split(os.sep)[-1:])).split(".")[:1])
        print("sheet表：{}".format(sheet))

        df.to_excel(writer, sheet_name=sheet, index=False)

    writer.save()

    print("合并结束！共合并{}个文件!".format(count))

    byebye = input()

    return df

if __name__ == "__main__":

    result_out()

