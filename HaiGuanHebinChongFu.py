# coding=utf-8

import sys
import os
import tkinter as tk

import pandas as pd
import time
import os.path
from tkinter import filedialog
import warnings
import tabulate

warnings.filterwarnings("ignore")
root = tk.Tk()
root.withdraw()


# 列出所有文件
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
            # 循环嵌套
            _files.extend(list_all_files(path, filekey_list))
        if os.path.isfile(path):
            if path.find("~") < 0:  # 带~符号表示临时文件，不读取
                if len(filekey) > 0:
                    for key in filekey:
                        # print(path)
                        filename = os.path.split(path)[1]
                        # filename = "".join(path.split("\\")[-1:])
                        if filename.find(key) >= 0:  # 只做文件名的过滤
                            _files.append(path)
                else:
                    _files.append(path)

    # print(_files)
    # 返回一个文件列表 list
    return _files


def get_all_files(rootdir, filekey):
    filelist = list_all_files(rootdir, filekey)
    mySeries = pd.Series(filelist)
    df = pd.DataFrame(mySeries)
    df.columns = ["filename"]
    # df["filename"] = df["filename"].apply(lambda x: x.strip().replace(" ", ""))

    print("\n要合并的文件总数:", df.shape[0])
    # print("去重前文件数:", df.shape[0])
    df["shortname"] = df["filename"].apply(lambda x: os.path.basename(x))
    # df.drop_duplicates(subset=["shortname"], keep="first",inplace=True)
    # print("去重后文件数:", df.shape[0])
    print(df.to_markdown())
    return df


def combine_excel():
    print('\n请选择要合并的文件目录（请确定所有excel的表头都是相同的哦！）:')
    filedir = ""
    filedir = filedialog.askdirectory()  # 获取文件夹
    print("\n你选择的路径是：", filedir)

    global default_dir
    default_dir = filedir

    if len(filedir) == 0:
        print("你没有输入任何目录 :(")
        sys.exit()
        return

    print('\n请输入要合并的文件名称或者后缀匹配符:（不输入表示所有文件都要合并！）')
    filekey = input()

    if len(filedir) == 0:
        print("你没有输入任何关键词 :(")
        filekey = ''
        # sys.exit()
        # return

    print("你希望在'{}'目录下找到所有的包含“{}”文件，然后合并。".format(filedir, filekey))

    df_box = read_all_excel(filedir, filekey)

    return df_box


def read_all_excel(rootdir, filekey):
    # title_key,bottom_key
    df_files = get_all_files(rootdir, filekey)
    df_box = []
    df_error = []
    dd_shape_sum = 0
    dd_shape = 0
    for index, file in df_files.iterrows():
        # 读取文件
        filename = file["filename"]
        if os.path.splitext(filename)[1].find("xls") >= 0:
            try:
                dd = pd.read_excel(filename, sheet_name="税单详情", skiprows=1,
                                   dtype="str")  # 实际应用时要根据excel实际情况调整sheet_name
                dd_shape = dd.shape[0]
            except Exception as err:
                df_error.append(filename)
        else:
            encoding_list = ["ansi", "utf-8", "gbk", "gb18030", "ansi"]

            try:
                print("encoding=", encoding_list[0])
                dd = pd.read_csv(filename, encoding=encoding_list[0], dtype=str, error_bad_lines=False,
                                 engine="python")
                dd_shape = dd.shape[0]
            except Exception as err:
                try:
                    print("encoding=", encoding_list[1])
                    dd = pd.read_csv(filename, encoding=encoding_list[1], dtype=str, error_bad_lines=False,
                                     engine="python")
                except Exception as err:
                    try:
                        print("encoding=", encoding_list[2])
                        dd = pd.read_csv(filename, encoding=encoding_list[2], dtype=str, error_bad_lines=False,
                                         engine="python")
                    except Exception as err:
                        try:
                            print("encoding=", encoding_list[3])
                            dd = pd.read_csv(filename, encoding=encoding_list[3], dtype=str,
                                             error_bad_lines=False,
                                             engine="python")
                        except Exception as err:
                            try:
                                print("encoding=", encoding_list[4])
                                dd = pd.read_csv(filename, encoding=encoding_list[4], dtype=str,
                                                 error_bad_lines=False, engine="python")
                            except Exception as err:
                                print(filename, " 异常:", err)
                                return 0

        if "dd" in vars():
            if "filename" in dd.columns:
                del dd["filename"]

            dd["filename"] = file["filename"]
            for col in dd.columns:
                if col.find('Unnamed') >= 0:
                    del dd[col]

            if "df_sum" not in vars():
                df_sum = dd
            else:
                df_sum = df_sum.append(dd)

            print("\n进度表：{}/{}  文件名{}，文件行数：{}".format(index + 1, df_files.shape[0], file["filename"], dd.shape[0]))
            print("合并后的数据行数:", df_sum.shape[0])
        dd_shape_sum += dd_shape

    df_box.append(df_sum)

    if dd_shape_sum != df_sum.shape[0]:
        print("\n合并前记录数({})与合并后记数({})不相等，请检查原文件！".format(dd_shape_sum, df_sum.shape[0]))

    if df_error != []:
        print("\n这些文件读取报错:")
        print(df_error)

    return df_box


def batch_convert_save():
    # 批量转换，保存
    # 合并所有文件
    df_box = combine_excel()
    index = 0
    for df in df_box:
        print("\n输出第{}个表格,记录数:{}".format(index + 1, df.shape[0]))
        print(df.head(10).to_markdown())
        # table_header = ['电子税单编号', '清单编号', '缴款书编号', '税单状态', '应征关税（元）', '应征消费税（元）',
        #                 '应征增值税（元）', '电商企业代码', '订单编号', '物流企业代码', '运单编号', '担保企业代码', '备注',
        #                 '生成时间']
        #
        # print(tabulate(df.head(10), headers=table_header,
        #                tablefmt='pipe'))  # “plain”,“simple”,“github”,“grid”,“fancy_grid”,“pipe”,“orgtbl”,“jira”,“presto”,“psql”,“rst”,“mediawiki”,“moinmoin”,“youtrack”,“html”,“latex”,“latex_raw”,“latex_booktabs”,“textile”
        # # print(df.head(5).to_markdown())

        for i in range(0, int(df.shape[0] / 500000) + 1):
            print("\n存储分页{}个表格，记录数:{} to:{}".format(i + 1, i * 500000, (i + 1) * 500000))
            df.iloc[i * 500000:(i + 1) * 500000].to_excel(default_dir + "\合并后第{}-{}个表格_{}.xlsx".format(index + 1, i + 1,
                                                                                                       time.strftime(
                                                                                                           "%Y-%m-%d %H:%M:%S",
                                                                                                           time.localtime()).replace(
                                                                                                           " ",
                                                                                                           "").replace(
                                                                                                           "-",
                                                                                                           "").replace(
                                                                                                           ":", "")),
                                                          index=False)
        index = index + 1


if __name__ == "__main__":
    print("开始:", time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()))
    try:
        batch_convert_save()
        print("\n合并结束！")
    except Exception as err:
        print(err)

    input('按任意键退出......')
