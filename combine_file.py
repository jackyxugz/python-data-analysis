# coding=utf-8

import sys
import os
import pandas as pd
import time
import os.path


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
            if ((path.find("~") < 0) and (path.find(".DS_Store") < 0)):  # 带~符号表示临时文件，不读取
                if len(filekey) > 0:
                    for key in filekey:
                        # print(path)
                        filename = "".join(path.split("\\")[-1:])
                        # print(filename)
                        if filename.find(key) >= 0:  # 只做文件名的过滤
                            _files.append(path)
                else:
                    _files.append(path)

    # print(_files)
    return _files


def read_excel(filename):
    print(filename)
    # try:
    if filename.find("xls") >= 0:
        temp_df = pd.read_excel(filename, dtype=str)
        for column_name in temp_df.columns:
            temp_df.rename(columns={column_name: column_name.replace(" ", "").replace("\n", "").strip()}, inplace=True)
    elif filename.find("csv") >= 0:
        try:
            temp_df = pd.read_csv(filename, dtype=str)
            for column_name in temp_df.columns:
                temp_df.rename(columns={column_name: column_name.replace(" ", "").replace("\n", "").strip()},
                               inplace=True)
        except Exception as e:
            temp_df = pd.read_csv(filename, dtype=str, encoding="gb18030")
            for column_name in temp_df.columns:
                temp_df.rename(columns={column_name: column_name.replace(" ", "").replace("\n", "").strip()},
                               inplace=True)
    else:
        print("不是xls或者csv文件！！")
        dict = {"filename": filename}
        temp_df = pd.DataFrame(dict, index=[0])

    # 下单时间	发货日期	付款日期	快递公司	快递单号	收货人姓名	平台站点	确认收货时间
    # temp_df["下单时间"] = temp_df["下单时间"].astype("datetime64[ns]")
    # temp_df["发货日期"] = temp_df["发货日期"].astype("datetime64[ns]")
    # temp_df["付款日期"] = temp_df["付款日期"].astype("datetime64[ns]")
    # temp_df["确认收货时间"] = temp_df["确认收货时间"].astype("datetime64[ns]")

    temp_df = temp_df.replace("\t", "", regex=True)  # 把订单号中的\t去掉
    temp_df["取号时间"] = temp_df["取号时间"].astype("datetime64[ns]")

    # temp_df["寄件日期"] = temp_df["寄件日期"].astype("datetime64[ns]")
    # if "运单号码" in temp_df.columns:
    #     temp_df.rename(columns={"运单号码":"运单号"},inplace=True)
    # if "序号" in temp_df.columns:
    #     del temp_df["序号"]

    # temp_df = temp_df.loc[temp_df["日期(Settlementdate)"].str.contains('-', na=False), :]
    print(temp_df.head(1).to_markdown())
    return temp_df


def get_all_files(rootdir, filekey):
    filelist = list_all_files(rootdir, filekey)
    # print(filelist)
    mySeries = pd.Series(filelist)
    df = pd.DataFrame(mySeries)
    df.columns = ["filename"]
    # df["filename"] = df["filename"].apply(lambda x: x.strip().replace(" ", ""))

    # print(df.to_markdown())
    print(df)
    count = len(df)
    global default_count
    default_count = count
    return df


def read_all_excel(rootdir, filekey):
    df_files = get_all_files(rootdir, filekey)
    for index, file in df_files.iterrows():
        if 'df' in locals().keys():  # 如果变量已经存在
            dd = read_excel(file["filename"])
            dd["filename"] = file["filename"]
            print("进度表：{}/{}  文件{}，行数{}".format(index + 1, df_files.shape[0], file["filename"], dd.shape[0]))
            df = df.append(dd)
        else:
            df = read_excel(file["filename"])
            df["filename"] = file["filename"]
            print("进度表：{}/{}  文件{}，行数{}".format(index + 1, df_files.shape[0], file["filename"], df.shape[0]))

    return df


def combine_excel():
    print('请选择要合并的文件目录（请确定所有excel的表头都是相同的哦！）:')
    # filedir=""
    filedir = input()

    print("你选择的路径是：", filedir)

    print("如果表头不是第一行，请输入需要跳过表头行数")
    skiprow = int(input())

    print("你要跳过的表头行数：", skiprow)

    global default_dir
    default_dir = filedir

    global default_skiprow
    default_skiprow = skiprow

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

    table = read_all_excel(filedir, filekey)

    del table["filename"]

    global default_filekey
    default_filekey = filekey

    return table


def xuguizhong():
    df = combine_excel()
    if 'df' in locals().keys():  # 如果变量已经存在
        # print(df.head(10).to_markdown())
        print(df.head(3))
        # df.to_clipboard(index=False)
        print(f"合并的总文件数量：{default_count}")
        print("正在处理，请耐心等候。。。")
        if len(df) > 500000:
            if len(default_filekey) == 0:
                df.to_csv(default_dir + "\合并表格.csv", index=False)
            else:
                df.to_csv(default_dir + "\合并表格_关键字_{}.csv".format(default_filekey), index=False)
        else:
            if len(default_filekey) == 0:
                df.to_excel(default_dir + "\合并表格.xlsx", index=False)
            else:
                df.to_excel(default_dir + "\合并表格_关键字_{}.xlsx".format(default_filekey), index=False)
        print("文件合并成功!")
        time.sleep(1)
        # byebye = input()
    else:
        print("不好意思，什么也没有做哦 :(")


if __name__ == "__main__":
    print("开始:", time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()))
    xuguizhong()
    print("程序执行完毕！")
