# coding=utf-8

import sys
import os
import pandas as pd
import time
import os.path
import xlrd
import xlwt
import math

import tabulate
import operator

# import win32api
# import win32ui
# import win32con
#
# import win32com
# from win32com.shell import shell
from tkinter import filedialog


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


file_columns_list=[]

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
                        if filename.find(key) > 0:  # 只做文件名的过滤
                            _files.append(path)
                else:
                    _files.append(path)

    # print(_files)
    return _files


def read_excel(filename):
    if os.path.splitext(filename)[1].find("xls")>=0:
        temp_df = pd.read_excel(filename,dtype=str)
    else:
        temp_df = pd.read_csv(filename,encoding="gbk", dtype=str,error_bad_lines=False).reset_index()  # ,decodeing="utf-8"

    temp_df.fillna("", inplace=True)
    print("抽查：",filename)
    print(temp_df.head(3).to_markdown())  #
    skiptop = 0
    skipbottom = 0
    skipbottom1 = 0
    skipbottom2 = 0
    shopname=""

    # print("忽略头部")

    for i in range(0, min(temp_df.shape[0]-1,15) ):
        row = ""
        for j in range(0, len(temp_df.columns)-1):
            # print(df.iloc[i, j] )
            row = row + "|" + str(temp_df.iloc[i, j])
            # if row.find("金额") >= 0:

        if row.find("店铺名称")>=0:
            shopname= row.replace("店铺名称：","").replace("|","")

        # 用关键词定位标题行
        if ((row.find("项目") >= 0) or (row.find("金额") >= 0)):
            skiptop = i + 1

        print(row)

    print("忽略尾部")
    large_row=""
    small_row="xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx"
    for i in range(1, min(temp_df.shape[0]-1,15)):
        row = ""
        for j in range(0, len(temp_df.columns)):
            # print(df.iloc[i, j] )
            row = row + "|" + str(temp_df.iloc[temp_df.shape[0]-i, j])

        if len(row)>len(large_row):
            large_row=row
        if len(row)<len(small_row):
            small_row=row
        # 如果这一行很短，说明这是合计行
        if len(small_row)<=len(large_row)/2:
            print("发现疑似合计行：",row)
            skipbottom1 = i-1
            break
        # print("跟踪1 small_row：",small_row)
        # print("跟踪2 large_row：",large_row)
        # print("跟踪3 row：",row)

        # print(row)

    for i in range(1, min(temp_df.shape[0] - 1, 15)):
        row = ""
        for j in range(0, len(temp_df.columns)):
            # print(df.iloc[i, j] )
            row = row + "|" + str(temp_df.iloc[temp_df.shape[0] - i, j])

        # 用关键词定位合计行
        if ((row.find("合计") > 0) | (row.find("总收入") > 0)):
            print("发现合计行：", row)
            skipbottom2 = i
            break

    skipbottom=max(skipbottom1,skipbottom2)
    print("末尾行数1:{}，行数2:{}，结果:{}".format(skipbottom1,skipbottom2,skipbottom))

        # print(row)

    #  按照正确的列名重新读取csv文件
    if os.path.splitext(filename)[1].find("xls")>=0:
        temp_df = pd.read_excel(filename,dtype=str, skiprows=skiptop)
    else:
        temp_df = pd.read_csv(filename, encoding="gbk", dtype=str, error_bad_lines=False, skiprows=skiptop) # ,skipbottom=skipbottom

    temp_df.fillna("",inplace=True)
    temp_df=temp_df[:-skipbottom]

    for col in temp_df.columns:
        # print("col=",col)
        if col.find("Unnamed:")>=0:
            del temp_df[col]

    # temp_df["shopname"]=shopname

    # print(temp_df.dtypes)
    print("查看结果:")
    print("文件： {} skiptop:{} skipbottom:{} 行数：{}".format( filename,   skiptop ,skipbottom,temp_df.shape[0]))
    print(temp_df.head(3).to_markdown())  #
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
    return df


def read_all_excel(rootdir, filekey):
    df_files = get_all_files(rootdir, filekey)
    df_box=[]
    for index, file in df_files.iterrows():

        dd = read_excel(file["filename"])
        dd["filename"] = file["filename"]
        print("进度表：{}/{}  文件{}，行数{}".format(index + 1, df_files.shape[0], file["filename"], dd.shape[0]))

        # 如果行数为0，则忽略，不需要关心格式
        if dd.shape[0]==0:
            print("忽略空表：", file["filename"])
        elif  file["filename"].find("错误")>0:
            print("忽略错误的表格：", file["filename"])
        else:
        # if dd.shape[0] > 0:
            # 将列名转换成list
            if_exist=False
            iseq=0
            dd_columns=dd.columns.to_list()
            for r in file_columns_list:
                # 如果相同，已经存在
                # print("对比1:",dd_columns)
                # print("对比2:",r[1:])

                if operator.eq(dd_columns,r[1:]):
                # if dd_columns==r[1:]:
                    if_exist=True
                    break
                iseq=iseq+1

            if if_exist:
                # print("模板已经存在！：", file["filename"])
                # print("模板1:", dd_columns)
                # print("模板2:",iseq)
                # for t in file_columns_list:
                #     print(t)

                if iseq == 0:
                    df0 = df0.append(dd)
                elif iseq == 1:
                    df1 = df1.append(dd)
                elif iseq == 2:
                    df2 = df2.append(dd)
                elif iseq == 3:
                    df3 = df3.append(dd)
                elif iseq == 4:
                    df4= df4.append(dd)
                elif iseq == 5:
                    df5 = df5.append(dd)
                elif iseq == 6:
                    df6 = df6.append(dd)
                elif iseq == 7:
                    df7 = df7.append(dd)
                elif iseq == 8:
                    df8 = df8.append(dd)
                elif iseq == 9:
                    df9 = df9.append(dd)
                elif iseq == 10:
                    df10 = df10.append(dd)
                elif iseq == 11:
                    df11 = df11.append(dd)
                elif iseq == 12:
                    df12 = df12.append(dd)
                elif iseq == 13:
                    df13 = df13.append(dd)


            else:
            #     print("新增字段模板：",file["filename"])
            #     print("模板1:", dd_columns)
                print("模板2:")
                for t in file_columns_list:
                    print(t)

                dd_columns.insert(0,file["filename"])
                file_columns_list.append(dd_columns)

                iseq=len(file_columns_list)-1
                if iseq==0:
                    df0 = dd
                elif iseq==1:
                    df1 = dd
                elif iseq==2:
                    df2 = dd
                elif iseq==3:
                    df3 = dd
                elif iseq==4:
                    df4 = dd
                elif iseq==5:
                    df5 = dd
                elif iseq==6:
                    df6 = dd
                elif iseq==7:
                    df7 = dd
                elif iseq==8:
                    df8 = dd
                elif iseq==9:
                    df9 = dd
                elif iseq==10:
                    df10 = dd
                elif iseq==11:
                    df11 = dd
                elif iseq==12:
                    df12= dd
                elif iseq==13:
                    df13 = dd

                # print(file["filename"], dd.shape[0])
                # print("进度表：{}/{}  文件{}，行数{}".format(index + 1, df_files.shape[0], file["filename"], dd.shape[0]))

    if 'df0' in vars():
        df_box.append(df0)
    if 'df1' in vars():
        df_box.append(df1)
    if 'df2' in vars():
        df_box.append(df2)
    if 'df3' in vars():
        df_box.append(df3)
    if 'df4' in vars():
        df_box.append(df4)
    if 'df5' in vars():
        df_box.append(df5)
    if 'df6' in vars():
        df_box.append(df6)
    if 'df7' in vars():
        df_box.append(df7)
    if 'df8' in vars():
        df_box.append(df8)
    if 'df9' in vars():
        df_box.append(df9)
    if 'df10' in vars():
        df_box.append(df10)
    if 'df11' in vars():
        df_box.append(df11)
    if 'df12' in vars():
        df_box.append(df12)
    if 'df13' in vars():
        df_box.append(df13)

    print("最终的字段列表：")
    print(file_columns_list)
    print("字段模板共有：{} 个".format(len(file_columns_list)))

    return df_box
    # return df0,df1,df2,df3,df4,df5,df6,df7,df8,df9,df10,df11,df12,df13


def combine_excel():
    print('请选择要合并的文件目录（请确定所有excel的表头都是相同的哦！）:')
    filedir=""
    # filedir = input()
    # myTuple = shell.SHBrowseForFolder(0, None, "", 64)
    # try:
    #     path = shell.SHGetPathFromIDList(myTuple[0])
    # except:
    #     print("你没有输入任何目录 :(")
    #     sys.exit()
    #     return
    #
    # filedir=path.decode('ansi')

    filedir=filedialog.askdirectory()  # 获取文件夹
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

    df_box = read_all_excel(filedir, filekey)
    # df0,df1,df2,df3,df4,df5,df6,df7,df8,df9,df10,df11,df12,df13

    return df_box


def district_save():
    df_box  = combine_excel()
    # df0,df1,df2,df3,df4,df5,df6,df7,df8,df9,df10,df11,df12,df13
    index=0
    for df in df_box:
        print("第{}个表格,记录数:{}".format(index,df.shape[0]))
        print(df.head(10).to_markdown())
        # df.to_excel(r"work/合并表格_test.xlsx")

        for i in range(0,int(df.shape[0]/500000)+1):
            print("存储分页：{}  from:{} to:{}".format(i ,i*500000,(i+1)*500000))
            # df.iloc[i*500000:(i+1)*500000].to_csv(default_dir + "\{}_合并表格_{}.csv".format(index,i))
            df.iloc[i*500000:(i+1)*500000].to_excel(default_dir + "\{}_合并表格_{}.xlsx".format(index,i))

            # df[i*500000:(i+1)*500000].to_csv( "work\{}_合并表格_{}.csv".format(index,i))

        index=index+1

    #
    # # if not (df0 is None):
    # if 'df0' in vars():
    #     for i in range(0,int(df0.shape[0]/500000)):
    #         df0[i:i*500000].to_csv(default_dir + r"\合并表格_0_{}.csv".format(i))
    #
    # # if not (df1 is None):
    # if 'df1' in locals().keys():
    #     for i in range(0, int(df1.shape[0] / 500000)):
    #         df1[i:i * 500000].to_csv(default_dir + r"\合并表格_1_{}.csv".format(i))
    #
    # # if not (df2 is None):
    # if 'df2' in locals().keys():
    #     for i in range(0, int(df2.shape[0] / 500000)):
    #         df2[i:i * 500000].to_csv(default_dir + r"\合并表格_2_{}.csv".format(i))
    #
    # # if not (df3 is None):
    # if 'df3' in locals().keys():
    #     for i in range(0, int(df3.shape[0] / 500000)):
    #         df3[i:i * 500000].to_csv(default_dir + r"\合并表格_3_{}.csv".format(i))
    #
    # # if not (df4 is None):
    # if 'df4' in locals().keys():
    #     for i in range(0, int(df4.shape[0] / 500000)):
    #         df4[i:i * 500000].to_csv(default_dir + r"\合并表格_4_{}.csv".format(i))
    #
    # # if not (df5 is None):
    # if 'df5' in locals().keys():
    #     for i in range(0, int(df5.shape[0] / 500000)):
    #         df5[i:i * 500000].to_csv(default_dir + r"\合并表格_5_{}.csv".format(i))
    #
    # # if not (df6 is None):
    # if 'df6' in locals().keys():
    #     for i in range(0, int(df6.shape[0] / 500000)):
    #         df6[i:i * 500000].to_csv(default_dir + r"\合并表格_6_{}.csv".format(i))
    #
    # # if not (df7 is None):
    # if 'df7' in locals().keys():
    #     for i in range(0, int(df7.shape[0] / 500000)):
    #         df7[i:i * 500000].to_csv(default_dir + r"\合并表格_7_{}.csv".format(i))
    #
    # # if not (df8 is None):
    # if 'df8' in locals().keys():
    #     for i in range(0, int(df8.shape[0] / 500000)):
    #         df8[i:i * 500000].to_csv(default_dir + r"\合并表格_8_{}.csv".format(i))
    #
    #
    # # if not (df9 is None):
    # if 'df9' in locals().keys():
    #     for i in range(0, int(df9.shape[0] / 500000)):
    #         df9[i:i * 500000].to_csv(default_dir + r"\合并表格_9_{}.csv".format(i))
    #
    # # if not (df10 is None):
    # if 'df10' in locals().keys():
    #     for i in range(0, int(df10.shape[0] / 500000)):
    #         df10[i:i * 500000].to_csv(default_dir + r"\合并表格_10_{}.csv".format(i))
    #
    # # if not (df11 is None):
    # if 'df11' in locals().keys():
    #     for i in range(0, int(df11.shape[0] / 500000)):
    #         df11[i:i * 500000].to_csv(default_dir + r"\合并表格_11_{}.csv".format(i))
    #
    # # if not (df12 is None):
    # if 'df12' in locals().keys():
    #     for i in range(0, int(df12.shape[0] / 500000)):
    #         df12[i:i * 500000].to_csv(default_dir + r"\合并表格_12_{}.csv".format(i))
    #
    # # if not (df13 is None):
    # if 'df13' in locals().keys():
    #     for i in range(0, int(df13.shape[0] / 500000)):
    #         df13[i:i * 500000].to_csv(default_dir + r"\合并表格_13_{}.csv".format(i))

    print("生成完毕，现在关闭吗？yes/no")
    byebye = input()
    print('bybye:', byebye)

def read_xiaohongshu(filename):


    # df=read_excel(r"/Users/kfz/PycharmProjects/Megacombine/data/caiwu/tandongmei/京东海外/京东Dicora UrbanFit海外旗舰店/1 妥投/结算单162159672妥投结算数据.csv")
    # df=read_excel(r"/Users/lichunlei/Library/Containers/com.tencent.xinWeChat/Data/Library/Application Support/com.tencent.xinWeChat/2.0b4.0.9/30b19085e36d0f4274a8809b5f07ae0a/Message/MessageTemp/f32808c113e684ade24f9fd133ae240a/File/2019各平台数据原表/小红书/小红书Dentyl Active旗舰店/小红书Dentyl Active旗舰店201901-账单.xlsx")
    df=read_excel(filename)
    print("测试")
    print(df)

    shopname=df.iloc[0,3]
    print("店铺名称==",shopname)

    del df["shopname"]

    # df=df.stack()
    df.columns=["项目","收支类型","金额"]
    df["金额"]=df.apply(lambda x: "{}{}".format(x["收支类型"],x["金额"].strip()) ,axis=1)
    del df["收支类型"]

    df=df.reset_index(drop=True)

    print(df.to_markdown())

    # 行列转置
    df=df.set_index("项目").T
    # 删除索引
    df = df.reset_index(drop=True)

    df["shopname"]=shopname

    print("最终结果:")
    print(df)

if __name__ == "__main__":
    print("开始:", time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()))
    # print("abc".find("x"))

    district_save()

    # filename=r"/Users/lichunlei/Library/Containers/com.tencent.xinWeChat/Data/Library/Application Support/com.tencent.xinWeChat/2.0b4.0.9/30b19085e36d0f4274a8809b5f07ae0a/Message/MessageTemp/f32808c113e684ade24f9fd133ae240a/File/2019各平台数据原表/小红书/小红书Dentyl Active旗舰店/小红书Dentyl Active旗舰店201901-账单.xlsx"
    # read_xiaohongshu(filename)

    # df=df.pivot_table(index=[0],columns=["","","","","",""],aggfunc="first").reset_index


    # df = df.stack().reset_index()



    # print(df.to_markdown())



    print("ok")

