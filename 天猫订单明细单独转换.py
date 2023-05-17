# __coding=utf8__
# /** 作者：zengyanghui **/

# __coding=utf8__
# /** 作者：zengyanghui **/
# __coding=utf8__
# /** 作者：zengyanghui **/

import sys
import os
import pandas as pd
import operator
# 显示所有列
pd.set_option('display.max_columns', None)
# 显示所有行
pd.set_option('display.max_rows', None)
# 设置value的显示长度为100，默认为50
pd.set_option('max_colwidth', 200)

import numpy as np
# from datetime import datetime
import datetime
import time
import os.path
import xlrd
import xlwt
import pprint
import math
import tabulate
import logging
# import Tkinter
import win32api
import win32ui
import win32con


def open_file():
    dlg = win32ui.CreateFileDialog(1)  # 1表示打开文件对话框
    dlg.SetOFNInitialDir('D:/')  # 设置打开文件对话框中的初始显示目录
    dlg.DoModal()
    filename = dlg.GetPathName()  # 获取选择的文件名称
    print("filename=",filename)
    print("read ok")
    return filename


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
                        # print("文件名:",filename)

                        key = key.replace("！", "!")

                        if key.find("!") >= 0:
                            # print("反向选择:",key)
                            if filename.find(key.replace("!", "")) >= 0:  # 此文件不要读取
                                # print("{} 不应该包含 {}，所以剔除:".format(filename,key ))
                                pass
                        elif filename.find(key) > 0:  # 只做文件名的过滤
                            _files.append(path)

                else:
                    _files.append(path)

    # print(_files)
    return _files


def read_excel(filename):
    print(filename)
    if ((filename.find("天猫") >= 0) | (filename.find("淘宝") >= 0)):
        # try:
        print("定位1")
        if filename.find("csv")>=0:
            df = pd.read_csv(filename, dtype=str, encoding="gb18030")
        else:
            df = pd.read_excel(filename, dtype=str)
        print(df.head().to_markdown())
        for column_name in df.columns:
            df.rename(columns={column_name: column_name.replace(" ", "").replace(" ", "").replace("\n", "").strip()}, inplace=True)

        # df["订单编号"] = df["订单编号"].str.replace("'", "")
        # df["订单编号"] = df["订单编号"].astype(str)


        print(df.head().to_markdown())
        # print(df["店铺名称"][1])
        if filename.find("天猫")>=0:
            plat = "TMALL"
            plat_chs = "天猫"
        elif filename.find("淘宝")>=0:
            plat = "TAOBAO"
            plat_chs = "淘宝"

        df2 = pd.DataFrame()
        df2["TITLE"] = df["标题"]
        if "SKU" in df.columns:
            df2["SKU"] = df["SKU"]
        else:
            df2["SKU"] = df["商家编码"]
        df2["PRICE"] = df["价格"]
        df2["PAYMENT"] = df["买家实际支付金额"]
        df2["PIC_PATH"] = ""
        df2["NUM"] = df["购买数量"]
        df2["PAY_TIME"] = df["订单付款时间"]
        df2["LOGISTICS_COMPANY"] = ""
        df2["INVOICE_NO"] = ""
        df2["TID"] = df["主订单编号"]
        df2["SKU_PROPERTIES_NAME"] = df["商品属性"]
        df2["OID"] = ""
        df2["REFUND_STATUS"] = ""
        df2["CONSIGN_TIME"] = ""
        df2["CREATED"] = df["订单创建时间"]
        df2 = df2.apply(lambda x: x.astype(str).str.replace("null", ""))
        df2["PAYMENT"] = df2["PAYMENT"].astype(float)
        df2["PRICE"] = df2["PRICE"].astype(float)
        df2["TID"] = df2.apply(lambda x:x["TID"] if ((len(x["TITLE"])>3)|(len(x["SKU"])>3)|(len(x["LOGISTICS_COMPANY"])>3)|(len(x["INVOICE_NO"])>3)) else "nan",axis=1)
        df2 = df2[~df2["TID"].str.contains("nan")]

        print("天猫-订单明细")
        print(df2.head(5).to_markdown())

    return df2


def get_baidushop(id,filename,type,df):
    # df = pd.read_pickle("data/百度-商品ID与店铺名称关系0(3)(2).pkl")
    shop_name = ""
    # print(df.to_markdown())
    if type == 1:
        if df[df["商品ID"].str.contains(id)].shape[0] > 0:
            # print("test1")
            # print(id)
            df = df[df["商品ID"].str.contains(id)]
            shop_name = df.iloc[0]["店铺名称"]
        else:
            # print("test2")
            # print(filename)
            shop = "".join(filename.split(os.sep)[-1:])
            # print(shop)
            if ((shop.find("广分电商") >= 0) | (shop.find("广西电商") >= 0)):
                shop_name = "-".join(shop.split("-")[:2])
                # print(shop_name)
            else:
                shop_name = "-".join(shop.split("-")[:1])
                # print(shop_name)
    else:
        if df[(df["店铺名称"].str.contains(id))].shape[0] > 0:
            # print("test2")
            df = df[df["店铺名称"].str.contains(id)]
            shop_name = df.iloc[0]["店铺代码"]
        else:
            shop_name = ""

    return shop_name


def error(self,filename):
    logging.basicConfig(filename=default_dir + "\错误日志.log",
                        format=f'[%(asctime)s-%(filename)s-%(levelname)s:%(message)s:{filename}]', level=logging.DEBUG,
                        filemode='a', datefmt='%Y-%m-%d%I:%M:%S %p')

    logging.error("这是一条error信息的打印")
    # logging.info("这是一条info信息的打印")
    # logging.warning("这是一条warn信息的打印")
    # logging.debug("这是一条debug信息的打印")


def taobao_shop(plat,filename,type):
    if plat == "TAOBAO":
        if filename.find("惠优购企业店") >= 0:
            if type == 1:  # 类型1返回店铺名称，类型2返回店铺代码
                shop = "惠优购企业店"
                return shop
            elif type == 2:
                shop = "hyg"
                return shop
        elif filename.find("麦凯莱品牌自营店") >= 0:
            if type == 1:  # 类型1返回店铺名称，类型2返回店铺代码
                shop = "麦凯莱品牌自营店"
                return shop
            elif type == 2:
                shop = "mklppzydzd"
                return shop
        elif filename.find("米兰站美妆") >= 0:
            if type == 1:  # 类型1返回店铺名称，类型2返回店铺代码
                shop = "米兰站美妆"
                return shop
            elif type == 2:
                shop = "mlzmz"
                return shop
        elif filename.find("前男友美妆") >= 0:
            if type == 1:  # 类型1返回店铺名称，类型2返回店铺代码
                shop = "前男友美妆"
                return shop
            elif type == 2:
                shop = "qnymz"
                return shop
        elif filename.find("时尚芭莎美妆店主") >= 0:
            if type == 1:  # 类型1返回店铺名称，类型2返回店铺代码
                shop = "时尚芭莎美妆店主"
                return shop
            elif type == 2:
                shop = "bsmzg"
                return shop
        elif filename.find("莎莎美妆品牌店") >= 0:
            if type == 1:  # 类型1返回店铺名称，类型2返回店铺代码
                shop = "莎莎美妆品牌店"
                return shop
            elif type == 2:
                shop = "ssmzppd"
                return shop
        elif filename.find("ceci茜茜美妆") >= 0:
            if type == 1:  # 类型1返回店铺名称，类型2返回店铺代码
                shop = "Ceci茜茜美妆"
                return shop
            elif type == 2:
                shop = "cecixxmz"
                return shop


def taobao_amount(endtime,amount,peo):
    # print(amount)
    if amount.find("nan") >= 0:
        if peo.find("否")<0:
            # print(f"定位1.1-{peo}")
            return peo
        elif endtime.find("null") < 0:
            # print(f"定位1.2-{endtime}")
            return endtime
        else:
            # print(f"定位1.3-{amount}")
            return amount
    elif amount.find("-") >= 0:
        if peo.find("否") < 0:
            # print(f"定位2.1-{peo}")
            return peo
        elif endtime.find("null") < 0:
            # print(f"定位2.2-{endtime}")
            return endtime
        else:
            # print(f"定位2.3-{amount}")
            return amount
    else:
        # print(f"定位3-{amount}")
        return amount


def taobao_endtime(shop,endtime,amount):
    if endtime.find("null")>=0:
        if amount.find("-")>=0:
            return amount
        else:
            return endtime
    elif str(endtime).isnumeric():
        if shop.find("-")>=0:
            return shop
        else:
            return endtime
    elif shop.find("nan")>=0:
        return shop
    else:
        return endtime

    # df1["END_TIME"] = df.apply(lambda x:x["打款商家金额"] if x["确认收货时间"].find("null")>=0 else x["确认收货时间"], axis=1)


def get_status(status):
    if ((status.find("交易成功") >= 0) | (status.find("交易完成") >= 0) | (status.find("完成") >= 0) | (
            status.find("收款确认(服务完成)") >= 0) | (status.find("已签收") >= 0) | (status.find("已分账") >= 0)):
        return "TRADE_FINISHED"
    elif ((status.find("交易关闭") >= 0) | (status.find("删除") >= 0) | (status.find("取消") >= 0) | (status.find("在线支付超时") >= 0) | (status.find("订单取消") >= 0) | (status.find("已关闭") >= 0)):
        return "TRADE_CLOSED"
    elif status.find("等待买家付款") >= 0:
        return "WAIT_BUYER_PAY"
    elif ((status.find("等待卖家发货") >= 0) | (status.find("等待打印") >= 0) | (status.find("等待发货") >= 0)):
        return  "WAIT_SELLER_SEND_GOODS"
    elif status.find("卖家部分发货") >= 0:
        return  "SELLER_CONSIGNED_PART"
    elif ((status.find("等待买家确认收货") >= 0) | (status.find("等待买家收货") >= 0) | (status.find("等待确认收货") >= 0) | (status.find("已发货") >= 0)):
        return  "WAIT_BUYER_CONFIRM_GOODS"
    elif status.find("交易进行中") >= 0:
        return "TRADE_ONGOING"
    elif ((status.find("交易异常") >= 0) | (status.find("暂停") >= 0) | (status.find("拒绝签收") >= 0)):
        return "TRADE_ABNORMAL"
    elif ((status.find("交易退款") >= 0)|(status.find("订单退款成功") >= 0)):
        return "TRADE_REFUND"
    elif ((status.find("未知") >= 0) | (status.find("调度中") >= 0)):
        return "UNKNOWN"
    else:
        return "UNKNOWN"


def get_refund_status(refund_status):
    if refund_status.find("无售后或售后关闭"):
        return 0
    elif refund_status.find("售后处理中"):
        return 1
    elif refund_status.find("退款成功"):
        return 2
    elif refund_status.find("部分退款成功"):
        return 3


def get_payway(pay):
    if ((pay.find("支付宝")>=0)|(pay.find("TAOBAO")>=0)|(pay.find("TMALL")>=0)):
        return 1
    elif ((pay.find("微信")>=0)|(pay.find("微信支付")>=0)):
        return 2
    elif ((pay.find("拼多多")>=0)|(pay.find("PDD")>=0)):
        return 3
    elif pay.find("货到付款")>=0:
        return 4
    elif pay.find("银行卡")>=0:
        return 5
    elif pay.find("余额")>=0:
        return 6
    elif pay.find("放心花")>=0:
        return 7
    elif pay.find("新卡支付")>=0:
        return 8
    elif pay.find("QQ")>=0:
        return 9
    elif pay.find("连连支付")>=0:
        return 10
    elif pay.find("邮局汇款")>=0:
        return 11
    elif pay.find("自提")>=0:
        return 12
    elif ((pay.find("在线支付")>=0)|(pay.find("DY")>=0)|(pay.find("JD")>=0)):
        return 13
    elif pay.find("公司转账")>=0:
        return 14
    elif pay.find("银行卡转账")>=0:
        return 15
    elif pay.find("财付通")>=0:
        return 16
    elif pay.find("有赞零钱")>=0:
        return 16
    elif pay.find("网易白条")>=0:
        return 163
    elif pay.find("网易宝SDK")>=0:
        return 164
    elif pay.find("唯品会支付")>=0:
        return 165
    else:
        return 99


def get_all_files(rootdir, filekey):
    filelist = list_all_files(rootdir, filekey)
    # print(filelist)
    mySeries = pd.Series(filelist)
    df = pd.DataFrame(mySeries)
    df.columns = ["filename"]
    # df["filename"] = df["filename"].apply(lambda x: x.strip().replace(" ", ""))

    df = df[~df["filename"].str.contains("汇总")]
    df = df[~df["filename"].str.contains("刷单")]

    return df


def read_all_excel(rootdir, filekey):
    df_files = get_all_files(rootdir, filekey)
    for index, file in df_files.iterrows():
        if 'df1' in locals().keys():  # 如果变量已经存在
            dd,ff = read_excel(file["filename"])
            dd["filename"] = file["filename"]
            ff["filename"] = file["filename"]
            print("进度表：{}/{}  文件{}，行数{}".format(index + 1, df_files.shape[0], file["filename"], dd.shape[0], ff.shape[0]))

            df1 = df1.append(dd)
            df2 = df2.append(ff)


        else:
            df1,df2 = read_excel(file["filename"])
            df1["filename"] = file["filename"]
            df2["filename"] = file["filename"]
            # print(file["filename"],  df.shape[0])
            print("进度表：{}/{}  文件{}，行数{}".format(index + 1, df_files.shape[0], file["filename"], df1.shape[0], df2.shape[0]))

    return df1,df2


def combine_excel():
    print('订单数据校对逻辑:')
    print('1.财务订单数据需要放在财务数据文件夹下，例如/校对数据/财务数据/...')
    print('2.导出订单数据需要放在导出数据文件夹下，例如/校对数据/导出数据/...')
    print("请输入财务订单和导出订单所在的文件夹：")
    # filedir=""
    # filedir = input()
    # myTuple = shell.SHBrowseForFolder(0, None, "", 64)
    global default_dir

    type = win32api.MessageBox(0, "请确认你要处理的文件,单个文件选是，多个文件选否！", "提醒", win32con.MB_YESNO)
    if type == 6:
        file = open_file()
        print("需要处理的文件{}".format(file))
        table2 = read_excel(file)
        dir = os.sep.join(file.split(os.sep)[:-1])
        print("需要处理的文件所在目录{}".format(dir))

        default_dir = dir
        outfile = "".join(file.split(".")[:-1])

    else:
        try:
            # path = shell.SHGetPathFromIDList(myTuple[0])
            filedir = input()
        except:
            print("你没有输入任何目录 :(")
            sys.exit()
            return

        # filedir=path.decode('ansi')
        print("你选择的路径是：", filedir)

        # global default_dir
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
        table2 = read_all_excel(filedir, filekey)
        del table2["filename"]
    print(f"表2处理前总行数:{len(table2)}")

    table2["TID"] = table2["TID"].astype(str)
    table2["TID"] = table2["TID"].str.replace("\s+", "")
    table2.replace("nan", np.nan, inplace=True)
    table2.dropna(subset=["TID"], axis=0, inplace=True)
    # table2.drop_duplicates(inplace=True)
    table2 = table2.sort_values(by=["TID"])

    print(f"表2处理后总行数:{len(table2)}")

    plat = os.sep.join(default_dir.split(os.sep)[-1:])
    index = 0
    # df.to_excel(r"work/合并表格_test.xlsx")
    print("账单总行数：")
    print(table2.shape[0])
    for i in range(0, int(table2.shape[0] / 200000) + 1):
        print("存储分页：{}  from:{} to:{}".format(i, i * 200000, (i + 1) * 200000))
        # df.iloc[i*500000:(i+1)*500000].to_csv(default_dir + "\{}_合并表格_{}.csv".format(index,i))
        if type == 6:
            writer = pd.ExcelWriter(outfile + "-处理后的订单{}.xlsx".format(i))
        else:
            writer = pd.ExcelWriter(default_dir + "\{}-处理后的订单{}.xlsx".format(plat, i))
        table3 = table2.iloc[i * 200000:(i + 1) * 200000]
        print(f"表2分页{i}总行数:{len(table3)}")
        table3.to_excel(writer, sheet_name="订单明细", index=0)
        writer.save()
        writer.close()

    # writer = pd.ExcelWriter(default_dir + "\处理后的订单.xlsx")
    # table1.to_excel(writer,sheet_name='订单基础',index=0)
    # table2.to_excel(writer,sheet_name='订单明细',index=0)
    # writer.save()
    # writer.close()

    # index = 0
    # print("第{}个表格,记录数:{}".format(index, table.shape[0]))
    # print(table.head(10).to_markdown())
    # # df.to_excel(r"work/合并表格_test.xlsx")
    # print("账单总行数：")
    # print(table.shape[0])
    # for i in range(0, int(table.shape[0] / 200000) + 1):
    #     print("存储分页：{}  from:{} to:{}".format(i, i * 200000, (i + 1) * 200000))
    #     # df.iloc[i*500000:(i+1)*500000].to_csv(default_dir + "\{}_合并表格_{}.csv".format(index,i))
    #     table.iloc[i * 200000:(i + 1) * 200000].to_excel(default_dir + "\{}处理后的订单{}.xlsx".format(index, i), index=False)

    # return table


def get_shopcode(PLATFORM,SHOPNAME,type):
    df = pd.read_excel("data/shopcode.xlsx")
    for index, row in df.iterrows():
        if PLATFORM.find(row["店铺平台"]) >= 0:
            # print(row["店铺平台"])
            # print("111")
            if SHOPNAME.find(row["店铺名称"]) >= 0:
                if type == 1:
                    shopname = row["店铺名称"]
                    return shopname
                else:
                    shopcode = row["店铺代码"]
                    # print("平台名:", PLATFORM, "店铺名称为：", SHOPNAME, "店铺代码为：", shopcode)
                    return shopcode
            else:
                # print("未找到店铺名称")
                pass
        else:
            # print("未找到店铺平台")
            pass
    return "N/A"
    # df["店铺代码"] = df["店铺代码"].loc[(df["店铺平台"].str.contains(PLATFORM) & df["店铺名称"].str.contains(SHOPNAME))]
    # df["店铺代码"] = ~df["店铺代码"].str.contains("nan")
    # print(df["店铺代码"])
    # return shop
    # shopcode = df["店铺代码"]
    # return shop
    # else:
    #     print("未找到店铺平台或者店铺名称")
    #     return "N/A"


if __name__ == "__main__":
    print("开始:", time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()))
    # print("abc".find("x"))

    combine_excel()

    # groupby_amt()
    # math_file()
    # get_shopcode("JD","dentylactive旗舰店")

    print("ok")