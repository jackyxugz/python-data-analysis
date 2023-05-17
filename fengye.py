# __coding=utf8__
# /** 作者：zengyanghui **/

# __coding=utf8__
# /** 作者：zengyanghui **/
# __coding=utf8__
# /** 作者：zengyanghui **/

import sys
import os
import pandas as pd

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


def list_all_files(rootdir, filekey_list):
    if len(filekey_list) > 0:
        filekey_list = filekey_list.replace(",", " ")
        filekey = filekey_list.split(" ")
    else:
        filekey = ''

    _files = []
    list_a = os.listdir(rootdir)  # 列出文件夹下所有的目录与文件
    for i in range(0, len(list_a)):
        path = os.path.join(rootdir, list_a[i])
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

    # if filename.find("天猫") > 0:
    #     try:
    #         dfx = pd.read_excel(filename, sheet_name="订单原表", dtype=str)
    #         print(f"订单表总行数:{len(dfx)}")
    #         dfy = pd.read_excel(filename, sheet_name="属性原表", dtype=str)
    #         print(f"属性表总行数:{len(dfy)}")
    #         df = pd.merge(dfx,dfy[["订单编号","标题","价格","购买数量","商品属性","备注","商家编码"]],how="left",on="订单编号")
    #         print(f"合并表总行数:{len(df)}")
    #     except Exception as e:
    #         print(f"读取{filename}文件报错！")
    #         error(e,filename)
    #         # logging.basicConfig(filename=default_dir + "\错误日志.log",format='[%(asctime)s-%(filename)s-%(levelname)s:%(message)s]', level = logging.DEBUG,filemode='a',datefmt='%Y-%m-%d%I:%M:%S %p')
    #         dict1 = {"PAYMENT": "", "MODIFIED": "", "PAY_TIME": "", "POST_FEE": "", "DISCOUNT_FEE": "", "STATUS": "",
    #                  "CREATED": "", "TID": "", "RECEIVED_PAYMENT": "", "BUYER_NICK": "", "SELLER_NICK": "",
    #                  "SELLER_MEMO": "", "CONSIGN_TIME": "", "END_TIME": "", "PLATFORM": "", "SHOPCODE": "",
    #                  "BUYER_MESSAGE": "", "PAY_NO": "", "PAY_WAY": "", "REFUND_STATUS": "", "RECEIVER_ADDRESS": "",
    #                  "RECEIVER_MOBILE": "", "RECEIVER_NAME": "", "RECEIVER_STATE": "", "RECEIVER_ZIP": "",
    #                  "RECEIVER_DISTRICT": "", "RECEIVER_CITY": "", "RECEIVER_TOWN": ""}
    #         df1 = pd.DataFrame(dict1,index=[0])
    #         dict2 = {"TITLE": "", "SKU": "", "CREATED": "", "STATUS": "", "PRICE": "", "PAYMENT": "", "PIC_PATH": "",
    #                  "NUM": "", "END_TIME": "", "PAY_TIME": "", "LOGISTICS_COMPANY": "", "INVOICE_NO": "", "TID": "",
    #                  "SKU_PROPERTIES_NAME": "", "OID": "", "BUYER_NICK": "", "SELLER_NICK": "", "REFUND_STATUS": "",
    #                  "PLATFORM": "", "SHOPCODE": "", "CONSIGN_TIME": ""}
    #         df2 = pd.DataFrame(dict2, index=[0])
    #         return df1, df2
    #
    #     for column_name in df.columns:
    #         df.rename(columns={column_name: column_name.replace(" ", "").replace("\n", "").strip()}, inplace=True)
    #     df["物流单号"] = df["物流单号"].str.replace("No:","")
    #     df["打款商家金额"] = df["打款商家金额"].str.replace("元","")
    #     print(df.head(5).to_markdown())
    #     print(df["店铺名称"][1])
    #     plat = "TMALL"
    #     df1 = pd.DataFrame()
    #     df1["PAYMENT"] = df["买家实际支付金额"]
    #     df1["MODIFIED"] = ""
    #     df1["PAY_TIME"] = df["订单付款时间"]
    #     df1["POST_FEE"] = df["买家应付邮费"]
    #     df1["DISCOUNT_FEE"] = ""
    #     df1["STATUS"] = df["订单状态"].apply(lambda x:get_status(x))
    #     df1["CREATED"] = df["订单创建时间"]
    #     df1["TID"] = df["订单编号"]
    #     df1["RECEIVED_PAYMENT"] = df["打款商家金额"]
    #     df1["BUYER_NICK"] = df["买家会员名"]
    #     df1["SELLER_NICK"] = df["店铺名称"]
    #     df1["SELLER_MEMO"] = df["订单备注"]
    #     df1["CONSIGN_TIME"] = ""
    #     df1["END_TIME"] = df["确认收货时间"]
    #     df1["PLATFORM"] = plat
    #     df1["SHOPCODE"] = get_shopcode(plat,df["店铺名称"][1])
    #     df1["BUYER_MESSAGE"] = df["买家留言"]
    #     df1["PAY_NO"] = ""
    #     df1["PAY_WAY"] = 1
    #     df1["REFUND_STATUS"] = ""
    #     df1["RECEIVER_ADDRESS"] = df["收货地址"]
    #     df1["RECEIVER_MOBILE"] = df.apply(lambda x:x["联系电话"] if pd.isnull(x["联系手机"]) else x["联系手机"],axis=1)
    #     df1["RECEIVER_NAME"] = df["收货人姓名"]
    #     df1["RECEIVER_STATE"] = df["收货地址"].apply(lambda x:"".join(x.split(" ")[:1]))
    #     df1["RECEIVER_ZIP"] = ""
    #     df1["RECEIVER_DISTRICT"] = df["收货地址"].apply(lambda x:"".join(x.split(" ")[2:3]))
    #     df1["RECEIVER_CITY"] = df["收货地址"].apply(lambda x:"".join(x.split(" ")[1:2]))
    #     df1["RECEIVER_TOWN"] = df["收货地址"].apply(lambda x:"".join(x.split(" ")[3:]))
    #
    #     df1["PAYMENT"] = df1["PAYMENT"].astype(float)
    #     df1["RECEIVED_PAYMENT"] = df1["RECEIVED_PAYMENT"].astype(float)
    #     df1["POST_FEE"] = df1["POST_FEE"].astype(float)
    #     print("天猫-订单基础")
    #     print(df1.head(5).to_markdown())
    #
    #     df2 = pd.DataFrame()
    #     df2["TITLE"] = df["标题"]
    #     df2["SKU"] = df["商家编码"]
    #     df2["CREATED"] = df["订单创建时间"]
    #     df2["STATUS"] = df["订单状态"].apply(lambda x:get_status(x))
    #     df2["PRICE"] = df["价格"]
    #     df2["PAYMENT"] = df["买家实际支付金额"]
    #     df2["PIC_PATH"] = ""
    #     df2["NUM"] = df["购买数量"]
    #     df2["END_TIME"] = ""
    #     df2["PAY_TIME"] = df["订单付款时间"]
    #     df2["LOGISTICS_COMPANY"] = df["物流公司"]
    #     df2["INVOICE_NO"] = df["物流单号"]
    #     df2["TID"] = df["订单编号"]
    #     df2["SKU_PROPERTIES_NAME"] = df["商品属性"]
    #     df2["OID"] = ""
    #     df2["BUYER_NICK"] = df["买家会员名"]
    #     df2["SELLER_NICK"] = df["店铺名称"]
    #     df2["REFUND_STATUS"] = ""
    #     df2["PLATFORM"] = plat
    #     df2["SHOPCODE"] = df1["SHOPCODE"]
    #     df2["CONSIGN_TIME"] = ""
    #
    #     df2["PAYMENT"] = df2["PAYMENT"].astype(float)
    #     df2["PRICE"] = df2["PRICE"].astype(float)
    #     print("天猫-订单明细")
    #     print(df2.head(5).to_markdown())
    if filename.find("麦凯莱") > 0:
        print("发现麦凯莱")
        try:
            df = pd.read_excel(filename, dtype=str)
            print(f"订单表总行数:{len(df)}")
            # df1 = df[["订单编号", "下单时间", "订单来源", "订单状态", "销售门店", "服务门店", "支付方式", "支付时间", "支付单号", "订单总金额",
            #           "实收金额", "物流费用", "买家备注", "发货时间", "收货人姓名", "收货人手机号", "收货人所在省份", "收货人所在城市",
            #           "收货人所在区/县", "收货人所在乡镇/街道", "收货人详细地址", "邮编", "客户昵称", "订单完成时间", "修改时间", "售后状态"]]
            # df1 = df1.drop_duplicates(subset=["订单编号"], keep="last")
            #
            # df2 = df[["订单编号", "下单时间", "订单来源", "订单状态", "销售门店", "服务门店", "支付时间", "订单总金额", "包裹id",
            #           "发货时间", "配送公司", "配送单号", "商品名称", "商家编码", "商品数量", "商品规格", "商品单价",
            #           "客户昵称", "订单完成时间", "售后状态"]]


        except Exception as e:
            print(f"读取{filename}文件报错！")
            error(e, filename)
            # logging.basicConfig(filename=default_dir + "\错误日志.log",format='[%(asctime)s-%(filename)s-%(levelname)s:%(message)s]', level = logging.DEBUG,filemode='a',datefmt='%Y-%m-%d%I:%M:%S %p')
            dict1 = {"PAYMENT": "", "MODIFIED": "", "PAY_TIME": "", "POST_FEE": "", "DISCOUNT_FEE": "", "STATUS": "",
                     "CREATED": "", "TID": "", "RECEIVED_PAYMENT": "", "BUYER_NICK": "", "SELLER_NICK": "",
                     "SELLER_MEMO": "", "CONSIGN_TIME": "", "END_TIME": "", "PLATFORM": "", "SHOPCODE": "",
                     "BUYER_MESSAGE": "", "PAY_NO": "", "PAY_WAY": "", "REFUND_STATUS": "", "RECEIVER_ADDRESS": "",
                     "RECEIVER_MOBILE": "", "RECEIVER_NAME": "", "RECEIVER_STATE": "", "RECEIVER_ZIP": "",
                     "RECEIVER_DISTRICT": "", "RECEIVER_CITY": "", "RECEIVER_TOWN": ""}
            df1 = pd.DataFrame(dict1, index=[0])
            dict2 = {"TITLE": "", "SKU": "", "CREATED": "", "STATUS": "", "PRICE": "", "PAYMENT": "", "PIC_PATH": "",
                     "NUM": "", "END_TIME": "", "PAY_TIME": "", "LOGISTICS_COMPANY": "", "INVOICE_NO": "", "TID": "",
                     "SKU_PROPERTIES_NAME": "", "OID": "", "BUYER_NICK": "", "SELLER_NICK": "", "REFUND_STATUS": "",
                     "PLATFORM": "", "SHOPCODE": "", "CONSIGN_TIME": ""}
            df2 = pd.DataFrame(dict2, index=[0])
            return df1, df2

        for column_name in df.columns:
            df.rename(columns={column_name: column_name.replace(" ", "").replace("\n", "").strip()}, inplace=True)

         # df1["TID"] = df["订单编号"]
         # df1["CREATED"] = df["下单时间"]



        # df1.rename(columns={"订单编号": "TID", "下单时间": "CREATED", "订单来源": "PLATFORM", "订单状态": "STATUS",
        #                     "销售门店": "SELLER_NICK", "服务门店": "SELLER_NICK", "支付方式": "PAY_WAY", "支付时间": "PAY_TIME",
        #                     "支付单号": "PAY_NO", "订单总金额": "PAYMENT", "实收金额": "RECEIVED_PAYMENT", "物流费用": "POST_FEE",
        #                     "买家备注": "BUYER_MESSAGE", "发货时间": "CONSIGN_TIME", "收货人姓名": "RECEIVER_NAME",
        #                     "收货人手机号": "RECEIVER_MOBILE", "收货人所在省份": "RECEIVER_STATE",
        #                     "收货人所在城市": "RECEIVER_CITY", "收货人所在区/县": "RECEIVER_DISTRICT",
        #                     "收货人所在乡镇/街道": "RECEIVER_TOWN", "收货人详细地址": "RECEIVER_ADDRESS",
        #                     "邮编": "RECEIVER_ZIP", "客户昵称": "BUYER_NICK", "订单完成时间": "END_TIME",
        #                     "售后状态": "REFUND_STATUS"})
        #
        # df2.rename(columns={"订单编号": "TID", "下单时间": "CREATED", "订单来源": "PLATFORM", "订单状态": "STATUS",
        #                     "销售门店": "SELLER_NICK", "服务门店": "SELLER_NICK",
        #                     "支付时间": "PAY_TIME", "订单总金额": "PAYMENT", "包裹id": "OID", "发货时间": "CONSIGN_TIME",
        #                     "配送公司": "LOGISTICS_COMPANY", "配送单号": "INVOICE_NO", "商品名称": "TITLE", "商家编码": "SKU",
        #                     "商品数量": "NUM", "商品规格": "SKU_PROPERTIES_NAME", "商品单价": "PRICE", "客户昵称": "BUYER_NICK",
        #                     "订单完成时间": "END_TIME", "售后状态": "REFUND_STATUS"})

        #
        # df1["TID"] = df["订单编号"]
        # df1["TID"] = df["订单编号"]
        # df1["TID"] = df["订单编号"]
        #

        # df1["MODIFIED"] = ""
        # df1["DISCOUNT_FEE"] = ""
        # df1["SELLER_MEMO"] = ""
        # df1["SHOPCODE"] = ""
        #
        # df2["PIC_PATH"] = ""
        # df2["SHOPCODE"] = ""

        #df["配送单号"] = df["配送单号"].str.replace("No:", "")
        #df["订单总金额"] = df["订单总金额"].str.replace("元", "")
        #print(df.head(5).to_markdown())
        # print(df["订单原表"][1])
        plat = "BD" #"百度店铺代码"
        df1 = pd.DataFrame()
        df1["PAYMENT"] = df["订单实付金额"]#df["实收金额"]
        df1["MODIFIED"] = ""#df["修改时间"]
        df1["PAY_TIME"] = ""#df["支付时间"]
        df1["POST_FEE"] = "" #物流费用
        df1["DISCOUNT_FEE"] = ""#df["折扣金额"]
        df1["STATUS"] = df["订单状态"].apply(lambda x: get_status(x))#df["订单状态"]
        df1["CREATED"] = df["下单时间"]#df["下单时间"]
        df1["TID"] = df["订单编号"]#df["订单编号"]
        df1["RECEIVED_PAYMENT"] = df["商品总价"]#df["实收金额"]
        df1["BUYER_NICK"] = ""#df["用户昵称"]
        df1["SELLER_NICK"] = df["商家店铺名称"]#df["店铺名称"]
        df1["SELLER_MEMO"] = ""#df["卖家备注"]
        df1["CONSIGN_TIME"] = ""#df["发货时间"]
        df1["END_TIME"] = ""#df["发货时间"]
        df1["PLATFORM"] = plat#df["电商平台"]
        df1["SHOPCODE"] = "mklzt" #get_shopcode(plat, df["店铺名称"][1])""
        df1["BUYER_MESSAGE"] = df["买家留言"]#df["买家留言"]
        df1["PAY_NO"] = ""#["支付单号"]
        df1["PAY_WAY"] = ""#["支付方式"]比如支付宝或微信
        df1["REFUND_STATUS"] = df["售后状态"].apply(lambda h: get_refund_status(h))
        df1["RECEIVER_ADDRESS"] = df["收货人地址"]#df["收货地址"]
        df1["RECEIVER_MOBILE"] = df["收货人电话"] #df.apply(lambda x: x["收货人手机号"] if pd.isnull(x["收货人手机号"]) else x["收货人手机号"], axis=1)
        df1["RECEIVER_NAME"] = df["收货人姓名"]
        df1["RECEIVER_STATE"] = df["收货人所在省份"] #df["收货地址"].apply(lambda x: "".join(x.split(" ")[:1]))
        df1["RECEIVER_ZIP"] = ""#df["邮编"]
        df1["RECEIVER_DISTRICT"] = df["区"] #df["收货地址"].apply(lambda x: "".join(x.split(" ")[2:3]))
        df1["RECEIVER_CITY"] = df["市"] #df["收货地址"].apply(lambda x: "".join(x.split(" ")[1:2]))
        df1["RECEIVER_TOWN"] = "" #df["收货地址"].apply(lambda x: "".join(x.split(" ")[3:]))

        df1["PAYMENT"] = df1["PAYMENT"].astype(float)
        df1["RECEIVED_PAYMENT"] = df1["RECEIVED_PAYMENT"].astype(float)
        # df1["POST_FEE"] = df1["POST_FEE"].astype(float)
        print("百度麦凯莱-订单基础")
        #print(df1.head(5).to_markdown())

        df2 = pd.DataFrame()
        df2["TITLE"] = df["选购产品"]#df["商品名称"]
        df2["SKU"] = df["sku编码"]#df["商家编码"]
        df2["CREATED"] = df["下单时间"]#df["下单时间"]
        df2["STATUS"] = df["订单状态"].apply(lambda x: get_status(x))#df["订单状态"]
        df2["PRICE"] = df["总价"]#df["商品单价"]
        df2["PAYMENT"] = df["总价"]#df["实收金额"]
        df2["PIC_PATH"] = ""#df["图片地址"]
        df2["NUM"] = df["数量"]#df["商品数量"]
        df2["END_TIME"] = ""#df["结算时间"]
        df2["PAY_TIME"] = ""#df["支付时间"]
        df2["LOGISTICS_COMPANY"] = df["快递公司"]#df["物流公司"]
        df2["INVOICE_NO"] = df["单号"]#df["物流单号"]
        df2["TID"] = df["订单编号"]#df["订单编号"]
        df2["SKU_PROPERTIES_NAME"] = df["数量明细"] #df["商品规格"]
        df2["OID"] = df["第三方订单ID"] #子订单号
        df2["BUYER_NICK"] = ""#df["客户昵称"]
        df2["SELLER_NICK"] = df["麦凯莱科技"]#df["销售门店"]
        df2["REFUND_STATUS"] = df["售后状态"].apply(lambda h: get_refund_status(h))#df["售后状态"]
        df2["PLATFORM"] = plat#df["电商平台"]
        df2["SHOPCODE"] = "mklzt"#df["电商平台代码"]比如百度为BD
        df2["CONSIGN_TIME"] = ""#df["发货时间"]

        df2["PAYMENT"] = df2["PAYMENT"].astype(float)
        df2["PRICE"] = df2["PRICE"].astype(float)
        print("百度麦凯莱-订单明细")
        # print(df2.head(5).to_markdown())
    elif filename.find("淘宝") > 0:
        print("没有发现！")

    return df1, df2


def error(self, filename):
    logging.basicConfig(filename=default_dir + "\错误日志.log",
                        format=f'[%(asctime)s-%(filename)s-%(levelname)s:%(message)s:{filename}]', level=logging.DEBUG,
                        filemode='a', datefmt='%Y-%m-%d%I:%M:%S %p')

    logging.error("这是一条error信息的打印")
    # logging.info("这是一条info信息的打印")
    # logging.warning("这是一条warn信息的打印")
    # logging.debug("这是一条debug信息的打印")


def get_status(status):
    a = str(status)
    if a.find("已发货") >= 0:
        return "WAIT_BUYER_CONFIRM_GOODS"
    elif a.find("交易关闭") >= 0:
        return "WAIT_SELLER_SEND_GOODS"
    elif a.find("待发货") >= 0:
        return "TRADE_CLOSED"
    elif a.find("交易成功") >= 0:
        return "UTRADE_FINISHED"
    elif a.find("未知") >= 0:
        return "UNKNOWN"



    # if status.find("等待买家付款") >= 0:
    #     return "WAIT_BUYER_PAY"
    # elif status.find("等待卖家发货") >= 0:
    #     return "WAIT_SELLER_SEND_GOODS"
    # elif status.find("卖家部分发货") >= 0:
    #     return "SELLER_CONSIGNED_PART"
    # elif ((status.find("等待买家确认收货") >= 0) | (status.find("等待买家收货") >= 0)):
    #     return "WAIT_BUYER_CONFIRM_GOODS"
    # elif status.find("交易进行中") >= 0:
    #     return "TRADE_ONGOING"
    # elif ((status.find("交易成功") >= 0) | (status.find("交易完成") >= 0)):
    #     return "TRADE_FINISHED"
    # elif status.find("交易异常") >= 0:
    #     return "TRADE_ABNORMAL"
    # elif status.find("交易退款") >= 0:
    #     return "TRADE_REFUND"
    # elif status.find("交易关闭") >= 0:
    #     return "TRADE_CLOSED"
    # elif status.find("未知") >= 0:
    #     return "UNKNOWN"


def get_refund_status(refund_status):
    refund_status1 = str(refund_status)
    if refund_status1.find("已退款") >= 0:
        return 2
    elif refund_status1.find("售后处理中") >= 0:
        return 1
    elif refund_status1.find("部分退款成功") >= 0:
        return 3

    # elif refund_status.find("售后处理中"):
    #     return 1
    # elif refund_status.find("退款成功"):
    #     return 2
    # elif refund_status.find("部分退款成功"):
    #     return 3


def get_payway(pay):
    pay1 = str(pay)
    if pay1.find("支付宝") >= 0:
        return 1
    elif pay1.find("微信") >= 0:
        return 2
    elif pay1.find("拼多多") >= 0:
        return 3
    elif pay1.find("货到付款") >= 0:
        return 4
    elif pay1.find("银行卡") >= 0:
        return 5
    elif pay1.find("余额") >= 0:
        return 6
    elif pay1.find("放心花") >= 0:
        return 7
    elif pay1.find("新卡支付") >= 0:
        return 8
    elif pay1.find("QQ") >= 0:
        return 9
    elif pay1.find("连连支付") >= 0:
        return 10
    elif pay1.find("邮局汇款") >= 0:
        return 11
    elif pay1.find("自提") >= 0:
        return 12
    elif pay1.find("线上支付") >= 0:
        return 13
    elif pay1.find("公司转账") >= 0:
        return 14
    elif pay1.find("银行卡转账") >= 0:
        return 15
    elif pay1.find("财付通") >= 0:
        return 16
    elif pay1.find("有赞零钱") >= 0:
        return 16
    elif pay1.find("网易白条") >= 0:
        return 163
    elif pay1.find("网易宝SDK") >= 0:
        return 164
    elif pay1.find("唯品会支付") >= 0:
        return 165
    elif pay1.find("无需支付") >= 0: # 未知
        return 99



    # if pay.find("支付宝") >= 0:
    #     return 1
    # elif pay.find("微信") >= 0:
    #     return 2
    # elif pay.find("拼多多") >= 0:
    #     return 3
    # elif pay.find("货到付款") >= 0:
    #     return 4
    # elif pay.find("银行卡") >= 0:
    #     return 5
    # elif pay.find("余额") >= 0:
    #     return 6
    # elif pay.find("放心花") >= 0:
    #     return 7
    # elif pay.find("新卡支付") >= 0:
    #     return 8
    # elif pay.find("QQ") >= 0:
    #     return 9
    # elif pay.find("连连支付") >= 0:
    #     return 10
    # elif pay.find("邮局汇款") >= 0:
    #     return 11
    # elif pay.find("自提") >= 0:
    #     return 12
    # elif pay.find("在线支付") >= 0:
    #     return 13
    # elif pay.find("公司转账") >= 0:
    #     return 14
    # elif pay.find("银行卡转账") >= 0:
    #     return 15
    # elif pay.find("财付通") >= 0:
    #     return 16
    # elif pay.find("有赞零钱") >= 0:
    #     return 16
    # elif pay.find("网易白条") >= 0:
    #     return 163
    # elif pay.find("网易宝SDK") >= 0:
    #     return 164
    # elif pay.find("唯品会支付") >= 0:
    #     return 165
    # elif pay.find("未知") >= 0:
    #     return 99


def get_all_files(rootdir, filekey):
    filelist = list_all_files(rootdir, filekey)
    # print(filelist)
    mySeries = pd.Series(filelist)
    df = pd.DataFrame(mySeries)
    df.columns = ["filename"]
    # df["filename"] = df["filename"].apply(lambda x: x.strip().replace(" ", ""))

    df = df[~df["filename"].str.contains("汇总")]

    return df


def read_all_excel(rootdir, filekey):
    df_files = get_all_files(rootdir, filekey)
    for index, file in df_files.iterrows():
        if 'df1' in locals().keys():  # 如果变量已经存在
            dd, ff = read_excel(file["filename"])
            dd["filename"] = file["filename"]
            ff["filename"] = file["filename"]
            print(
                "进度表：{}/{}  文件{}，行数{}".format(index + 1, df_files.shape[0], file["filename"], dd.shape[0], ff.shape[0]))

            df1 = df1.append(dd)
            df2 = df2.append(ff)

        else:
            df1, df2 = read_excel(file["filename"])
            df1["filename"] = file["filename"]
            df2["filename"] = file["filename"]
            # print(file["filename"],  df.shape[0])
            print("进度表：{}/{}  文件{}，行数{}".format(index + 1, df_files.shape[0], file["filename"], df1.shape[0],
                                                df2.shape[0]))

    return df1, df2


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
    table1, table2 = read_all_excel(filedir, filekey)
    # del table1["filename"]
    # del table2["filename"]
    print(f"表1处理前总行数:{len(table1)}")
    print(f"表2处理前总行数:{len(table2)}")

    # table.to_excel(default_dir + "/处理后的订单.xlsx", index=False)
    table1["TID"] = table1["TID"].astype(str)
    table1.replace("nan", np.nan, inplace=True)
    table1.dropna(subset=["TID"], axis=0, inplace=True)
    table1.drop_duplicates(inplace=True)
    table1 = table1.sort_values(by=["TID", "CREATED"])

    table2["TID"] = table2["TID"].astype(str)
    table2.replace("nan", np.nan, inplace=True)
    table2.dropna(subset=["TID"], axis=0, inplace=True)
    table2.drop_duplicates(inplace=True)
    table2 = table2.sort_values(by=["TID", "CREATED"])

    print(f"表1处理后总行数:{len(table1)}")
    print(f"表2处理后总行数:{len(table2)}")

    index = 0
    # print("第{}个表格,记录数:{}".format(index, table1.shape[0]))
    # print(table1.head(10).to_markdown())
    # df.to_excel(r"work/合并表格_test.xlsx")
    print("账单总行数：")
    print(table1.shape[0])
    print(table2.shape[0])
    for i in range(0, int(table1.shape[0] / 200000) + 1):
        print("存储分页：{}  from:{} to:{}".format(i, i * 200000, (i + 1) * 200000))
        # df.iloc[i*500000:(i+1)*500000].to_csv(default_dir + "\{}_合并表格_{}.csv".format(index,i))
        writer = pd.ExcelWriter(default_dir + "\{}处理后的订单{}.xlsx".format(index, i))
        table3 = table1.iloc[i * 200000:(i + 1) * 200000]
        print(f"表1分页{i}总行数:{len(table3)}")
        table3.to_excel(writer, sheet_name="订单基础", index=0)
        table4 = table2[table2["TID"].isin(table3["TID"])]
        print(f"表2分页{i}总行数:{len(table4)}")
        table4.to_excel(writer, sheet_name="订单明细", index=0)
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


def get_shopcode(PLATFORM, SHOPNAME):
    df = pd.read_excel("data/shopcode.xlsx")
    for index, row in df.iterrows():
        if PLATFORM.find(row["店铺平台"]) >= 0:
            # print(row["店铺平台"])
            # print("111")
            if SHOPNAME.find(row["店铺名称"]) >= 0:
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
