import pandas as pd

import pandas as pd
import numpy as np
from collections import Counter
import os
# import Tkinter
import win32api
import win32ui
import win32con
import tabulate
import math

def get_fn():
    choice = win32api.MessageBox(0, "请选择要处理的财务文件，否则退出程序！", "提醒", win32con.MB_YESNO)
    if choice == 6:
        # print("请选择财务文件！")
        file_path = open_file()
        df = pd.read_excel(file_path,dtype=str)
        df = pd.DataFrame(df)
        print(df.head(5).to_markdown())
        for column_name in df.columns:
            df.rename(columns={column_name: column_name.replace("\n", "").replace(" ", "").strip()}, inplace=True)
        # df2 = df[["入库单号","系统单号","接单日期", "商品条码", "商品名称", "入库数量","订货价(不含税)", "入库总额（不含税）","订货金额（不含税）"]].copy()
        df2 = df[["ERP单号","订单号","入库时间","货物名称","实际入库数量","成本含税","送货金额含税"]].copy()
        print(df2.head(5).to_markdown())
        df2.rename(columns={"ERP单号": "so","订单号": "customer","入库时间":"date","货物名称": "product",
                            "实际入库数量": "qty", "成本含税": "price", "送货金额含税": "amt"},
                   inplace=True)
        print(df2.head(5).to_markdown())
        df2["so"] = df2["so"].astype(str)
        df2["product"] = df2["product"].astype(str)
        df2.qty.fillna("0", inplace=True)
        df2["qty"] = df2["qty"].astype(int)
        df2.price.fillna("0", inplace=True)
        df2["price"] = df2["price"].apply(lambda x: str(x).replace(" ", "0").replace("nan", "0").replace("NAN", "0").replace("-", "0"))
        # # df2["price"] = df2["price"].astype(float)
        df2["amt"] = df2["amt"].astype(float)
        df2["amt"] = df2["amt"].map(lambda x: "{:.2f}".format(x))
        # df2["total_amt"] = df2["total_amt"].astype(float)
        # df2["total_amt"] = df2["total_amt"].map(lambda x: "{:.2f}".format(x))
        # df2["date"] = df2["date"].astype(str).apply(lambda x: x.replace(".", "-").replace("-2-29", "-2-28"))
        # df2["date"] = df2["date"].astype("datetime64[ns]")
        # df2["iyear"]=df2["date"].apply(lambda x: x.year)
        # df2["imonth"] = df2["date"].apply(lambda x: x.month)
        # df2["year_month"] = df2.apply(lambda x: "{}-{}".format(x["date"].year, x["date"].month), axis=1)

        df2 = df2.apply(lambda x: x.astype(str).str.replace(" ", "").str.replace(",", "").str.replace("\n", "").str.strip())
        df2 = df2.apply(lambda x: x.astype(str).str.replace("(", "）").str.replace(")", "）"))

        # file_type = win32api.MessageBox(0, "选择是否要输出处理后的明细", "提醒", win32con.MB_YESNO)
        # if file_type == 6:
        #     item = win32api.MessageBox(0, "选择是否输出方式，是：按月保存文件，否：全部保存一个文件！","提醒",  win32con.MB_YESNO)
        #     if item == 6:
        #         df3 = df2[["so","erp_so","customer","sku","product","year_month","order_qty","real_qty", "amt"]]
        #         df3 = df3.sort_values(by=["so","erp_so","psku","real_qty"])
        #         print(df3.head(5).to_markdown())
        #         # df3.to_excel("data/fn-item2019.xlsx")
        #         for x in range(1,13):
        #             s="{}".format(x)
        #             df3[df3.year_month.str.contains(s)].to_excel("data/fn-item-{}.xlsx".format(x),index=False)
        #     else:

        df3 = df2[["so","customer","date","product","qty","price", "amt"]]
        # df3 = df3.drop_duplicates(["erp_so"], keep='first')
        df3 = df3.sort_values(by=["so","customer","qty"])
        print(df3.head(5).to_markdown())
        print(df3.head(5).to_markdown())
        df3.to_excel(r"C:\Users\mega\Downloads\对账\2019\处理后\fn-19永辉.xlsx",index=False)
        # else:
        #     print("未输出明细文件！")

        # sum_file = win32api.MessageBox(0, "选择是否要输出按月汇总文件", "提醒", win32con.MB_YESNO)
        # if sum_file == 6:
        #     group_date = df2.groupby(["so","erp_so","iyear","imonth"]).agg({"order_qty":"sum","real_qty":"sum","amt":"sum"})
        #     group_date = pd.DataFrame(group_date).reset_index()
        #     print(group_date.head(5).to_markdown())
        #     group_date.to_excel("data/fn-sum.xlsx",index=False)
        # else:
        #     print("未输出按月汇总文件！")
    else:
        print("退出财务文件处理！")
        pass

def get_it():
    choice = win32api.MessageBox(0, "请选择要处理的IT文件，否则退出程序！", "提醒", win32con.MB_YESNO)
    if choice == 6:
        file_path = open_file()
        df = pd.read_excel(file_path, dtype=str)
        # df = pd.read_pickle("data/it-item_table.pkl")
        # df2 = df[["销售明细行/订单关联/客户参考", "销售明细行/订单关联", "产品/条码", "产品", "完成数量", "取值价格", "小计","税额总计"]].copy()
        # df2.rename(columns={"销售明细行/订单关联/客户参考": "customer", "销售明细行/订单关联": "so", "产品/条码": "sku", "产品": "product",
        #                     "完成数量": "qty", "取值价格": "price", "小计": "amt", "税额总计": "tax"}, inplace=True)
        df2 = df[["销售明细行/订单关联/客户参考", "销售明细行/订单关联"]].copy()
        df2.rename(columns={"销售明细行/订单关联/客户参考": "customer", "销售明细行/订单关联": "so"}, inplace=True)

        # print(df2.sort_values(by=["date"] ,ascending=False).head(20).to_markdown())
        df2["so"] = df2["so"].astype(str)
        # df2["erp_so"] = df2["erp_so"].astype(str)
        df2["customer"] = df2["customer"].astype(str)
        # df2["sku"] = df2["sku"].astype(str)
        # df2["product"] = df2["product"].astype(str)
        # df2.qty.fillna("0", inplace=True)
        # df2["qty"] = df2["qty"].astype(float)
        # df2.price.fillna("0", inplace=True)
        # df2["price"] = df2["price"].astype(float)
        # df2["amt"] = df2["amt"].astype(float)
        # df2["date"] = df2["date"].astype("datetime64[ns]")
        # df2["iyear"]=df2["date"].apply(lambda x: x.year)
        # df2["imonth"] = df2["date"].apply(lambda x: x.month)
        # df2["year_month"] =df2.apply(lambda x: "{}-{}".format(x["date"].year,x["date"].month) ,axis=1)
        df2 = df2.apply(lambda x: x.astype(str).str.replace(" ", "").str.replace(",", "").str.replace("\n", "").str.strip())
        df2 = df2.apply(lambda x: x.astype(str).str.replace("(", "）").str.replace(")", "）").str.replace("=", "").str.replace("'", "").str.replace('"', ''))

        # file_type = win32api.MessageBox(0, "选择是否要输出处理后的明细", "提醒", win32con.MB_YESNO)
        # if file_type == 6:
        #     item = win32api.MessageBox(0, "选择是否输出方式，是：按月保存文件，否：全部保存一个文件！","提醒",  win32con.MB_YESNO)
        #     if item == 6:
        #         df3 = df2[["so","erp_so","customer","sku","product","year_month","order_qty","real_qty","amt"]]
        #         df3 = df3.sort_values(by=["so","erp_so","sku","real_qty"])
        #         print(df3.head(5).to_markdown())
        #         for x in range(1,13):
        #             s="2019-{}".format(x)
        #             df3[df3.year_month.str.contains(s)].to_excel("data/it-item-{}.xlsx".format(x),index=False)
        #     else:

        # df2 = df2.loc[df2["type"].str.contains("销售出库|销售退货入库")]
        # df2 = df2.loc[df2["customer"].str.contains("京东7FRESH(北京四季优选)")]

        # df2 = df2.query('type=="销售出库" | type=="销售退货入库"')
        # df2 = df2.query('customer=="华润万家"')

        df3 = df2[["so","customer"]]
        df3 = df3.sort_values(by=["so","customer"])
        df3.dropna(subset=["so"],axis=0,inplace=True)
        df3["so"] = df3["so"].astype(str)
        df3["customer"] = df3["customer"].astype(str)
        df3 = df3[~df3["so"].str.contains("nan") & ~df3["customer"].str.contains("nan")]
        # df3 = df3.drop_duplicates(["erp_so"], keep='first')
        print(df3.head(5).to_markdown())
        print(len(df3))
        df3.to_excel(r"C:\Users\mega\Downloads\补充ERP单号\it-item.xlsx",index=False)
        # else:
        #     print("未输出明细文件！")

        # sum_file = win32api.MessageBox(0, "选择是否要输出按月汇总文件", "提醒", win32con.MB_YESNO)
        # if sum_file == 6:
        #     group_date = df2.groupby(["so","erp_so","iyear","imonth"]).agg({"order_qty":"sum","real_qty":"sum","amt":"sum"})
        #     group_date = pd.DataFrame(group_date).reset_index()
        #
        #     print(group_date.head(5).to_markdown())
        #     group_date.to_excel("data/it-sum.xlsx",index=False)
        # else:
        #     print("未输出按月汇总文件！")
    else:
        print("退出IT文件处理！")
        pass



def hebing():
    stock_picking=pd.read_csv("data/stock_picking.csv",encoding="utf-8")
    stock_move = pd.read_csv("data/stock_move1.csv", encoding="utf-8")
    product = pd.read_csv("data/Result_49.csv", encoding="utf-8")

    product.rename(columns={"id":"product_id","name":"product_name"},inplace=True )

    # 删除空行
    stock_move.dropna()

    df1=stock_picking.merge(stock_move,how="inner",left_on="id",right_on="picking_id")
    print(df1.head(10).to_markdown())

    df2 = df1.merge(product, how="inner", on="product_id")
    df2["default_code"] = "["+df2["default_code"]+"]"

    print(df2.head(10).to_markdown())

    df3 = df2[["origin","default_code","product_name","product_uom_qty"]]

    print(df3.head(10).to_markdown())
    
    df3.to_excel("data/it行项目2.xlsx")

def hebing2():
    stock_pick = pd.read_csv("data/Result_4.csv", encoding="utf-8")
    product = pd.read_csv("data/Result_49.csv", encoding="utf-8")

    product.rename(columns={"id": "product_id", "name": "product_name"}, inplace=True)

    stock_pick["product_uom_qty"] = stock_pick["product_uom_qty"].astype(int)
    # 删除空行
    # stock_move.dropna()
    #
    # df1 = stock_picking.merge(stock_move, how="inner", left_on="id", right_on="picking_id")
    # print(df1.head(10).to_markdown())

    df1 = stock_pick.merge(product, how="inner", on="product_id")
    df1["default_code"] = "[" + df1["default_code"] + "]"

    print(df1.head(10).to_markdown())

    df2 = df1[["origin", "default_code", "product_name", "product_uom_qty"]]

    print(df2.head(10).to_markdown())

    df2.to_excel("data/it行项目3.xlsx")

def open_file():
    dlg = win32ui.CreateFileDialog(1)  # 1表示打开文件对话框
    dlg.SetOFNInitialDir('E:/Python')  # 设置打开文件对话框中的初始显示目录
    dlg.DoModal()
    filename = dlg.GetPathName()  # 获取选择的文件名称
    print("filename=",filename)
    print("read ok")
    return filename

def file_cover():
    file_path = open_file()
    df = pd.read_excel(file_path, dtype=str)
    df.to_pickle("data/it-item_table.pkl")


if __name__=="__main__" :

    # hebing2()
    # get_fn2019()
    get_fn()
    get_it()
    # file_cover()