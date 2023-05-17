# __coding=utf8__

import pandas as pd
import os
import xlrd
import math

# 定义参考文件路径
shop_file = r"D:\work\OMS理帐\21年店铺.xlsx"
sku_file = r"D:\work\OMS理帐\2021年产品明细.xls"
product_file = r"D:\work\OMS理帐\product.template.xls"
warehouse_file = r"D:\work\OMS理帐\21年采购仓库.xls"

file_columns_list = []


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
                        if filename.find(key) >= 0:  # 只做文件名的过滤
                            _files.append(path)
                else:
                    _files.append(path)

    # print(_files)
    # 返回一个文件列表 list
    return _files


def convert_sales(filename, targe_dir, newfile):
    # 源头
    # 主体	主体	2C\2C发出商品	订单日期	订单日期	店铺名		固定	固定			内部编码	单位	实际卖出	平均价格	税率

    # 读取参考表格
    df_shop = pd.read_excel(shop_file)
    df_sku = pd.read_excel(sku_file)
    df_warehouse = pd.read_excel(warehouse_file)

    # 原始文件
    df = pd.read_excel(filename)
    # print(df.to_markdown())
    print(df.head(10).to_markdown())

    cnt = df.shape[0]
    print("原表行数:", df.shape[0])

    df_sales = df.copy()
    sale_type = "2C发出商品" if filename.find("2C发出商品") >= 0 else "2C"
    df_sales.rename({"主体": "配送仓库", "店铺": "客户"}, inplace=True)

    df_product = pd.read_excel(product_file)
    df_sales["商家编码"] = df_sales["商家编码"].astype(str)
    df_product["条码"] = df_product["条码"].astype(str)
    df_sales = df_sales.merge(df_product[["条码", "名称"]], how="left", left_on=["商家编码"], right_on=["条码"])

    if (df_sales[df_sales["名称"].isnull()].shape[0] > 0):
        df_sales[df_sales["名称"].isnull()].to_excel(newfile.replace(".xls", "_销售条码异常.xls"))
    else:
        print("条码检查合格")

    del df_sales["条码"]

    df_sales["进销存标识"] = sale_type
    df_sales["承诺日期"] = df_sales["支付日期"]
    df_sales["订单日期"] = df_sales["支付日期"]
    df_sales["价格表"] = ""
    df_sales["跟单员"] = "陆俊秀"
    df_sales["业务团队"] = "潘勤"
    df_sales["源单据"] = ""
    df_sales["客户参考"] = ""

    df_sales = df_sales.merge(df_shop[["平台", "OMS店铺名称", "Odoo店铺名称"]], how="left", left_on=["平台", "店铺"],
                              right_on=["平台", "OMS店铺名称"])
    print("转换1", "成功" if cnt == df.shape[0] else "失败:{}-{}".format(cnt, df_sales.shape[0]))

    print("sku_x")
    print(df_sales.head(10).to_markdown())
    df_sales = df_sales.merge(df_sku[["条码", "内部参考", "计量单位/显示名称"]], how="left", left_on=["商家编码"],
                              right_on=["条码"])
    print("转换2", "成功" if cnt == df.shape[0] else "失败:{}-{}".format(cnt, df_sales.shape[0]))

    print("sku_y")
    print(df_sales.head(10).to_markdown())

    # df_warehouse.rename(columns={"公司": "仓库公司"}, inplace=True)
    df_sales = df_sales.merge(df_warehouse[["公司", "仓库/显示名称"]], how="left", left_on=["主体"],
                              right_on=["公司"])
    # df_sales = df_sales.merge(df_warehouse[["仓库公司", "仓库/显示名称"]], how="left", left_on=["主体"], right_on=["仓库公司"])

    # 重命名
    df_sales.rename(
        columns={"Odoo店铺名称": "客户", "内部参考": "订单行/产品", "计量单位/显示名称": "订单行/计量单位", "实际卖出": "订单行/订购数量", "平均价格": "订单行/单价",
                 "仓库/显示名称": "配送仓库"},
        inplace=True)

    # 订单行/产品
    # df_sales=df_sales.merge(df_sku[["条码","内部参考"]],how="left",left_on=["商家编码"],right_on=["条码"])

    # print(df_sales[df_sales["税率"].isnull()])
    # df_sales["税率"].fillna(99, inplace=True)

    df_sales["订单行/税率"] = df_sales["税率"].apply(lambda x: "税收{}％（含）".format(int(x * 100)))
    df_sales["订单行/关税税额"] = ""
    df_sales["订单行/报关单号"] = ""
    df_sales["订单行/汇率"] = ""
    df_sales["公司"] = df_sales["主体"]
    # df_sales["配送仓库"] = df_warehouse["仓库/显示名称"]

    # df_sales["配送仓库"] = df_sales['仓库/显示名称']

    print("结果行数:", df_sales.shape[0])

    print("转换 ", "成功" if cnt == df_sales.shape[0] else "失败")

    print("查看结果:")
    print(df_sales.head(10).to_markdown())
    df_sales = df_sales[
        ["公司", "配送仓库", "进销存标识", "订单日期", "承诺日期", "客户", "价格表", "跟单员", "业务团队", "源单据", "客户参考", "订单行/产品", "订单行/计量单位",
         "订单行/订购数量", "订单行/单价", "订单行/税率", "订单行/关税税额", "订单行/报关单号", "订单行/汇率"]]

    # mubiao = r"D:\work\OMS理帐\2C发出商品"
    df_sales.to_excel("{}\{}".format(targe_dir, newfile), index=False)

    # 目标
    #  公司	配送仓库	进销存标识	订单日期	承诺日期	客户	价格表	跟单员	业务团队	源单据	客户参考	订单行/产品	订单行/计量单位	订单行/订购数量	订单行/单价	订单行/税率


def convert_purchase(filename, targe_dir, newfile):
    # 源头
    # 主体	主体	2C\2C发出商品	订单日期	订单日期	店铺名		固定	固定			内部编码	单位	实际卖出	平均价格	税率

    # 读取参考表格

    df_shop = pd.read_excel(shop_file)
    df_sku = pd.read_excel(sku_file)
    df_warehouse = pd.read_excel(warehouse_file)

    # 公司	交货到/数据库 ID

    # 原始文件
    # filename = r"C:\Users\sjit27\Desktop\文件\212C铲喜官（深圳）日用品有限公司\212C铲喜官（深圳）日用品有限公司\2C铲喜官（深圳）日用品有限公司销售订单.xlsx"
    df = pd.read_excel(filename)
    print(df.head(10).to_markdown())

    cnt = df.shape[0]
    print("原表行数:", df.shape[0])

    df_purcharse = df.copy()
    sale_type = "2C发出商品" if filename.find("2C发出商品") >= 0 else "2C"
    # df_purcharse["公司"] = df_purcharse["供应商"]

    df_product = pd.read_excel(product_file)
    df_purcharse["商家编码"] = df_purcharse["商家编码"].astype(str)
    df_product["条码"] = df_product["条码"].astype(str)
    df_purcharse = df_purcharse.merge(df_product[["条码", "名称"]], how="left", left_on=["商家编码"], right_on=["条码"])

    print("条码为空")
    print(df_purcharse[df_purcharse["条码"].isnull()].to_markdown())

    if (df_purcharse[df_purcharse["名称"].isnull()].shape[0] > 0):
        df_purcharse[df_purcharse["名称"].isnull()].to_excel(newfile.replace(".xls", "条码异常.xls"))
    else:
        print("条码检查合格")

    df_purcharse["公司"] = df_purcharse["主体"]
    df_purcharse["供应商"] = "深圳市麦凯莱科技有限公司"
    df_purcharse["date_order"] = df_purcharse["订单日期"]
    df_purcharse["进销存标识"] = sale_type
    df_purcharse["订单行/计划日期"] = df_purcharse["订单日期"]
    df_purcharse["订单行/单价"] = df_purcharse["采购含税单价"]
    # df_purcharse["订单行/订购数量"] = df_purcharse["实际卖出"]
    df_purcharse["订单行/订购数量"] = df_purcharse["数量"]
    df_purcharse["币种"] = "CNY"
    df_purcharse["采购员"] = "公司"
    df_purcharse["订单行/关税税额"] = ""
    df_purcharse["订单行/报关单号"] = ""
    df_purcharse["订单行/汇率"] = ""
    df_purcharse["源单据"] = ""

    # 交货到/数据库 ID
    df_purcharse = df_purcharse.merge(df_warehouse[["公司", "交货到/数据库 ID"]], how="left", left_on=["公司"], right_on=["公司"])

    # df_purcharse = df_purcharse.merge(df_shop[["平台", "OMS店铺名称", "Odoo店铺名称"]], how="left", left_on=["平台", "店铺"],
    #                                   right_on=["平台", "OMS店铺名称"])
    print("转换1", "成功" if cnt == df.shape[0] else "失败:{}-{}".format(cnt, df_purcharse.shape[0]))
    df_purcharse["商家编码"] = df_purcharse["商家编码"].astype(str)
    df_purcharse = df_purcharse.merge(df_sku[~df_sku["条码"].isnull()][["条码", "内部参考", "计量单位/显示名称"]], how="left",
                                      left_on=["商家编码"], right_on=["条码"])
    print("转换2", "成功" if cnt == df.shape[0] else "失败:{}-{}".format(cnt, df_purcharse.shape[0]))

    # 重命名
    df_purcharse.rename(
        columns={"内部参考": "订单行/产品", "计量单位/显示名称": "订单行/计量单位", "实际卖出数量": "订单行/订购数量"},
        inplace=True)

    # 订单行/产品
    df_purcharse["订单行/税率"] = df_purcharse["税率"].apply(lambda x: "税收{}％（含）".format(int(x * 100)))
    df_purcharse["订单行/关税税额"] = ""
    df_purcharse["订单行/报关单号"] = ""
    df_purcharse["订单行/汇率"] = ""

    print("结果行数:", df_purcharse.shape[0])

    print("转换 ", "成功" if cnt == df_purcharse.shape[0] else "失败")

    print("查看结果:")
    print(df_purcharse.head(10).to_markdown())

    df_purcharse = df_purcharse[
        ["公司", "交货到/数据库 ID", "进销存标识", "供应商", "date_order", "采购员", "订单行/计划日期", "币种", "订单行/产品", "订单行/计量单位", "订单行/订购数量",
         "订单行/单价", "订单行/税率", "订单行/关税税额", "订单行/报关单号", "订单行/汇率", "源单据"]]

    print("{}\{}".format(targe_dir, newfile))
    df_purcharse.to_excel("{}\{}".format(targe_dir, newfile), index=False)


def convert_all_purchase(rootdir, targe_dir, filekey):
    filelist = list_all_files(rootdir, filekey)

    print(filelist)
    for _file in filelist:
        print("文件名:", _file)
        new_filename = _file.split("\\")[-1]
        if new_filename.find("采购订单") >= 0:
            print("新文件名:", new_filename)
            convert_purchase(_file, targe_dir, new_filename.replace(".xlsx", "_转换后.xlsx"))


def convert_all_sales(rootdir, targe_dir, filekey):
    filelist = list_all_files(rootdir, filekey)

    print(filelist)
    for _file in filelist:
        print("文件名:", _file)
        # new_filename=os.path.splitext(_file)[0]
        new_filename = _file.split("\\")[-1]
        if new_filename.find("采购订单") < 0:
            print("新文件名:", new_filename)
            convert_sales(_file, targe_dir, new_filename.replace(".xlsx", "_销售订单_转换后.xlsx"))


if __name__ == "__main__":
    # 来源数据
    source_dir = r"D:\work\OMS理帐\转换中"
    # 目标目录
    targe_dir = r"D:\work\OMS理帐\转换odoo格式后"
    convert_all_purchase(source_dir, targe_dir, '')
    convert_all_sales(source_dir, targe_dir, '')
    print("ok")
