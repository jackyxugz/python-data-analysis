# __coding=utf8__

import pandas as pd
import sys
import os
import time
import os.path


def convert(filename, newfile):
    # 源头
    # 主体	主体	2C\2C发出商品	订单日期	订单日期	店铺名		固定	固定			内部编码	单位	实际卖出	平均价格	税率

    # 读取参考表格
    df_shop = pd.read_excel(r"D:\work\OMS理帐\21年店铺.xlsx")
    df_sku = pd.read_excel(r"D:\work\OMS理帐\2021年产品明细.xls")

    # 原始文件
    # filename = r"C:\Users\sjit27\Desktop\文件\212C铲喜官（深圳）日用品有限公司\212C铲喜官（深圳）日用品有限公司\2C铲喜官（深圳）日用品有限公司销售订单.xlsx"
    df = pd.read_excel(filename)
    print(df.to_markdown())

    cnt = df.shape[0]
    print("原表行数:", df.shape[0])

    df_purcharse = df.copy()
    sale_type = "2C发出商品" if filename.find("2C发出商品") else "2C"
    df_purcharse.rename({"主体": "配送仓库", "店铺": "客户"}, inplace=True)
    df_purcharse["进销存标识"] = sale_type
    df_purcharse["承诺日期"] = df_purcharse["支付日期"]
    df_purcharse["订单日期"] = df_purcharse["支付日期"]
    df_purcharse["价格表"] = ""
    df_purcharse["跟单员"] = "陆俊秀"
    df_purcharse["业务团队"] = "潘勤"
    df_purcharse["源单据"] = "潘勤"
    df_purcharse["客户参考"] = ""

    df_purcharse = df_purcharse.merge(df_shop[["平台", "OMS店铺名称", "Odoo店铺名称"]], how="left", left_on=["平台", "店铺"],
                                      right_on=["平台", "OMS店铺名称"])
    print("转换1", "成功" if cnt == df.shape[0] else "失败:{}-{}".format(cnt, df_purcharse.shape[0]))

    df_purcharse["商家编码"]=df_purcharse["商家编码"].astype(str)
    # df_sku["条码"]=df_sku["条码"].astype(str)
    df_purcharse = df_purcharse.merge(df_sku[["条码", "内部参考", "计量单位/显示名称"]], how="left", left_on=["商家编码"],
                                      right_on=["条码"])
    print("转换2", "成功" if cnt == df.shape[0] else "失败:{}-{}".format(cnt, df_purcharse.shape[0]))

    # 重命名
    df_purcharse.rename(
        columns={"Odoo店铺名称": "客户", "内部参考": "订单行/产品", "计量单位/显示名称": "订单行/计量单位", "实际卖出数量": "订单行/订购数量", "平均价格": "订单行/单价"},
        inplace=True)

    # 订单行/产品
    # df_purcharse=df_purcharse.merge(df_sku[["条码","内部参考"]],how="left",left_on=["商家编码"],right_on=["条码"])

    df_purcharse["订单行/税率"] = df_purcharse["税率"].apply(lambda x: "税收{}％（含）".format(x * 100))
    df_purcharse["订单行/关税税额"] = ""
    df_purcharse["订单行/报关单号"] = ""
    df_purcharse["订单行/汇率"] = ""
    df_purcharse["公司"] = df_purcharse["主体"]
    df_purcharse["配送仓库"] = df_purcharse["主体"]

    print("结果行数:", df_purcharse.shape[0])

    print("转换 ", "成功" if cnt == df_purcharse.shape[0] else "失败")

    print("查看结果:")
    print(df_purcharse.to_markdown())
    df_purcharse = df_purcharse[
        ["公司", "配送仓库", "进销存标识", "订单日期", "承诺日期", "客户", "价格表", "跟单员", "业务团队", "源单据", "客户参考", "订单行/产品", "订单行/计量单位",
         "订单行/订购数量", "订单行/单价", "订单行/税率"]]

    mubiao = r"D:\work\OMS理帐\2C发出商品"
    df_purcharse.to_excel("{}\{}".format(mubiao, newfile), index=False)

    # 目标
    #  公司	配送仓库	进销存标识	订单日期	承诺日期	客户	价格表	跟单员	业务团队	源单据	客户参考	订单行/产品	订单行/计量单位	订单行/订购数量	订单行/单价	订单行/税率


if __name__ == "__main__":
    dir = r"D:\work\OMS理帐\212C樱语（深圳）日用品有限公司"
    convert(dir + r"\2C樱语（深圳）日用品有限公司销售订单.xlsx", "2C樱语（深圳）日用品有限公司销售订单1.xlsx")

    # combine_excel()
