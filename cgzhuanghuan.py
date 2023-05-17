# __coding=utf8__

import pandas as pd

# 定义参考文件路径
shop_file = r"Z:\it审计处理需求\odoo导入\21年店铺.xlsx"
sku_file = r"Z:\it审计处理需求\odoo导入\2021年产品明细.xls"
product_file = r"Z:\it审计处理需求\odoo导入\product.template.xls"
warehouse_file = r"Z:\it审计处理需求\odoo导入\21年采购仓库.xls"


def convert_sales(filename, targe_dir, newfile):
    # 源头
    # 主体	主体	2C\2C发出商品	订单日期	订单日期	店铺名		固定	固定			内部编码	单位	实际卖出	平均价格	税率

    # 读取参考表格
    df_shop = pd.read_excel(shop_file)
    df_sku = pd.read_excel(sku_file)

    # 原始文件
    df = pd.read_excel(filename)
    print(df.to_markdown())

    cnt = df.shape[0]
    print("原表行数:", df.shape[0])

    df_sales = df.copy()
    sale_type = "2C发出商品" if filename.find("2C发出商品") else "2C"
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
    df_sales["源单据"] = "潘勤"
    df_sales["客户参考"] = ""

    df_sales = df_sales.merge(df_shop[["平台", "OMS店铺名称", "Odoo店铺名称"]], how="left", left_on=["平台", "店铺"],
                              right_on=["平台", "OMS店铺名称"])
    print("转换1", "成功" if cnt == df.shape[0] else "失败:{}-{}".format(cnt, df_sales.shape[0]))

    df_sales = df_sales.merge(df_sku[["条码", "内部参考", "计量单位/显示名称"]], how="left", left_on=["商家编码"],
                              right_on=["条码"])
    print("转换2", "成功" if cnt == df.shape[0] else "失败:{}-{}".format(cnt, df_sales.shape[0]))

    # 重命名
    df_sales.rename(
        columns={"Odoo店铺名称": "客户", "内部参考": "订单行/产品", "计量单位/显示名称": "订单行/计量单位", "实际卖出": "订单行/订购数量", "平均价格": "订单行/单价"},
        inplace=True)

    # 订单行/产品
    # df_sales=df_sales.merge(df_sku[["条码","内部参考"]],how="left",left_on=["商家编码"],right_on=["条码"])

    df_sales["订单行/税率"] = df_sales["税率"].apply(lambda x: "税收{}％（含）".format(x * 100))
    df_sales["订单行/关税税额"] = ""
    df_sales["订单行/报关单号"] = ""
    df_sales["订单行/汇率"] = ""
    df_sales["公司"] = df_sales["主体"]
    df_sales["配送仓库"] = df_sales["主体"]

    print("结果行数:", df_sales.shape[0])

    print("转换 ", "成功" if cnt == df_sales.shape[0] else "失败")

    print("查看结果:")
    print(df_sales.to_markdown())
    df_sales = df_sales[
        ["公司", "配送仓库", "进销存标识", "订单日期", "承诺日期", "客户", "价格表", "跟单员", "业务团队", "源单据", "客户参考", "订单行/产品", "订单行/计量单位",
         "订单行/订购数量", "订单行/单价", "订单行/税率"]]

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
    print(df.to_markdown())

    cnt = df.shape[0]
    print("原表行数:", df.shape[0])

    df_purcharse = df.copy()
    sale_type = "2C发出商品" if filename.find("2C发出商品") else "2C"
    # df_purcharse["公司"] = df_purcharse["供应商"]

    df_product = pd.read_excel(product_file)
    df_purcharse["商家编码"] = df_purcharse["商家编码"].astype(str)
    df_product["条码"] = df_product["条码"].astype(str)
    df_purcharse = df_purcharse.merge(df_product[["条码", "名称"]], how="left", left_on=["商家编码"], right_on=["条码"])

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
    df_purcharse["币种"] = "RMB"
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
                                      left_on=["商家编码"],
                                      right_on=["条码"])
    print("转换2", "成功" if cnt == df.shape[0] else "失败:{}-{}".format(cnt, df_purcharse.shape[0]))

    # 重命名
    df_purcharse.rename(
        columns={"内部参考": "订单行/产品", "计量单位/显示名称": "订单行/计量单位", "实际卖出数量": "订单行/订购数量"},
        inplace=True)

    # 订单行/产品
    df_purcharse["订单行/税率"] = df_purcharse["税率"].apply(lambda x: "税收{}％（含）".format(x * 100))
    df_purcharse["订单行/关税税额"] = ""
    df_purcharse["订单行/报关单号"] = ""
    df_purcharse["订单行/汇率"] = ""

    print("结果行数:", df_purcharse.shape[0])

    print("转换 ", "成功" if cnt == df_purcharse.shape[0] else "失败")

    print("查看结果:")
    print(df_purcharse.to_markdown())

    df_purcharse = df_purcharse[
        ["公司", "交货到/数据库 ID", "进销存标识", "供应商", "date_order", "采购员", "订单行/计划日期", "币种", "订单行/产品", "订单行/计量单位", "订单行/订购数量",
         "订单行/单价", "订单行/税率",
         "订单行/关税税额", "订单行/报关单号", "订单行/汇率", "源单据"]]

    # mubiao = r"D:\work\OMS理帐\2C发出商品"
    print("{}\{}".format(targe_dir, newfile))
    df_purcharse.to_excel("{}\{}".format(targe_dir, newfile), index=False)


if __name__ == "__main__":
    # 来源数据
    # source_dir = r"Z:\it审计处理需求\odoo导入\转换前\2C"
    source_dir = r"Z:\it审计处理需求\odoo导入\转换前\2C发出商品"
    # 目标目录
    targe_dir = r"D:\odoo数据处理"

    # convert_purchase("{}\{}".format(source_dir, "2C发出商品芭葆兔（深圳）日用品有限公司_采购订单.xlsx"), targe_dir,"2C发出商品芭葆兔（深圳）日用品有限公司_采购订单_转换后.xlsx")
    # convert_sales("{}\{}".format(source_dir, "2C樱语（深圳）日用品有限公司.xlsx"), targe_dir, "2C樱语（深圳）日用品有限公司_销售订单_转换后.xlsx")
    convert_sales("{}\{}".format(source_dir, "2C一片珍芯（深圳）化妆品有限公司.xlsx"), targe_dir, "2C一片珍芯（深圳）化妆品有限公司.xlsx")
    convert_purchase("{}\{}".format(source_dir, "2C一片珍芯（深圳）化妆品有限公司_采购订单.xlsx"), targe_dir,
                     "2C一片珍芯（深圳）化妆品有限公司_采购订单.xlsx")
    convert_sales("{}\{}".format(source_dir, "2C不是很酷（深圳）服装有限公司.xlsx"), targe_dir, "2C不是很酷（深圳）服装有限公司.xlsx")
    convert_purchase("{}\{}".format(source_dir, "2C不是很酷（深圳）服装有限公司_采购订单.xlsx"), targe_dir, "2C不是很酷（深圳）服装有限公司_采购订单.xlsx")
    convert_sales("{}\{}".format(source_dir, "2C不酷（深圳）商贸有限公司.xlsx"), targe_dir, "2C不酷（深圳）商贸有限公司.xlsx")
    convert_purchase("{}\{}".format(source_dir, "2C不酷（深圳）商贸有限公司_采购订单.xlsx"), targe_dir, "2C不酷（深圳）商贸有限公司_采购订单.xlsx")
    convert_sales("{}\{}".format(source_dir, "2C冰川女神（深圳）化妆品有限公司.xlsx"), targe_dir, "2C冰川女神（深圳）化妆品有限公司.xlsx")
    convert_purchase("{}\{}".format(source_dir, "2C冰川女神（深圳）化妆品有限公司_采购订单.xlsx"), targe_dir,
                     "2C冰川女神（深圳）化妆品有限公司_采购订单.xlsx")
    # convert_sales("{}\{}".format(source_dir, "2C勃狄（深圳）化妆品有限公司_error_价格偏高请检查.xls"), targe_dir,"2C勃狄（深圳）化妆品有限公司_error_价格偏高请检查.xlsx")
    # convert_purchase("{}\{}".format(source_dir, "2C勃狄（深圳）化妆品有限公司_error_税率问题.xls"), targe_dir, "2C勃狄（深圳）化妆品有限公司_error_税率问题.xlsx")
    convert_sales("{}\{}".format(source_dir, "2C可瘾（广州）化妆品有限公司.xlsx"), targe_dir, "2C可瘾（广州）化妆品有限公司.xlsx")
    convert_purchase("{}\{}".format(source_dir, "2C可瘾（广州）化妆品有限公司_采购订单.xlsx"), targe_dir, "2C可瘾（广州）化妆品有限公司_采购订单.xlsx")
    convert_sales("{}\{}".format(source_dir, "2C可隐（深圳）化妆品有限公司.xlsx"), targe_dir, "2C可隐（深圳）化妆品有限公司.xlsx")
    convert_purchase("{}\{}".format(source_dir, "2C可隐（深圳）化妆品有限公司_采购订单.xlsx"), targe_dir, "2C可隐（深圳）化妆品有限公司_采购订单.xlsx")
    convert_sales("{}\{}".format(source_dir, "2C吉维儿（深圳）化妆品有限公司.xlsx"), targe_dir, "2C吉维儿（深圳）化妆品有限公司.xlsx")
    convert_purchase("{}\{}".format(source_dir, "2C吉维儿（深圳）化妆品有限公司_采购订单.xlsx"), targe_dir, "2C吉维儿（深圳）化妆品有限公司_采购订单.xlsx")
    convert_sales("{}\{}".format(source_dir, "2C商魂信息（深圳）有限公司.xlsx"), targe_dir, "2C商魂信息（深圳）有限公司.xlsx")
    convert_purchase("{}\{}".format(source_dir, "2C商魂信息（深圳）有限公司_采购订单.xlsx"), targe_dir, "2C商魂信息（深圳）有限公司_采购订单.xlsx")
    convert_sales("{}\{}".format(source_dir, "2C喝啥（深圳）食品有限公司.xlsx"), targe_dir, "2C喝啥（深圳）食品有限公司.xlsx")
    convert_purchase("{}\{}".format(source_dir, "2C喝啥（深圳）食品有限公司_采购订单.xls"), targe_dir, "2C喝啥（深圳）食品有限公司_采购订单.xlsx")
    convert_sales("{}\{}".format(source_dir, "2C多瑞（深圳）日用品有限公司.xlsx"), targe_dir, "2C多瑞（深圳）日用品有限公司.xlsx")
    convert_purchase("{}\{}".format(source_dir, "2C多瑞（深圳）日用品有限公司_采购订单.xlsx"), targe_dir, "2C多瑞（深圳）日用品有限公司_采购订单.xlsx")
    convert_sales("{}\{}".format(source_dir, "2C妈咪港湾（深圳）化妆品有限公司.xlsx"), targe_dir, "2C妈咪港湾（深圳）化妆品有限公司.xlsx")
    convert_purchase("{}\{}".format(source_dir, "2C妈咪港湾（深圳）化妆品有限公司_采购订单.xlsx"), targe_dir,
                     "2C妈咪港湾（深圳）化妆品有限公司_采购订单.xlsx")
    convert_sales("{}\{}".format(source_dir, "2C宅星人（深圳）食品有限公司.xlsx"), targe_dir, "2C宅星人（深圳）食品有限公司.xlsx")
    convert_purchase("{}\{}".format(source_dir, "2C宅星人（深圳）食品有限公司_采购订单.xlsx"), targe_dir, "2C宅星人（深圳）食品有限公司_采购订单.xlsx")
    convert_sales("{}\{}".format(source_dir, "2C宝贝港湾（深圳）化妆品有限公司.xlsx"), targe_dir, "2C宝贝港湾（深圳）化妆品有限公司.xlsx")
    convert_purchase("{}\{}".format(source_dir, "2C宝贝港湾（深圳）化妆品有限公司_采购订单.xlsx"), targe_dir,
                     "2C宝贝港湾（深圳）化妆品有限公司_采购订单.xlsx")
    convert_sales("{}\{}".format(source_dir, "2C宝贝配方师（深圳）日用品有限公司.xlsx"), targe_dir, "2C宝贝配方师（深圳）日用品有限公司.xlsx")
    convert_purchase("{}\{}".format(source_dir, "2C宝贝配方师（深圳）日用品有限公司_采购订单.xlsx"), targe_dir,
                     "2C宝贝配方师（深圳）日用品有限公司_采购订单.xlsx")
    convert_sales("{}\{}".format(source_dir, "2C宝贝魔术师（深圳）日用品有限公司.xlsx"), targe_dir, "2C宝贝魔术师（深圳）日用品有限公司.xlsx")
    convert_purchase("{}\{}".format(source_dir, "2C宝贝魔术师（深圳）日用品有限公司_采购订单.xlsx"), targe_dir,
                     "2C宝贝魔术师（深圳）日用品有限公司_采购订单.xlsx")
    convert_sales("{}\{}".format(source_dir, "2C尚隐（深圳）计生用品有限公司.xlsx"), targe_dir, "2C尚隐（深圳）计生用品有限公司.xlsx")
    convert_purchase("{}\{}".format(source_dir, "2C尚隐（深圳）计生用品有限公司_采购订单.xlsx"), targe_dir, "2C尚隐（深圳）计生用品有限公司_采购订单.xlsx")
    convert_sales("{}\{}".format(source_dir, "2C平湖宏炽贸易有限公司.xlsx"), targe_dir, "2C平湖宏炽贸易有限公司.xlsx")
    convert_purchase("{}\{}".format(source_dir, "2C平湖宏炽贸易有限公司_error_价格偏高请检查.xlsx"), targe_dir,
                     "2C平湖宏炽贸易有限公司_error_价格偏高请检查.xlsx")
    convert_purchase("{}\{}".format(source_dir, "2C平湖宏炽贸易有限公司_采购订单.xlsx"), targe_dir, "2C平湖宏炽贸易有限公司_采购订单.xlsx")
    convert_sales("{}\{}".format(source_dir, "2C平湖鑫桂贸易有限公司.xlsx"), targe_dir, "2C平湖鑫桂贸易有限公司.xlsx")
    convert_purchase("{}\{}".format(source_dir, "2C平湖鑫桂贸易有限公司_error_价格偏高请检查.xlsx"), targe_dir,
                     "2C平湖鑫桂贸易有限公司_error_价格偏高请检查.xlsx")
    convert_purchase("{}\{}".format(source_dir, "2C平湖鑫桂贸易有限公司_采购订单.xlsx"), targe_dir, "2C平湖鑫桂贸易有限公司_采购订单.xlsx")
    convert_sales("{}\{}".format(source_dir, "2C平湖鲁文国际贸易有限公司.xlsx"), targe_dir, "2C平湖鲁文国际贸易有限公司.xlsx")
    convert_purchase("{}\{}".format(source_dir, "2C平湖鲁文国际贸易有限公司_采购订单.xlsx"), targe_dir, "2C平湖鲁文国际贸易有限公司_采购订单.xlsx")
    convert_sales("{}\{}".format(source_dir, "2C广州市尚西国际贸易有限公司.xlsx"), targe_dir, "2C广州市尚西国际贸易有限公司.xlsx")
    convert_purchase("{}\{}".format(source_dir, "2C广州市尚西国际贸易有限公司_采购订单.xlsx"), targe_dir, "2C广州市尚西国际贸易有限公司_采购订单.xlsx")
    convert_sales("{}\{}".format(source_dir, "2C广州斌闻贸易有限公司.xlsx"), targe_dir, "2C广州斌闻贸易有限公司.xlsx")
    convert_purchase("{}\{}".format(source_dir, "2C广州斌闻贸易有限公司_采购订单.xlsx"), targe_dir, "2C广州斌闻贸易有限公司_采购订单.xlsx")
    convert_sales("{}\{}".format(source_dir, "2C广州驰骄服饰有限公司.xlsx"), targe_dir, "2C广州驰骄服饰有限公司.xlsx")
    convert_purchase("{}\{}".format(source_dir, "2C广州驰骄服饰有限公司_采购订单.xlsx"), targe_dir, "2C广州驰骄服饰有限公司_采购订单.xlsx")
    convert_sales("{}\{}".format(source_dir, "2C惠优购（深圳）日用品有限公司.xlsx"), targe_dir, "2C惠优购（深圳）日用品有限公司.xlsx")
    convert_purchase("{}\{}".format(source_dir, "2C惠优购（深圳）日用品有限公司_采购订单.xlsx"), targe_dir, "2C惠优购（深圳）日用品有限公司_采购订单.xlsx")
    convert_sales("{}\{}".format(source_dir, "2C戏酱（深圳）食品有限公司.xlsx"), targe_dir, "2C戏酱（深圳）食品有限公司.xlsx")
    convert_purchase("{}\{}".format(source_dir, "2C戏酱（深圳）食品有限公司_采购订单.xlsx"), targe_dir, "2C戏酱（深圳）食品有限公司_采购订单.xlsx")
    convert_sales("{}\{}".format(source_dir, "2C控师（深圳）化妆品有限公司.xlsx"), targe_dir, "2C控师（深圳）化妆品有限公司.xlsx")
    convert_purchase("{}\{}".format(source_dir, "2C控师（深圳）化妆品有限公司_采购订单.xlsx"), targe_dir, "2C控师（深圳）化妆品有限公司_采购订单.xlsx")
    convert_sales("{}\{}".format(source_dir, "2C播地艾（广州）化妆品有限公司.xlsx"), targe_dir, "2C播地艾（广州）化妆品有限公司.xlsx")
    convert_purchase("{}\{}".format(source_dir, "2C播地艾（广州）化妆品有限公司_采购订单.xlsx"), targe_dir, "2C播地艾（广州）化妆品有限公司_采购订单.xlsx")
    convert_sales("{}\{}".format(source_dir, "2C无极爽（深圳）日用品有限公司.xlsx"), targe_dir, "2C无极爽（深圳）日用品有限公司.xlsx")
    convert_purchase("{}\{}".format(source_dir, "2C无极爽（深圳）日用品有限公司_采购订单.xlsx"), targe_dir, "2C无极爽（深圳）日用品有限公司_采购订单.xlsx")
    convert_sales("{}\{}".format(source_dir, "2C末隐师（广州）化妆品有限公司.xlsx"), targe_dir, "2C末隐师（广州）化妆品有限公司.xlsx")
    convert_purchase("{}\{}".format(source_dir, "2C末隐师（广州）化妆品有限公司_采购订单.xlsx"), targe_dir, "2C末隐师（广州）化妆品有限公司_采购订单.xlsx")
    convert_sales("{}\{}".format(source_dir, "2C柚选（深圳）化妆品有限公司.xlsx"), targe_dir, "2C柚选（深圳）化妆品有限公司.xlsx")
    convert_purchase("{}\{}".format(source_dir, "2C柚选（深圳）化妆品有限公司_采购订单.xlsx"), targe_dir, "2C柚选（深圳）化妆品有限公司_采购订单.xlsx")
    convert_sales("{}\{}".format(source_dir, "2C植之璨（深圳）化妆品有限公司.xlsx"), targe_dir, "2C植之璨（深圳）化妆品有限公司.xlsx")
    convert_purchase("{}\{}".format(source_dir, "2C植之璨（深圳）化妆品有限公司_采购订单.xlsx"), targe_dir, "2C植之璨（深圳）化妆品有限公司_采购订单.xlsx")
    convert_sales("{}\{}".format(source_dir, "2C樱语（深圳）日用品有限公司.xlsx"), targe_dir, "2C樱语（深圳）日用品有限公司.xlsx")
    convert_purchase("{}\{}".format(source_dir, "2C樱语（深圳）日用品有限公司_采购订单.xlsx"), targe_dir, "2C樱语（深圳）日用品有限公司_采购订单.xlsx")
    convert_sales("{}\{}".format(source_dir, "2C泉州市航星贸易有限公司.xlsx"), targe_dir, "2C泉州市航星贸易有限公司.xlsx")
    convert_purchase("{}\{}".format(source_dir, "2C泉州市航星贸易有限公司_采购订单.xlsx"), targe_dir, "2C泉州市航星贸易有限公司_采购订单.xlsx")
    convert_sales("{}\{}".format(source_dir, "2C泡研（深圳）化妆品有限公司.xlsx"), targe_dir, "2C泡研（深圳）化妆品有限公司.xlsx")
    convert_purchase("{}\{}".format(source_dir, "2C泡研（深圳）化妆品有限公司_采购订单.xlsx"), targe_dir, "2C泡研（深圳）化妆品有限公司_采购订单.xlsx")
    convert_sales("{}\{}".format(source_dir, "2C浙江魔湾电子有限公司.xlsx"), targe_dir, "2C浙江魔湾电子有限公司.xlsx")
    convert_purchase("{}\{}".format(source_dir, "2C浙江魔湾电子有限公司_采购订单.xlsx"), targe_dir, "2C浙江魔湾电子有限公司_采购订单.xlsx")
    convert_sales("{}\{}".format(source_dir, "2C深圳大前海物流有限公司.xlsx"), targe_dir, "2C深圳大前海物流有限公司.xlsx")
    convert_purchase("{}\{}".format(source_dir, "2C深圳大前海物流有限公司_采购订单.xlsx"), targe_dir, "2C深圳大前海物流有限公司_采购订单.xlsx")
    # convert_sales("{}\{}".format(source_dir, "2C深圳市二十四小时七天商贸有限公司.xls"), targe_dir, "2C深圳市二十四小时七天商贸有限公司.xlsx")
    # convert_purchase("{}\{}".format(source_dir, "2C深圳市二十四小时七天商贸有限公司_error_税率问题.xls"), targe_dir,"2C深圳市二十四小时七天商贸有限公司_error_税率问题.xlsx")
    convert_sales("{}\{}".format(source_dir, "2C深圳市卖家优选实业有限公司.xlsx"), targe_dir, "2C深圳市卖家优选实业有限公司.xlsx")
    convert_purchase("{}\{}".format(source_dir, "2C深圳市卖家优选实业有限公司_采购订单.xlsx"), targe_dir, "2C深圳市卖家优选实业有限公司_采购订单.xlsx")
    convert_sales("{}\{}".format(source_dir, "2C深圳市卖家联合商贸有限公司.xlsx"), targe_dir, "2C深圳市卖家联合商贸有限公司.xlsx")
    convert_purchase("{}\{}".format(source_dir, "2C深圳市卖家联合商贸有限公司_采购订单.xlsx"), targe_dir, "2C深圳市卖家联合商贸有限公司_采购订单.xlsx")
    convert_sales("{}\{}".format(source_dir, "2C深圳市博滴日用品有限公司.xlsx"), targe_dir, "2C深圳市博滴日用品有限公司.xlsx")
    convert_purchase("{}\{}".format(source_dir, "2C深圳市博滴日用品有限公司_采购订单.xlsx"), targe_dir, "2C深圳市博滴日用品有限公司_采购订单.xlsx")
    convert_sales("{}\{}".format(source_dir, "2C深圳市白皮书文化传媒有限公司.xlsx"), targe_dir, "2C深圳市白皮书文化传媒有限公司.xlsx")
    convert_purchase("{}\{}".format(source_dir, "2C深圳市白皮书文化传媒有限公司_采购订单.xlsx"), targe_dir, "2C深圳市白皮书文化传媒有限公司_采购订单.xlsx")
    # convert_purchase("{}\{}".format(source_dir, "2C深圳市精酿商贸有限公司_error_税率问题.xls"), targe_dir, "2C深圳市精酿商贸有限公司_error_税率问题.xlsx")
    convert_sales("{}\{}".format(source_dir, "2C深圳市艾法商贸有限公司.xlsx"), targe_dir, "2C深圳市艾法商贸有限公司.xlsx")
    # convert_purchase("{}\{}".format(source_dir, "2C深圳市艾法商贸有限公司_error_价格偏高请检查.xls"), targe_dir, "2C深圳市艾法商贸有限公司_error_价格偏高请检查.xlsx")
    convert_purchase("{}\{}".format(source_dir, "2C深圳市艾法商贸有限公司_采购订单.xlsx"), targe_dir, "2C深圳市艾法商贸有限公司_采购订单.xlsx")
    convert_sales("{}\{}".format(source_dir, "2C深圳市配颜师生物科技有限公司.xlsx"), targe_dir, "2C深圳市配颜师生物科技有限公司.xlsx")
    convert_purchase("{}\{}".format(source_dir, "2C深圳市配颜师生物科技有限公司_采购订单.xlsx"), targe_dir, "2C深圳市配颜师生物科技有限公司_采购订单.xlsx")
    convert_sales("{}\{}".format(source_dir, "2C深圳市魔湾游戏科技有限公司.xlsx"), targe_dir, "2C深圳市魔湾游戏科技有限公司.xlsx")
    convert_purchase("{}\{}".format(source_dir, "2C深圳市魔湾游戏科技有限公司_采购订单.xlsx"), targe_dir, "2C深圳市魔湾游戏科技有限公司_采购订单.xlsx")
    convert_sales("{}\{}".format(source_dir, "2C深圳市麦凯莱科技有限公司.xlsx"), targe_dir, "2C深圳市麦凯莱科技有限公司.xlsx")
    # convert_purchase("{}\{}".format(source_dir, "2C深圳市麦凯莱科技有限公司_error_价格偏高请检查.xls"), targe_dir, "2C深圳市麦凯莱科技有限公司_error_价格偏高请检查.xlsx")
    convert_purchase("{}\{}".format(source_dir, "2C深圳市麦凯莱科技有限公司_采购订单.xlsx"), targe_dir, "2C深圳市麦凯莱科技有限公司_采购订单.xlsx")
    # convert_purchase("{}\{}".format(source_dir, "2C深圳樱岚护肤品有限公司_error_税率问题.xls"), targe_dir, "2C深圳樱岚护肤品有限公司_error_税率问题.xlsx")
    convert_purchase("{}\{}".format(source_dir, "2C深圳睿旗科技有限公司.xlsx"), targe_dir, "2C深圳睿旗科技有限公司.xlsx")
    # convert_purchase("{}\{}".format(source_dir, "2C深圳睿旗科技有限公司_error_价格偏高请检查.xls"), targe_dir,"2C深圳睿旗科技有限公司_error_价格偏高请检查.xlsx")
    convert_purchase("{}\{}".format(source_dir, "2C深圳睿旗科技有限公司_采购订单.xlsx"), targe_dir, "2C深圳睿旗科技有限公司_采购订单.xlsx")
    convert_sales("{}\{}".format(source_dir, "2C深圳造白化妆品有限公司.xlsx"), targe_dir, "2C深圳造白化妆品有限公司.xlsx")
    convert_purchase("{}\{}".format(source_dir, "2C深圳造白化妆品有限公司_采购订单.xlsx"), targe_dir, "2C深圳造白化妆品有限公司_采购订单.xlsx")
    convert_sales("{}\{}".format(source_dir, "2C深圳魔湾电子有限公司.xlsx"), targe_dir, "2C深圳魔湾电子有限公司.xlsx")
    convert_purchase("{}\{}".format(source_dir, "2C深圳魔湾电子有限公司_采购订单.xlsx"), targe_dir, "2C深圳魔湾电子有限公司_采购订单.xlsx")
    convert_sales("{}\{}".format(source_dir, "2C燃威（深圳）食品有限公司.xlsx"), targe_dir, "2C燃威（深圳）食品有限公司.xlsx")
    convert_purchase("{}\{}".format(source_dir, "2C燃威（深圳）食品有限公司_采购订单.xlsx"), targe_dir, "2C燃威（深圳）食品有限公司_采购订单.xlsx")
    convert_sales("{}\{}".format(source_dir, "2C珍芯漾肤（深圳）化妆品有限公司.xlsx"), targe_dir, "2C珍芯漾肤（深圳）化妆品有限公司.xlsx")
    convert_purchase("{}\{}".format(source_dir, "2C珍芯漾肤（深圳）化妆品有限公司_采购订单.xlsx"), targe_dir,
                     "2C珍芯漾肤（深圳）化妆品有限公司_采购订单.xlsx")
    convert_sales("{}\{}".format(source_dir, "2C白卿（深圳）化妆品有限公司.xlsx"), targe_dir, "2C白卿（深圳）化妆品有限公司.xlsx")
    convert_purchase("{}\{}".format(source_dir, "2C白卿（深圳）化妆品有限公司_采购订单.xlsx"), targe_dir, "2C白卿（深圳）化妆品有限公司_采购订单.xlsx")
    convert_sales("{}\{}".format(source_dir, "2C盈养泉（深圳）化妆品有限公司.xlsx"), targe_dir, "2C盈养泉（深圳）化妆品有限公司.xlsx")
    convert_purchase("{}\{}".format(source_dir, "2C盈养泉（深圳）化妆品有限公司_采购订单.xlsx"), targe_dir, "2C盈养泉（深圳）化妆品有限公司_采购订单.xlsx")
    convert_sales("{}\{}".format(source_dir, "2C秀美颜（广州）化妆品有限公司.xlsx"), targe_dir, "2C秀美颜（广州）化妆品有限公司.xlsx")
    convert_purchase("{}\{}".format(source_dir, "2C秀美颜（广州）化妆品有限公司_采购订单.xlsx"), targe_dir, "2C秀美颜（广州）化妆品有限公司_采购订单.xlsx")
    convert_sales("{}\{}".format(source_dir, "2C秀美颜（深圳）化妆品有限公司.xlsx"), targe_dir, "2C秀美颜（深圳）化妆品有限公司.xlsx")
    convert_purchase("{}\{}".format(source_dir, "2C秀美颜（深圳）化妆品有限公司_采购订单.xlsx"), targe_dir, "2C秀美颜（深圳）化妆品有限公司_采购订单.xlsx")
    convert_sales("{}\{}".format(source_dir, "2C肌密泉（深圳）化妆品有限公司.xlsx"), targe_dir, "2C肌密泉（深圳）化妆品有限公司.xlsx")
    convert_purchase("{}\{}".format(source_dir, "2C肌密泉（深圳）化妆品有限公司_采购订单.xlsx"), targe_dir, "2C肌密泉（深圳）化妆品有限公司_采购订单.xlsx")
    convert_sales("{}\{}".format(source_dir, "2C肌沫（深圳）化妆品有限公司.xlsx"), targe_dir, "2C肌沫（深圳）化妆品有限公司.xlsx")
    convert_purchase("{}\{}".format(source_dir, "2C肌沫（深圳）化妆品有限公司_采购订单.xlsx"), targe_dir, "2C肌沫（深圳）化妆品有限公司_采购订单.xlsx")
    convert_sales("{}\{}".format(source_dir, "2C肯妮诗（深圳）化妆品有限公司.xlsx"), targe_dir, "2C肯妮诗（深圳）化妆品有限公司.xlsx")
    convert_purchase("{}\{}".format(source_dir, "2C肯妮诗（深圳）化妆品有限公司_采购订单.xls"), targe_dir, "2C肯妮诗（深圳）化妆品有限公司_采购订单.xlsx")
    convert_sales("{}\{}".format(source_dir, "2C芭葆兔（深圳）日用品有限公司.xls"), targe_dir, "2C芭葆兔（深圳）日用品有限公司.xlsx")
    convert_purchase("{}\{}".format(source_dir, "2C芭葆兔（深圳）日用品有限公司_采购订单.xls"), targe_dir, "2C芭葆兔（深圳）日用品有限公司_采购订单.xlsx")
    convert_sales("{}\{}".format(source_dir, "2C若蘅（深圳）化妆品有限公司.xls"), targe_dir, "2C若蘅（深圳）化妆品有限公司.xlsx")
    convert_purchase("{}\{}".format(source_dir, "2C若蘅（深圳）化妆品有限公司_采购订单.xls"), targe_dir, "2C若蘅（深圳）化妆品有限公司_采购订单.xlsx")
    convert_sales("{}\{}".format(source_dir, "2C茉小桃（深圳）化妆品有限公司.xls"), targe_dir, "2C茉小桃（深圳）化妆品有限公司.xlsx")
    convert_purchase("{}\{}".format(source_dir, "2C茉小桃（深圳）化妆品有限公司_采购订单.xls"), targe_dir, "2C茉小桃（深圳）化妆品有限公司_采购订单.xlsx")
    # convert_purchase("{}\{}".format(source_dir, "2C茱莉珂丝（深圳）化妆品有限公司_error_缺少有效的数据.xls"), targe_dir, "2C茱莉珂丝（深圳）化妆品有限公司_error_缺少有效的数据.xlsx")
    convert_sales("{}\{}".format(source_dir, "2C萌洁齿（深圳）日用品有限公司.xls"), targe_dir, "2C萌洁齿（深圳）日用品有限公司.xlsx")
    convert_purchase("{}\{}".format(source_dir, "2C萌洁齿（深圳）日用品有限公司_采购订单.xls"), targe_dir, "2C萌洁齿（深圳）日用品有限公司_采购订单.xlsx")
    convert_sales("{}\{}".format(source_dir, "2C萦丝茧（深圳）化妆品有限公司.xls"), targe_dir, "2C萦丝茧（深圳）化妆品有限公司.xlsx")
    convert_purchase("{}\{}".format(source_dir, "2C萦丝茧（深圳）化妆品有限公司_采购订单.xls"), targe_dir, "2C萦丝茧（深圳）化妆品有限公司_采购订单.xlsx")
    convert_sales("{}\{}".format(source_dir, "2C补舍（深圳）食品有限公司.xls"), targe_dir, "2C补舍（深圳）食品有限公司.xlsx")
    convert_purchase("{}\{}".format(source_dir, "2C补舍（深圳）食品有限公司_采购订单.xls"), targe_dir, "2C补舍（深圳）食品有限公司_采购订单.xlsx")
    convert_sales("{}\{}".format(source_dir, "2C谷口（深圳）化妆品有限公司.xls"), targe_dir, "2C谷口（深圳）化妆品有限公司.xlsx")
    convert_purchase("{}\{}".format(source_dir, "2C谷口（深圳）化妆品有限公司_采购订单.xls"), targe_dir, "2C谷口（深圳）化妆品有限公司_采购订单.xlsx")
    convert_sales("{}\{}".format(source_dir, "2C贝贝港湾（深圳）化妆品有限公司.xls"), targe_dir, "2C贝贝港湾（深圳）化妆品有限公司.xlsx")
    convert_purchase("{}\{}".format(source_dir, "2C贝贝港湾（深圳）化妆品有限公司_采购订单.xls"), targe_dir, "2C贝贝港湾（深圳）化妆品有限公司_采购订单.xlsx")
    convert_sales("{}\{}".format(source_dir, "2C造味（深圳）食品有限公司.xls"), targe_dir, "2C造味（深圳）食品有限公司.xlsx")
    convert_purchase("{}\{}".format(source_dir, "2C造味（深圳）食品有限公司_采购订单.xls"), targe_dir, "2C造味（深圳）食品有限公司_采购订单.xlsx")
    convert_sales("{}\{}".format(source_dir, "2C造白（广州）化妆品有限公司.xls"), targe_dir, "2C造白（广州）化妆品有限公司.xlsx")
    convert_purchase("{}\{}".format(source_dir, "2C造白（广州）化妆品有限公司_采购订单.xls"), targe_dir, "2C造白（广州）化妆品有限公司_采购订单.xlsx")
    convert_sales("{}\{}".format(source_dir, "2C配颜师（嘉兴）生物科技有限公司.xls"), targe_dir, "2C配颜师（嘉兴）生物科技有限公司.xlsx")
    convert_purchase("{}\{}".format(source_dir, "2C配颜师（嘉兴）生物科技有限公司_采购订单.xls"), targe_dir, "2C配颜师（嘉兴）生物科技有限公司_采购订单.xlsx")
    convert_sales("{}\{}".format(source_dir, "2C配颜师（深圳）化妆品有限公司.xls"), targe_dir, "2C配颜师（深圳）化妆品有限公司.xlsx")
    convert_purchase("{}\{}".format(source_dir, "2C配颜师（深圳）化妆品有限公司_采购订单.xls"), targe_dir, "2C配颜师（深圳）化妆品有限公司_采购订单.xlsx")
    convert_sales("{}\{}".format(source_dir, "2C铲喜官（深圳）日用品有限公司.xls"), targe_dir, "2C铲喜官（深圳）日用品有限公司.xlsx")
    convert_purchase("{}\{}".format(source_dir, "2C铲喜官（深圳）日用品有限公司_采购订单.xls"), targe_dir, "2C铲喜官（深圳）日用品有限公司_采购订单.xlsx")
    convert_sales("{}\{}".format(source_dir, "2C魔妆（深圳）化妆品有限公司.xls"), targe_dir, "2C魔妆（深圳）化妆品有限公司.xlsx")
    convert_purchase("{}\{}".format(source_dir, "2C魔妆（深圳）化妆品有限公司_采购订单.xls"), targe_dir, "2C魔妆（深圳）化妆品有限公司_采购订单.xlsx")

    print("finish")
