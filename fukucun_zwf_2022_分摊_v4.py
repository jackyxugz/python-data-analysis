# coding=utf-8
import shutil
import sys
import os
import pandas as pd
import numpy  as np
import time
import os.path
import xlrd
import xlwt
import math
import zipfile
import tabulate

'''
自动分摊及调账
（一）、自动分摊
1、删除2020年销售出库>0的条码
2、2020年销售出库<0的商品摊销到2021年相同平台+店铺+条码的记录上
3、2021年销售出库<0的商品摊销到2021年相同平台+店铺+条码的记录上
4、金额<=2的替换为1111111111115
5、11111打头的记录（一版是纸巾），如果相同平台+店铺+日期的正常商品销售有记录，则将该纸巾删除，金额分摊到其他相关记录上
6、11111打头的记录（一版是纸巾），如果没有相同平台+店铺+日期的正常商品销售相关记录，则保留该纸巾记录，并替换为销量前31的商品条码（如果销售的sku不足31条，则重复追加凑够31条记录）
7、如果替换纸巾发现由于金额太小，导致数量==0，则将该条码替换为最便宜的商品，数量和价格做相应调整
8、如果数量仍然==0，则删除该条码
（二）、调账
IT金额>财务，请减少金额(a)
1、删除行，累计减少b
2、选择行，扣减金额(a-b)
2、数量=int(总金额/单价)   总金额=原总金额-(a-b)
3、单价=总金额/数量   
IT金额<财务，请增加金额 a
1、选择行，增加金额a
2、数量=int(总金额/单价)    总金额=原总金额+a
2、单价=总金额/数量 

'''

debug = False

# 财务核对文件
fn_file = r"C:\Users\ns2033\Downloads\2021年订单回款-明细表（20220610154403-2021稿）（发出商品20220614093726-2022稿）-2022.6.16提供核对.xlsx"
# fn_file=r"data/odoo\2021年订单回款-明细表（20220610154403-2021稿）（发出商品20220614093726-2022稿）-2022.6.16提供核对.xlsx"

# 工作文件夹
work_dir_2C = r"C:\Users\ns2033\Downloads\摊销\需处理的\2C"
work_dir_2C发出商品 = r"C:\Users\ns2033\Downloads\摊销\需处理的\2C发出商品"

# 输出目录
# 转格式
output_dir_2C = r"C:\Users\ns2033\Downloads\摊销\自动处理结果\2C"
output_dir_2C发出商品 = r"C:\Users\ns2033\Downloads\摊销\自动处理结果\2C发出商品"

# 与财务汇总表的对比结果存放目录
output_dir = r"C:\Users\ns2033\Downloads"
# output_dir="data/odoo/处理完成"

# 定义参考文件路径
shop_file = r"Z:\it审计处理需求\odoo导入\2022年\22年店铺.xls"
sku_file = r"Z:\it审计处理需求\odoo导入\2022年\product.template.xls"
warehouse_file = r"Z:\it审计处理需求\odoo导入\2022年\22年采购仓库.xls"
tax_file = r"Z:\it审计处理需求\odoo导入\2022年\22年税率.xlsx"
super_sku_file = r"Z:\it审计处理需求\odoo导入\2022年\可用于替换的条码.xlsx"

file_columns_list = []


def refund_amount_zwf(df1, to_file):
    # 老版本算法
    print("处理退货！")
    shijishoukuan = df1["实际收款"].sum()
    # print("处理退货后：", shop)
    print(shijishoukuan)

    other_kouchu = 0

    # 价格小于等于0的记录 单独 拷贝出来
    # df_fushu = df1[((df1["平均价格"] < 0) | (df1["实际卖出"] < 0))].copy()
    # 直接用标记的信息,退货需要分摊
    df_fushu = df1[((df1["分类备注"].str.contains("做退货") | (df1["实际收款"] < 0) | (df1["实际卖出"] < 0)))]
    print("所有需要分摊的负数")
    print(df_fushu.head(10).to_markdown())

    if debug:
        df_fushu.to_excel(to_file.replace(".xlsx", "_退货记录.xlsx"), index=False)

    if df_fushu.shape[0] > 0:
        print("退货明细")
        print(df_fushu.head(10).to_markdown())

    df_fushu = df_fushu.groupby(["平台", "店铺", "商家编码"]).agg(实际收款=("实际收款", "sum"), 发货时间=("发货时间", "max")).reset_index()
    print("退货统计 核对退货总金额：", df_fushu["实际收款"].sum())
    print(df_fushu.to_markdown())
    if debug:
        df_fushu.to_excel(to_file.replace(".xlsx", "_退货汇总.xlsx"), index=False)

    print("追查数据:")
    print(df1.head(10).to_markdown())
    # 删除退货,排除掉需要分摊的纪录
    # df1 = df1[df1["平均价格"] > 0]
    # df1 = df1[df1["实际卖出"] > 0]

    # df1=df1[~(df1["金额需做分摊"].str.contains("分摊")  & (df1["实际收款"]<0)) ]
    df1 = df1[~(df1["分类备注"].str.contains("分摊") & (df1["实际收款"] < 0))]
    # df1=df1[~(df1["分类备注"].str.contains("分摊")  ) ]

    print("删除退货后:")
    print(df1.head(10).to_markdown())
    print("需要分摊的负数的总金额：", df_fushu["实际收款"].sum(), df_fushu.shape[0])
    print("删除退货后的总金额：", df1["实际收款"].sum(), df1.shape[0])
    print("删除退货后的总金额(加总)：", df1["实际收款"].sum() + df_fushu["实际收款"].sum())

    df1 = df1.reset_index()
    df1["iid"] = df1.index
    df1["总金额"] = df1["总金额"].astype("float64")
    df1["实际收款"] = df1["实际收款"].astype("float64")
    print(df1.head(5).to_markdown())
    # 摊销到前面日期最大销售额的
    del_rows = []
    deduct_rows = []
    i = 0
    for index1, row1 in df_fushu.iterrows():
        platform = row1["平台"]
        shop = row1["店铺"]
        sku = row1["商家编码"]
        fushu_amnt = abs(-row1["实际收款"])
        sum_amnt = 0
        # _del_rows = []
        _deduct_rows = []
        print("处理退货1:", platform, shop, sku, fushu_amnt)
        print(df1.groupby(["平台", "店铺"]).agg({"实际收款": np.sum}).head(5).to_markdown())
        print(df1[((df1["平台"] == platform) & (df1["店铺"].str.upper().str.contains(shop.upper())))].groupby(
            ["平台", "店铺", "商家编码"]).agg({"实际收款": np.sum}).reset_index().head(5).to_markdown())
        for index2, row2 in df1[
            ((df1["平台"] == platform) & (df1["店铺"].str.upper() == shop.upper()) & (df1["商家编码"] == sku) & (
                    df1["实际收款"] > 0))].sort_values(
            "发货时间", ascending=False).iterrows():
            # 按日期倒序摊销金额
            print("i=", i)
            # print(df1.loc[index2].to_markdown())
            if sum_amnt + abs(row2["实际收款"]) <= abs(fushu_amnt):
                sum_amnt = sum_amnt + abs(-row2["实际收款"])
                # 删除当前行
                del_rows.append(row2["iid"])
                # print("删除行:", row2["iid"], row2["总金额"])
            else:
                print("pre_sum_amnt=", sum_amnt)
                print("要扣减的当前行:", index2, row2["实际收款"])
                deduct_value = abs(abs(fushu_amnt) - abs(sum_amnt))
                sum_amnt = abs(sum_amnt) + abs(-fushu_amnt)
                print("added_sum_amnt=", sum_amnt)

                _deduct_rows.append([row2["iid"], deduct_value])
                # _deduct_rows.append(list2)
                break
            print("sum_amnt=", sum_amnt)
            i = i + 1

        # 如果不够扣减，请人工扣减其他sku
        if sum_amnt < fushu_amnt:
            debug_file = to_file.replace(".xlsx", "_需要更换条码消除退货_扣金额_{}.xlsx".format(fushu_amnt - sum_amnt))
            print("如果不够扣减，请人工扣减其他sku:", debug_file)

            if "df_rengong_kouchu" in vars():
                df_rengong_kouchu = df_rengong_kouchu.append(df_fushu.iloc[index1])
            else:
                df_rengong_kouchu = df_fushu.iloc[index1]

            other_kouchu = fushu_amnt - sum_amnt

            print("to_file:", to_file)
            _deduct_rows.clear()
        else:
            # deduct_rows.append(_deduct_rows)
            deduct_rows.extend(_deduct_rows)

        print("处理退货2:")
        print("删除:", del_rows)
        print("扣减:", deduct_rows)

    sum_tuihuo = 0
    print("要删除的记录,删除前:", df1.shape[0])
    print(del_rows)

    df_del = pd.DataFrame(del_rows, columns=["iid"])
    df_del["delete"] = 1
    print(df_del)

    df1 = df1.merge(df_del, how="left", on=["iid"])

    print("要删除的记录,删除后:", df1.shape[0])

    print("要扣减的记录")
    print(deduct_rows)
    df_deducat = pd.DataFrame(deduct_rows, columns=["iid", "deduct"]).reset_index()
    # , columns = ['行号', '扣减金额']
    print(df_deducat.to_markdown())

    df1 = df1.merge(df_deducat[["iid", "deduct"]], how="left", on=["iid"])
    print(df1[~df1["deduct"].isnull()].to_markdown())

    if debug:
        df1[~df1["deduct"].isnull()].to_excel(to_file.replace(".xlsx", "_检查扣减.xlsx"), index=False)

    print("核对退货总金额:", sum_tuihuo, df1[~df1["deduct"].isnull()]["deduct"].sum())

    # df1[~df1["deduct"].isnull()].to_excel(to_file.replace(".xlsx", "_检查扣减_结束.xlsx"), index=False)
    if debug:
        df1.to_excel(to_file.replace(".xlsx", "_处理退货_debug.xlsx"), index=False)

    # 删除要删除掉的
    print("删除前总金额:", df1["实际收款"].sum())
    df1 = df1[df1["delete"].isnull()]
    print("删除后总金额:", df1["实际收款"].sum())

    # print("追查零价格_debug1")
    # print(df1[df1["平均价格"] == 0].to_markdown())

    print("实际收款为0的记录11:", df1[df1["实际收款"] == 0].shape[0])
    print(df1[df1["实际收款"] == 0].head(5).to_markdown())

    # 扣减后重新计算  ?是否生效了？
    df1["deduct"].fillna(0)
    # df1["总金额"] = df1.apply(lambda x: x["总金额"] - x["deduct"]  if  x["deduct"]>0 else x["总金额"], axis=1)

    # 扣除后金额为0，怎么办？
    df1["实际收款"] = df1.apply(lambda x: x["实际收款"] - x["deduct"] if x["deduct"] > 0 else x["实际收款"], axis=1)
    df1["总金额"] = df1["实际收款"]
    df1["实际卖出"] = df1.apply(lambda x: int(x["实际收款"] / x["平均价格"]) if x["deduct"] > 0 else x["实际卖出"], axis=1)
    df1["平均价格"] = df1.apply(
        lambda x: x["平均价格"] if x["deduct"] == 0 else x["实际收款"] / x["实际卖出"] if x["实际卖出"] > 0 else x["平均价格"], axis=1)
    # df1[df1["deduct"] > 0]["平均价格"] = df1.apply(lambda x: x["实际收款"] / x["实际卖出"], axis=1)

    print("实际收款为0的记录12:", df1[df1["实际收款"] == 0].shape[0])
    print(df1[df1["实际收款"] == 0].head(5).to_markdown())

    # 再次扣除需要人工处理的金额
    sel_id = df1[df1["实际收款"] > other_kouchu * 2]["iid"].iloc[0]
    if sel_id != None:
        df1["实际收款"] = df1.apply(lambda x: x["实际收款"] - other_kouchu if x["iid"] == sel_id else x["实际收款"], axis=1)
        df1["总金额"] = df1["实际收款"]
        df1["实际卖出"] = df1.apply(lambda x: int(x["实际收款"] / x["平均价格"]) if x["deduct"] > 0 else x["实际卖出"], axis=1)
        df1["平均价格"] = df1.apply(
            lambda x: x["平均价格"] if x["deduct"] == 0 else x["实际收款"] / x["实际卖出"] if x["实际卖出"] > 0 else x["平均价格"], axis=1)
    else:
        print("实在没有找到可替换的数据！")
        if "df_rengong_kouchu" in vars():
            # df_fushu.iloc[index1].to_excel(debug_file, index=False)
            df_rengong_kouchu.to_excel(debug_file, index=False)

    # print("追查零价格_debug2")
    # print(df1[df1["平均价格"] == 0].to_markdown())

    print("扣减后总金额:", df1["实际收款"].sum())

    del df1["deduct"]
    del df1["iid"]
    del df1["delete"]

    return df1


def refund_amount_new(df, to_file):
    print("按百分比，撒胡椒面处理  退货 ！index_plat_shop_month")
    print(df.head(3).to_markdown())

    # df1["月份"] = df1["发货时间"].apply(lambda x: x.replace("/", "-").split('-')[1])
    # df1["月份"] = df1["发货时间"].apply(lambda x:  x.strftime("%Y-%m-%d").replace("/", "-").split('-')[1])
    # df1["月份"] = df1["发货时间"].apply(lambda x:  x.strftime("%m"))
    # df["发货时间"]=df["发货时间"].apply(lambda x:x.strftime("%Y-%m-%d") )
    shijishoukuan = df["实际收款"].sum()
    # print("处理退货后：", shop)

    # df=df.reset_index()
    df["id"] = df.index

    # 如果发货时间和回款日期都是空

    df["发货时间"] = df.apply(
        lambda x: "2022-01-01" if ((x["发货时间"] != x["发货时间"]) & (x["回款日期"] != x["回款日期"])) else x["发货时间"], axis=1)
    df["发货时间"] = df.apply(lambda x: "2022-01-31" if str(x["发货时间"]).find("2021") >= 0 else x["发货时间"], axis=1)
    df["发货时间"].fillna("2022-01-01", inplace=True)
    # df["发货时间"] = df["发货时间"].apply(lambda x: x.strftime("%Y-%m-%d"))
    df["发货时间"] = df["发货时间"].astype(str)

    # df["月份"] = df.apply(lambda x: x["发货时间"].strftime("%m") if x["发货时间"].strftime("%y") == 2022 else '01', axis=1)
    # df["月份"] = df.apply(lambda x: "01" if x["发货时间"].find("2021")>=0 else x["发货时间"].split("-")[1], axis=1)

    df.to_excel("d:\kkktesss.xlsx")

    df["月份"] = df.apply(lambda x: x["发货时间"].replace("/", "-").split("-")[1], axis=1)

    print("测试测试")
    # print(df.head(10).to_markdown())
    # df1["月份"] = df1["发货时间"].apply(lambda x: x.strftime("%m"))

    if df.shape[0] > 0:
        df["index_plat_shop"] = df.apply(lambda x: "{}{}".format(x["平台"], x["店铺"]), axis=1)
        df["index_plat_shop_paydate"] = df.apply(lambda x: "{}{}{}".format(x["平台"], x["店铺"], x["发货时间"]), axis=1)
        df["index_plat_shop_recvdate"] = df.apply(lambda x: "{}{}{}".format(x["平台"], x["店铺"], x["回款日期"]), axis=1)
        df["index_plat_shop_month"] = df.apply(lambda x: "{}{}{}".format(x["平台"], x["店铺"], x["月份"]), axis=1)
    else:
        df["index_plat_shop"] = ""
        df["index_plat_shop_paydate"] = ""
        df["index_plat_shop_recvdate"] = ""
        df["index_plat_shop_month"] = ""

    # df = df[df["店铺名称"].str.contains(shop)]

    print("抽查数据")
    print(df.head(3).to_markdown())

    print("实际收款为0的记录1:", df[df["实际收款"] == 0].shape[0])
    print(df[df["实际收款"] == 0].head(5).to_markdown())

    print("商家编码为空...1")
    print(df[df["商家编码"].str.contains("nan")].to_markdown())
    other_kouchu = 0

    print("摊销负数之前金额:", shijishoukuan)

    # 价格小于等于0的记录 单独 拷贝出来
    # df_fushu = df1[((df1["平均价格"] < 0) | (df1["实际卖出"] < 0))].copy()
    # 直接用标记的信息,退货需要分摊
    # df_fushu = df1[  ( df1["金额需做分摊"].str.contains("分摊") &(df1["实际收款"]<0) )   ]
    # df_fushu = df1[  (  ( df1["分类备注"].str.contains("分摊") |  df1["分类备注"].str.contains("做退货"))   &(df1["实际收款"]<0) )   ]
    tag = "1111111111"

    # 需要摊销负数的索引
    # df1["分类备注"].str.contains("分摊") |
    index_need_refund = df["分类备注"].str.contains("做退货") | (df["实际收款"] < 0) | (df["实际卖出"] < 0)
    df_fushu = df[index_need_refund]

    # 需要分摊的正数的索引
    index_need_tanxiao = df["商家编码"].str.contains(tag) | df["分类备注"].str.contains("分摊")
    df_fentan = df[index_need_tanxiao]
    df_fentan = df_fentan[~df_fentan.id.isin(df_fushu.id)]  # 避免有交集

    # df_fushu = df1[((df1["分类备注"].str.contains("做退货") | (df1["实际收款"] < 0) | (df1["实际卖出"] < 0)))]
    # 摊销保留到下一个环节去执行
    # df1 = df[ (~index_need_refund) & (~index_need_tanxiao)]

    # 正常的数据,nomal
    df1 = df[~df.id.isin(df_fushu.id)]
    df1 = df1[~df.id.isin(df_fentan.id)]

    # df1 = df1[~ (index_need_refund )]
    print("金额 {} 分解：".format(df["实际收款"].sum()))
    print("正常金额：", df1["实际收款"].sum())
    print("负数金额：", df_fushu["实际收款"].sum())
    print("纸巾摊销金额：", df_fentan["实际收款"].sum())

    # 当月除了纸巾就是负数，属于新店启动，刷单，没有任何正产销售的年月
    df_newshop = df[~df["index_plat_shop_month"].isin(df1["index_plat_shop_month"])]

    # 排除掉新店
    df_fushu = df_fushu[~df_fushu.index_plat_shop_month.isin(df_newshop.index_plat_shop_month)]
    df_fentan = df_fentan[~df_fentan.index_plat_shop_month.isin(df_newshop.index_plat_shop_month)]

    print(df_fushu.head(5).to_markdown())
    if df_fushu.shape[0] > 0:
        print("发现负数了,需要做退货!")
        # df_fushu["index_plat_shop_month"] = df_fushu.apply(lambda x: "{}{}{}".format(x["平台"], x["店铺"], x["月份"]), axis=1)

        # 剩下是正数，需要分摊的部分，也不要进入正常数据集
        # print("发现奇葩数据：")
        # print(df1[df1["回款日期"].str.contains("2022-05-11,2022-05-09") & df1["平台"].str.contains("百度") ].to_markdown())

        # df1["index_plat_shop_month"] = df1.apply(lambda x: "{}{}{}".format(x["平台"], x["店铺"], x["月份"]), axis=1)

        # df_fushu = df1[  ( df1["分类备注"].str.contains("分摊")   )   ]
        print("所有需要分摊的负数:", df_fushu["实际收款"].sum())
        print(df_fushu.head(10).to_markdown())

        if debug:
            df_fushu.to_excel(to_file.replace(".xlsx", "_退货记录.xlsx"), index=False)

        if df_fushu.shape[0] > 0:
            print("退货明细")
            print(df_fushu.head(10).to_markdown())

        # 算出每个月需要摊销的负金额
        df_fushu_group = df_fushu.groupby(["index_plat_shop_month"]).agg(实际收款=("实际收款", "sum")).reset_index()
        df_fushu_group.columns = ["index_plat_shop_month", "摊销负数金额"]
        print("退货统计 核对退货总金额：", df_fushu_group["摊销负数金额"].sum())
        print(df_fushu_group.head(10).to_markdown())
        if debug:
            df_fushu_group.to_excel(to_file.replace(".xlsx", "_退货汇总.xlsx"), index=False)

        # 如果当月只有负数怎么办？
        df_fushu_group["月份"] = df_fushu_group["index_plat_shop_month"].apply(lambda x: x[-2:])

        # df_yiwai= df_fushu_group[~df_fushu_group["index_plat_shop_month"].isin(df1["index_plat_shop_month"])]
        # # df_yiwai["月份"]=df_yiwai["index_plat_shop_month"].apply(lambda x: x[-2:])
        # if df_yiwai.shape[0]>0:
        #     # max_month= df1[df1["index_plat_shop_month"].isin(df_fushu_group["index_plat_shop_month"])]["月份"].agg("max").reset_index().iloc[0,0]
        #     max_month= df1[df1["index_plat_shop_month"].isin(df_fushu_group["index_plat_shop_month"])     ]["月份"].agg("min")  # ?????
        #
        #     # df_yiwai[["月份"]].iloc[0,0]
        #     print("太意外了！如果当月只有负数怎么办?累计到最早一个月份!",max_month) # ,df1["店铺名称"],df1["销售分类"]
        #     print("需要换条码！")
        #     print(df_yiwai.head(10).to_markdown())
        #
        #     df_fushu_group=df_fushu_group.merge(df_yiwai,how="left",on=["index_plat_shop_month"])
        #     df_fushu_group["index_plat_shop_month"]=df_fushu_group.apply(lambda x:  "{}{}".format(x["index_plat_shop_month"][:-2],str(max_month)) if  str(x["月份_y"]) !='nan'  else x["index_plat_shop_month"] ,axis=1 )
        #
        #     del df_fushu_group["月份_x"]
        #     del df_fushu_group["月份_y"]
        #     del df_fushu_group["摊销负数金额_y"]
        #     df_fushu_group.rename(columns={"摊销负数金额_x":"摊销负数金额"},inplace=True)
        #     df_fushu_group=df_fushu_group.groupby(["index_plat_shop_month"])["摊销负数金额"].agg("sum")
        #     print(df_fushu_group.to_markdown())

        print("查看要摊销的负数金额")
        print(df_fushu_group.head(12).to_markdown())

        # df1["正常销售金额"]=df1.groupby(["平台", "店铺", "月份"])["实际收款"].agg("cumsum")
        df1["正常销售金额"] = df1.groupby(by=["index_plat_shop_month"])["实际收款"].transform("sum")
        df1 = df1.merge(df_fushu_group, how="left", on=["index_plat_shop_month"])

        print("商家编码为空...")
        print(df1[df1["商家编码"].str.contains("nan")].to_markdown())

        df1.to_excel("d:\c数摊销前结果.xlsx")

        df1["摊销负数金额"].fillna(0, inplace=True)  # 默认是负数
        df1["摊销比例"] = df1.apply(lambda x: x["摊销负数金额"] / x["正常销售金额"] if x["摊销负数金额"] < 0 else 0, axis=1)
        df1["实际收款"] = df1.apply(lambda x: x["实际收款"] + x["实际收款"] * x["摊销比例"], axis=1)

        df1["总金额"] = df1["实际收款"]
        df1["实际卖出"] = df1.apply(lambda x: math.ceil(x["实际收款"] / x["平均价格"]), axis=1)  # 这里报错
        df1["平均价格"] = df1.apply(lambda x: x["实际收款"] / x["实际卖出"], axis=1)

        # df1.to_excel("d:\负数摊销后结果.xlsx")

        print("查看摊销后的效果")
        print(df1.head(20).to_markdown())

        print("分摊前 {}，分摊后 {} ,另外需要摊销的正数有:{}".format(shijishoukuan, df1["实际收款"].sum(), df_fentan["实际收款"].sum()))

        df1.to_excel(r"d:\撒胡椒面摊销以后.xlsx")

        # print("删除退货后:")
        # print(df1.head(10).to_markdown())
        # print("需要分摊的负数的总金额：", df_fushu["实际收款"].sum(), df_fushu.shape[0])
        # print("删除退货后的总金额：", df1["实际收款"].sum(), df1.shape[0])
        # print("删除退货后的总金额(加总)：", df1["实际收款"].sum() + df_fushu["实际收款"].sum())

        print("扣减后总金额:", df1["实际收款"].sum())
        print("纸巾摊销金额：", df_fentan["实际收款"].sum())

        del df1["摊销负数金额"]
        del df1["正常销售金额"]
        del df1["摊销比例"]
        # del df1["index_plat_shop_month"]
    else:
        print("没有发现退货！")

    # del df1["月份"]
    default_code = "066005"
    print("没有替换条码之前金额:", df1["实际收款"].sum())
    if df_newshop.shape[0] > 0:
        df1 = change_sku(df1, df_newshop, default_code)
    print("替换条码以后的金额:", df1["实际收款"].sum())

    return df1, df_fentan


def deduct_amount(df1, to_file, platform, shop, sku, amnt):
    print("比财务的多，要扣减")
    # print("检查 level_0")
    # print(df1.head(5).to_markdown())
    df1["商家编码"] = df1["商家编码"].astype(str)
    if df1.columns.__contains__("level_0"):
        del df1["level_0"]

    df1 = df1.reset_index()
    df1["iid"] = df1.index
    # print(df1.head(50).to_markdown())
    # 摊销到前面日期最大销售额的
    del_rows = []
    deduct_rows = []
    i = 0
    sum_amnt = 0
    # _del_rows = []
    _deduct_rows = []
    print("处理退货1111:", platform, shop, sku, amnt)
    print("符合扣除记录的列表:")
    # 按金额扣减
    df1["商家编码"] = df1["商家编码"].astype(str)
    df1["总金额"] = df1["总金额"].astype("float64")
    # koujian_db = df1[df1["平台"] == platform]
    # koujian_db = koujian_db[koujian_db["店铺"] == shop]
    # koujian_db = koujian_db[koujian_db["商家编码"].str.contains(str(sku))]
    # koujian_db = koujian_db[koujian_db["总金额"] >0 ]
    koujian_db = df1[((df1["平台"] == platform) & (df1["店铺"] == shop) & (df1["商家编码"].str.contains(str(sku))) & (
                df1["总金额"] > 0))].sort_values("总金额", ascending=False)
    # koujian_db=koujian_db.sort_values("总金额", ascending=False)

    if koujian_db.shape[0] > 0:
        print(koujian_db.head(10).to_markdown())
        for index2, row2 in koujian_db.sort_values(
                "总金额", ascending=False).iterrows():
            # 按日期倒序摊销金额
            print("循环,", index2)
            print("i=", i)
            # print(df1.loc[index2].to_markdown())
            if sum_amnt + abs(row2["总金额"]) < abs(amnt):
                sum_amnt = sum_amnt + abs(row2["总金额"])
                # df1.iloc[df1.id==row2["id"],""]
                # 删除当前行
                # df1=df1[df1.id!=index2]
                # del_rows.append(index2)
                del_rows.append(row2["iid"])
                # print("删除行:", row2["iid"], row2["总金额"])
            else:
                print("pre_sum_amnt=", sum_amnt)
                print("要扣减的当前行:", index2, "目前金额:", row2["总金额"])
                deduct_value = abs(abs(amnt) - abs(sum_amnt))
                sum_amnt = abs(sum_amnt) + abs(-amnt)
                print("added_sum_amnt=", sum_amnt)
                # df1.iloc[df1.id == row2["id"], "总回款"]=df1["总回款"]-(amnt-sum_amnt)
                # print("当前行:", row2["index"], abs(-row2["总金额"]), " 累计:", abs(sum_amnt) + abs(-amnt), " 更新金额为:",
                #       -(deduct_value))

                # print("当前行:", row2["iid"], abs(-row2["总金额"]), " 累计:", abs(sum_amnt) + abs(-amnt), " 更新金额为:",
                #       -(deduct_value))

                _deduct_rows.append([row2["iid"], deduct_value])
                # _deduct_rows.append(list2)
                break
            print("sum_amnt=", sum_amnt)
            i = i + 1

        # 如果不够扣减，请人工扣减其他sku
        if sum_amnt < amnt:
            print("不够扣减", sum_amnt, ",", amnt)
            # df_fushu.to_markdown(to_file.replace(".xlsx", "_需要更换条码消除退货_扣金额_{}_2.xlsx".format(amnt - sum_amnt)), index=False)
            # df_fushu.to_markdown(to_file.replace(".xlsx", "_需要更换条码消除退货_扣金额_2.xlsx"), index=False)
            _deduct_rows.clear()
        else:
            # deduct_rows.append(_deduct_rows)
            deduct_rows.extend(_deduct_rows)
    else:
        print("没有找到符合扣减标准的任何纪录！", platform, shop, sku)
        # print(df1[
        #           ((df1["平台"] == platform) & (df1["店铺"] == shop) & (df1["商家编码"] == sku)  )].sort_values(
        #     "总金额", ascending=False).head(100).to_markdown())
        # print(df1[
        # ((df1["平台"] == platform) & (df1["店铺"] == shop)   )].sort_values(
        # "总金额", ascending=False).head(100).to_markdown())

        # print(df1.groupby(["平台","店铺","商家编码"]).agg({"总金额":np.max,"总金额":np.sum}).reset_index().to_markdown())
        print(
            df1.groupby(["平台", "店铺", "商家编码"]).agg(最大总金额=("总金额", "max"), 合计总金额=("总金额", "sum")).reset_index().sort_values(
                "合计总金额", ascending=False).to_markdown())

    print("处理退货2:")
    print("删除:", del_rows)
    print("扣减:", deduct_rows)

    sum_tuihuo = 0
    print("要删除的记录,删除前:", df1.shape[0])
    print(del_rows)

    df_del = pd.DataFrame(del_rows, columns=["iid"])
    df_del["delete"] = 1
    print(df_del)

    df1 = df1.merge(df_del, how="left", on=["iid"])

    print("要删除的记录,删除后2:", df1.shape[0])
    # print("要扣减的记录")
    print(deduct_rows)
    df_deducat = pd.DataFrame(deduct_rows, columns=["iid", "deduct"]).reset_index()
    # , columns = ['行号', '扣减金额']
    print(df_deducat.to_markdown())

    df1 = df1.merge(df_deducat[["iid", "deduct"]], how="left", on=["iid"])
    print(df1[~df1["deduct"].isnull()].to_markdown())

    if debug:
        df1[~df1["deduct"].isnull()].to_excel(to_file.replace(".xlsx", "_检查扣减2.xlsx"), index=False)

    print("核对退货总金额:", sum_tuihuo, df1[~df1["deduct"].isnull()]["deduct"].sum())

    # df1[~df1["deduct"].isnull()].to_excel(to_file.replace(".xlsx", "_检查扣减_结束.xlsx"), index=False)
    if debug:
        df1.to_excel(to_file.replace(".xlsx", "_处理退货_debug2.xlsx"), index=False)

    # 删除要删除掉的
    print("删除前总金额:", df1["总金额"].sum())
    df1 = df1[df1["delete"].isnull()]
    print("删除后总金额:", df1["总金额"].sum())

    # 扣减后重新计算  ?是否生效了？
    df1["deduct"].fillna(0)
    # df1["总金额"] = df1.apply(lambda x: x["总金额"] - x["deduct"]  if  x["deduct"]>0 else x["总金额"], axis=1)
    df1["实际收款"] = df1.apply(lambda x: x["实际收款"] - x["deduct"] if x["deduct"] > 0 else x["实际收款"], axis=1)
    df1["总金额"] = df1["实际收款"]
    df1["实际卖出"] = df1.apply(lambda x: int(x["实际收款"] / x["平均价格"]) if x["deduct"] > 0 else x["实际卖出"], axis=1)
    df1["平均价格"] = df1.apply(lambda x: x["平均价格"] if x["deduct"] == 0 else x["实际收款"] / x["实际卖出"] if x["实际卖出"] > 0 else 0,
                            axis=1)
    # df1[df1["deduct"] > 0]["平均价格"] = df1.apply(lambda x: x["实际收款"] / x["实际卖出"], axis=1)

    print("扣减后总金额:", df1["总金额"].sum())

    del df1["deduct"]
    del df1["iid"]
    del df1["delete"]

    return df1


def add_amount(df1, platform, shop, amnt):
    print("比财务少,要增加金额")
    # print("检查 level_0")
    # print(df1.head(5).to_markdown())
    if df1.columns.__contains__("level_0"):
        del df1["level_0"]

    df1["商家编码"] = df1["商家编码"].astype(str)
    df1 = df1.reset_index()
    df1["iid"] = df1.index

    df1["总金额"] = df1["总金额"].astype("float64")

    # print(df1.head(50).to_markdown())
    # 摊销到前面日期最大销售额的
    add_rows = []
    print("处理增加1111:", platform, shop, amnt)
    print("符合增加记录的列表:")
    # 增加金额
    print(df1[
              ((df1["平台"] == platform) & (df1["店铺"] == shop))].sort_values(
        "总金额", ascending=False).head(10).to_markdown())

    df1["总金额"] = df1["总金额"].astype("float64")
    # 找到销量最大的产品，然后找到最后的销售日期
    print("寻找销量最大的产品")

    print("看所有")
    print(df1.groupby(["平台", "店铺"]).agg({"总金额": np.sum}).to_markdown())

    # print(df1.sort_values(
    #     "总金额", ascending=False).head(10).to_markdown())
    #
    # print("看平台")
    # print(df1[
    #           ((df1["平台"].str.contains(platform)) )].sort_values(
    #     "总金额", ascending=False).head(10).to_markdown())
    #
    # print("看店铺")
    # print(df1[
    #     ((df1["平台"].str.contains(platform)) & (df1["店铺"].str.upper().str.contains(shop.upper())  )  )].sort_values(
    #     "总金额", ascending=False).head(10).to_markdown())

    sku = iid = df1[
        ((df1["平台"].str.contains(platform)) & (df1["店铺"].str.upper().str.contains(shop.upper())))].sort_values(
        "总金额", ascending=False).iloc[0]["商家编码"]

    # iid=df1[
    #     ((df1["平台"] == platform) & (df1["店铺"] == shop) & (df1["商家编码"].str.contains(sku)))].sort_values(
    #     "总金额", ascending=False).iloc[0]["iid"]

    iid = df1[
        ((df1["平台"] == platform) & (df1["店铺"].str.upper() == shop.upper()) & (
            df1["商家编码"].str.contains(sku)))].sort_values(
        "发货时间", ascending=False).iloc[0]["iid"]

    add_rows.append([iid, amnt])

    print("处理追加:", add_rows)
    # 删除要删除掉的
    print("增加前总金额:", df1["总金额"].sum())
    # print("增加后后总金额:", df1["总金额"].sum())
    print(df1[df1["iid"] == iid].to_markdown())
    # df1[df1["iid"] == iid].to_excel(r"d:\test_debug.xlsx")

    # 扣减后重新计算  ?是否生效了？
    df1["总金额"] = df1.apply(lambda x: x["总金额"] + amnt if x["iid"] == iid else x["总金额"], axis=1)
    df1["实际收款"] = df1.apply(lambda x: x["实际收款"] + amnt if x["iid"] == iid else x["实际收款"], axis=1)
    df1["实际卖出"] = df1.apply(
        lambda x: 1 if x["平均价格"] == 0 else int(x["实际收款"] / x["平均价格"]) if x["iid"] == iid else x["实际卖出"], axis=1)
    # df1["平均价格"] = df1.apply(lambda x:  x["平均价格"] if x["iid"]==iid else x["实际收款"] / x["实际卖出"] if x["实际卖出"]>0 else 0  , axis=1)
    df1["平均价格"] = df1.apply(lambda x: x["实际收款"] / x["实际卖出"] if x["实际卖出"] > 0 else x["实际收款"], axis=1)

    print("增加后后总金额:", df1["总金额"].sum())

    del df1["iid"]

    return df1


def patch_advise(df, del_fn, add_fn, to_file):
    i = 0
    if len(del_fn) > 0:
        for k in del_fn:
            platform = k[0]
            shop = k[1]
            sku = k[2]
            amnt = k[3]

            print("扣减财务多余的数据:", k)
            if ((platform > '') & (shop > '') & (str(sku) > '')):
                df = deduct_amount(df, to_file, platform, shop, sku, amnt)
                # print(filename, "扣减财务（{}）后的总金额：".format("".join(k)), df["总金额"].sum(), df.shape[0])
                print("扣减财务后的总金额{}：".format(i), amnt, df["总金额"].sum(), df.shape[0])
                i = i + 1

    print("检查 level_0")
    print(df.head(5).to_markdown())
    if df.columns.__contains__("level_0"):
        del df["level_0"]

    i = 0
    if len(add_fn) > 0:
        for k in add_fn:
            platform = k[0]
            shop = k[1]
            amnt = k[2]

            print("增加不足的数据:", k)
            if ((platform > '') & (shop > '') & (sku > '')):
                df = add_amount(df, platform, shop, amnt)
                # print(filename, "扣减财务（{}）后的总金额：".format("".join(k)), df["总金额"].sum(), df.shape[0])
                print("扣减财务后的总金额{}：".format(i), amnt, df["总金额"].sum(), df.shape[0])
                i = i + 1

    # print("检查 level_0")
    # print(df.head(5).to_markdown())
    if df.columns.__contains__("level_0"):
        del df["level_0"]

    return df


def change_sku(df1, df_newshop, default_code):
    # 换品
    # df_newshop
    # change_type 类型：指定，最畅销，最低价
    # if df_fentan_2.shape[0] > 0:
    print("需要采用第二种方法,换品")

    df_super_sku = pd.read_excel(super_sku_file, dtype=str)
    print(df_super_sku.to_markdown())
    df_super_sku["内部参考"] = df_super_sku["内部参考"].astype(str)
    df_super_sku = df_super_sku[df_super_sku["内部参考"].str.contains(default_code)]
    df_super_sku.rename(columns={"产品类别": "类别", "产品条码": "商家编码", "产品名称": "商品名称", "平均售价": "平均价格"}, inplace=True)
    df_super_sku["type1"] = 1
    df_super_sku["平均价格"] = df_super_sku["平均价格"].astype("float")

    print("查看替补产品:")
    print(df_super_sku.to_markdown())

    sale_type = df_newshop["销售分类"].iloc[0]

    # 当月的汇总成一天
    df_newshop["月份"] = df_newshop["发货时间"].apply(lambda x: x.split("-")[1])

    # 按日期 汇总  1111
    # 主体	店铺	平台	品牌	类别	支付日期	回款日期	商家编码	商品名称	总回款	总退款	币种	外币回款汇总	外币退款汇总	外币转换港币回款汇总	外币转换港币退款汇总	实际收款	回款产品	退款产品	实际卖出	平均价格	总金额	未税金额	税率
    df_fentan_2 = df_newshop.groupby(by=["主体", "店铺", "平台", "月份"]).agg(回款日期=("回款日期", lambda x: ",".join(x)),
                                                                      总回款=("总回款", "sum"), 总退款=("总退款", "sum")
                                                                      , 外币回款汇总=("外币回款汇总", "sum"),
                                                                      外币退款汇总=("外币退款汇总", "sum"),
                                                                      外币转换港币回款汇总=("外币转换港币回款汇总", "sum"),
                                                                      外币转换港币退款汇总=("外币转换港币退款汇总", "sum"),
                                                                      实际收款=("实际收款", "sum"),
                                                                      实际卖出=("实际卖出", "sum"),
                                                                      总金额=("总金额", "sum"),
                                                                      未税金额=("未税金额", "sum"),
                                                                      税率=("税率", "max"),
                                                                      币种=("币种", "max"),
                                                                      发货时间=("发货时间", "max")
                                                                      ).reset_index()

    df_fentan_2["memo"] = ""
    df_fentan_2["备注"] = ""
    df_fentan_2["回款日期"] = df_fentan_2["回款日期"].apply(lambda x: ",".join(list(set(x.split(",")))))  # 回款日期去重
    print("汇总后的需要分摊数据")
    # print(df_fentan_2.head(10).to_markdown())
    # df_fentan_2.to_excel(to_file.replace(".xlsx", "_需要分摊的独孤的纸巾.xlsx"), index=False)

    # 最多31天，也就是31行
    # df_fentan_2=df_fentan_2.reset_index()
    df_fentan_2["id"] = df_fentan_2.index
    df_fentan_2["type1"] = 1
    df_fentan_2["销售分类"] = sale_type
    # 替换条码 ，分摊到前几名的常卖产品上
    # df_hot = df[~df["商家编码"].str.contains("1111")].sort_values(by=["实际卖出"], ascending=False)
    # df_hot = df[~df["商家编码"].str.contains("xxxxxxxxxx")].sort_values(by=["实际卖出"], ascending=False)

    # 指定商品来替代
    # 这里假设开新店，当月销售额不为负数
    # 如果为负数，则需要摊销到其他月份
    if True:
        print("拼接前行数:", df_fentan_2.shape[0])
        df_fentan_2 = df_fentan_2.merge(df_super_sku[["type1", "类别", "内部参考", "商家编码", "商品名称", "品牌", "平均价格"]], how="left",
                                        on=["type1"])
        print("拼接后。行数:", df_fentan_2.shape[0])

        print("立即查看换品拼接后的结果:")
        print(df_fentan_2.head(10).to_markdown())

        # df_fentan_2.to_excel("d:\debug_分摊与热卖.xlsx")
        if df_fentan_2[df_fentan_2["税率"].isnull()].shape[0] > 0:
            print("替换后的数据，寻找空税率")
            print(df_fentan_2[df_fentan_2["税率"].isnull()].head(10).to_markdown())
            # df_fentan_2.to_excel(to_file.replace(".xlsx", "_替换后的数据——抽查价格.xlsx"), index=False)

        print("替换后的数据,实际卖出不能==0")
        # df_fentan_2.to_excel(to_file.replace(".xlsx", "_这里不出错.xlsx"), index=False)
        # df_fentan_2["平均价格"] = df_fentan_2.apply(lambda x: (x["实际收款"] / x["实际卖出"]) if x["实际卖出"] > 0 else 0, axis=1)

        df_fentan_2["实际卖出"] = df_fentan_2.apply(lambda x: math.ceil(x["实际收款"] / x["平均价格"]) if x["平均价格"] > 0 else 1,
                                                axis=1)

        df_fentan_2["平均价格"] = df_fentan_2.apply(lambda x: (x["实际收款"] / x["实际卖出"]) if x["实际卖出"] > 0 else 0, axis=1)

        # df_fentan_2["实际卖出"]=df_fentan_2.apply(lambda x:  math.ceil(x["总金额"]/x["平均价格"])  if x["平均价格"]>0 else 0    ,axis=1 )
        # df_fentan_2.to_excel(to_file.replace(".xlsx", "_这里出错.xlsx"), index=False)

        # df_fentan_2.loc[df_fentan_2["实际卖出"] > 0]["memo"]="第一批替换条码"
        df_fentan_2["备注"] = df_fentan_2.apply(lambda x: "{},当月开新店替换条码".format(x["备注"]) if x["实际卖出"] > 0 else x["备注"],
                                              axis=1)

        if df_fentan_2[df_fentan_2["平均价格"] == 0].shape[0] > 0:
            print("替换后的数据,平均价格仍然为0")
            print(df_fentan_2[df_fentan_2["平均价格"] == 0].head(10).to_markdown())
            # if df_fentan_2[df_fentan_2["平均价格"]==0].shape[0]>0:
            #     df_fentan_2.to_excel(to_file.replace(".xlsx", "_替换后的数据平均价格为0.xlsx"), index=False)

        # df_fentan_2.to_excel(to_file.replace(".xlsx", "_替换后的数据——更新数量.xlsx"), index=False)
        print("第一批最低价商品替换完（纸巾换洗发水）:", df_fentan_2[df_fentan_2["实际卖出"] > 0]["实际收款"].sum())

    # 用畅销品来替代
    if False:
        df_hot = df_nomal.sort_values(by=["实际卖出"], ascending=False)
        print("正常好卖的商品有：")
        print(df_hot.head(10).to_markdown())
        # 凑到30天的记录
        i = 0
        if df_hot.shape[0] > 0:
            while df_hot.shape[0] <= df_fentan_2.shape[0]:  # 31天不够用，可能涉及多个月份
                # print("跟踪:",i)
                # print(df_hot.head(5).to_markdown())
                df_hot = df_hot.append(df_hot)

        df_hot = df_hot.reset_index()
        # df_hot = df_hot.head(31)
        df_hot["id"] = df_hot.index

        print("畅销产品排行榜！")
        print(df_hot.head(10).to_markdown())
        # print(df_hot[df_hot["税率"].isnull()].to_markdown())

        # print("凑够1个月的可用于分摊的数据")
        # print(df_hot.head(10).to_markdown())
        # print(df_fentan_2.head(10).to_markdown())
        df_hot.rename(columns={"内部参考号": "内部参考"}, inplace=True)

        # 这里进行替换条码,将需要分摊的产品替换成畅销产品的条码，品牌，类别，名称，价格都自动带出来
        df_fentan_2 = df_fentan_2.merge(df_hot[["id", "品牌", "类别", "内部参考", "商家编码", "商品名称", "平均价格"]], how="left",
                                        on=["id"])

        # df_fentan_2.to_excel("d:\debug_分摊与热卖.xlsx")

        if df_fentan_2[df_fentan_2["税率"].isnull()].shape[0] > 0:
            print("替换后的数据，寻找空税率")
            print(df_fentan_2[df_fentan_2["税率"].isnull()].head(10).to_markdown())
            # df_fentan_2.to_excel(to_file.replace(".xlsx", "_替换后的数据——抽查价格.xlsx"), index=False)

        if df_fentan_2[df_fentan_2["平均价格"] == 0].shape[0] > 0:
            print("替换后的数据,平均价格为0")
            print(df_fentan_2[df_fentan_2["平均价格"] == 0].head(10).to_markdown())
            # if df_fentan_2[df_fentan_2["平均价格"]==0].shape[0]>0:
            #     df_fentan_2.to_excel(to_file.replace(".xlsx", "_替换后的数据平均价格为0.xlsx"), index=False)

        print("替换后的数据,实际卖出不能==0")
        # df_fentan_2.to_excel(to_file.replace(".xlsx", "_这里不出错.xlsx"), index=False)
        df_fentan_2["实际卖出"] = df_fentan_2.apply(lambda x: math.ceil(x["实际收款"] / x["平均价格"]) if x["平均价格"] > 0 else 1,
                                                axis=1)
        df_fentan_2["平均价格"] = df_fentan_2.apply(lambda x: (x["实际收款"] / x["实际卖出"]) if x["实际卖出"] > 0 else 0, axis=1)
        # df_fentan_2["实际卖出"]=df_fentan_2.apply(lambda x:  math.ceil(x["总金额"]/x["平均价格"])  if x["平均价格"]>0 else 0    ,axis=1 )
        # df_fentan_2.to_excel(to_file.replace(".xlsx", "_这里出错.xlsx"), index=False)

        # df_fentan_2.loc[df_fentan_2["实际卖出"] > 0]["memo"]="第一批替换条码"
        df_fentan_2["memo"] = df_fentan_2.apply(lambda x: "第一批替换条码" if x["实际卖出"] > 0 else x["memo"], axis=1)

        # df_fentan_2.to_excel(to_file.replace(".xlsx", "_替换后的数据——更新数量.xlsx"), index=False)
        print("第一批最低价商品替换完（纸巾换洗发水）:", df_fentan_2[df_fentan_2["实际卖出"] > 0]["实际收款"].sum())
        # print(df_fentan_2[df_fentan_2["实际卖出"]> 0].head(3).to_markdown())

    # 打补丁，实际卖出==0，找到价格最低的销售商品清单进行替换
    print("第二批替换条码,实际卖出==0")
    if False:
        if df_fentan_2[df_fentan_2["实际卖出"] == 0].shape[0] > 0:
            # print(filename, "实际卖出==0 1")
            print("还需要打补丁：")
            db_too_low = df_fentan_2[df_fentan_2["实际卖出"] == 0].copy()
            print(db_too_low[
                      ["主体", "平台", "店铺", "发货时间", "总回款", "总退款", "实际收款", "总金额", "实际卖出", "平均价格", "id", "品牌", "类别", "商家编码",
                       "商品名称"]].head(10).to_markdown())
            melt_money = db_too_low["实际收款"].sum()  # 需要被摊销掉的最后一笔金额
            print("最后一批替换及删除的纸巾:", melt_money)
            # df1.to_excel(to_file.replace(".xlsx", "_error_实际卖出为0.xlsx"), index=False)

            print("价格太低，单价为0的全部摊销到其他商品上!")
            df_fentan_2 = df_fentan_2[df_fentan_2["实际卖出"] > 0]

            if df_fentan_2.shape[0] > 0:
                print("摊销到最后一笔摊销过的销售记录上")
                print(df_fentan_2.tail(1).to_markdown())
                # print(df_fentan_2.head(1).to_markdown())
                last_index = df_fentan_2.tail(1)["id"].iloc[0]
                df_fentan_2["实际收款"] = df_fentan_2.apply(
                    lambda x: x["实际收款"] + melt_money if x["id"] == last_index else x["实际收款"], axis=1)
                df_fentan_2["总金额"] = df_fentan_2["实际收款"]
                df_fentan_2["memo"] = df_fentan_2.apply(
                    lambda x: x["memo"] + ",把金额太小的最后摊销掉" if x["id"] == last_index else x["memo"], axis=1)

                df_fentan_2["实际卖出"] = df_fentan_2.apply(lambda x: int(x["实际收款"] / x["平均价格"]) if x["平均价格"] > 0 else 1,
                                                        axis=1)
                df_fentan_2["平均价格"] = df_fentan_2.apply(lambda x: (x["实际收款"] / x["实际卖出"]) if x["实际卖出"] > 0 else 0,
                                                        axis=1)

                if df_fentan_2[df_fentan_2["实际卖出"] == 0].shape[0] > 0:
                    print("岂有此理，仍然还有无法摊销的！")
            else:
                print("没有一笔能正常摊销，那就摊销到正常销售记录上")
                print(df1.tail(1).to_markdown())
                # print(df1.head(1).to_markdown())
                last_index = df1.tail(1)["id"].iloc[0]
                df1["实际收款"] = df1.apply(
                    lambda x: x["实际收款"] + melt_money if x["id"] == last_index else x["实际收款"], axis=1)
                df1["总金额"] = df1["实际收款"]
                df1["memo"] = df1.apply(
                    lambda x: x["memo"] + ",把金额太小的最后摊销掉" if x["id"] == last_index else x["memo"], axis=1)

                df1["实际卖出"] = df1.apply(
                    lambda x: int(x["实际收款"] / x["平均价格"]) if x["平均价格"] > 0 else 1,
                    axis=1)
                df1["平均价格"] = df1.apply(lambda x: (x["实际收款"] / x["实际卖出"]) if x["实际卖出"] > 0 else 0,
                                        axis=1)

                if df1[df1["实际卖出"] == 0].shape[0] > 0:
                    print("岂有此理，仍然还有无法摊销的!!!")

            # df1.to_excel("d:\debug_分摊第二步.xlsx")

            if False:
                # 找到价格最低的销售商品清单
                df_cheap = df_nomal.sort_values(by=["平均价格"])[["商家编码", "商品名称", "平均价格", "品牌", "类别"]].head(1)

                print("找到价格最低的销售商品清单")
                print(df_cheap.head(5).to_markdown())
                item_code = ''
                item_name = ''
                item_price = 0
                item_brand = ''
                item_type = ''

                if df_cheap.shape[0] > 0:
                    item_code = df_cheap.iloc[0, 0]
                    item_name = df_cheap.iloc[0, 1]
                    item_price = df_cheap.iloc[0, 2]
                    item_brand = df_cheap.iloc[0, 3]
                    item_type = df_cheap.iloc[0, 4]
                # min_price=df_cheap.iloc[0,1]
                # df1=df1.merge(df_cheap[["cheap_id","商家编码","平均价格"]])
                # df_fentan_2["min_price"] = item_price
                print("最低价格:", item_code, item_name, item_price)

                df_fentan_2["memo"] = df_fentan_2.apply(
                    lambda x: "更换最低价格商品" if ((x["实际卖出"] == 0) & (x["总金额"] > item_price)) else "价格太低" if x[
                                                                                                            "实际卖出"] == 0 else
                    x["memo"],
                    axis=1)

                print("df_fentan_2")
                print(df_fentan_2.head(10).to_markdown())

                # print(df_fentan_2[df_fentan_2["实际卖出"] == 0].head(10).to_markdown())
                # df_fentan_2.to_excel(to_file.replace(".xlsx", "_替换后的数据——最后两条记录.xlsx"), index=False)

                # df_fentan_2.loc[df_fentan_2["memo"]=="更换最低价格商品" ]["min_price"]=item_price
                df_fentan_2["min_price"] = df_fentan_2.apply(lambda x: item_price if x["memo"] == "更换最低价格商品" else "",
                                                             axis=1)
                # print("最后一批,替换及删除：")
                # print(df_fentan_2[df_fentan_2["实际卖出"] == 0].to_markdown())

                # 如果金额>最低价格，则匹配此商品
                # df1[df1.平均价格<df1.min_price]["支付日期"]=df[~df["商家编码"].str.contains("11111")].sort_values(by=["平均价格"])[["支付日期"]].head(1).iloc[0,0]
                # df1[df1.平均价格<df1.min_price]["回款日期"]=df[~df["商家编码"].str.contains("11111")].sort_values(by=["平均价格"])[["回款日期"]].head(1).iloc[0,0]
                # df_fentan_2[df_fentan_2.总金额.astype("float64") > df_fentan_2.min_price.astype("float64")]["memo"] = "更换最低价格商品"

                df_fentan_2["商家编码"] = df_fentan_2.apply(lambda x: item_code if x["memo"] == "更换最低价格商品" else x["商家编码"],
                                                        axis=1)
                df_fentan_2["商品名称"] = df_fentan_2.apply(lambda x: item_name if x["memo"] == "更换最低价格商品" else x["商品名称"],
                                                        axis=1)
                df_fentan_2["平均价格"] = df_fentan_2.apply(lambda x: item_price if x["memo"] == "更换最低价格商品" else x["平均价格"],
                                                        axis=1)
                df_fentan_2["实际卖出"] = df_fentan_2.apply(
                    lambda x: x["实际卖出"] if item_price == 0 else math.ceil(x["总金额"] / item_price) if x[
                                                                                                        "memo"] == "更换最低价格商品" else
                    x["实际卖出"], axis=1)
                df_fentan_2["平均价格"] = df_fentan_2.apply(
                    lambda x: x["平均价格"] if x["实际卖出"] == 0 else x["总金额"] / x["实际卖出"] if x["memo"] == "更换最低价格商品" else x[
                        "平均价格"], axis=1)
                df_fentan_2["品牌"] = df_fentan_2.apply(lambda x: item_brand if x["memo"] == "更换最低价格商品" else x["品牌"],
                                                      axis=1)
                df_fentan_2["类别"] = df_fentan_2.apply(lambda x: item_type if x["memo"] == "更换最低价格商品" else x["类别"],
                                                      axis=1)

                # df_fentan_2["销售分类"] = sale_type
                # df_fentan_2.to_excel(to_file.replace(".xlsx", "_更换最低价格商品_跟踪.xlsx"), index=False)

                # df_fentan_2.loc[df_fentan_2.memo.str.contains("更换最低价格商品")]["实际卖出"] = df_fentan_2.apply(lambda x: math.ceil(x["总金额"] / item_price),
                #                                                            axis=1)

                # print("查看更换完最低价格情况:")
                # print(df_fentan_2[df_fentan_2.memo == "更换最低价格商品"].to_markdown())
                # df_fentan_2.to_excel(to_file.replace(".xlsx", "_更换最低价格商品.xlsx"), index=False)

                # df_fentan_2[df_fentan_2.memo.str.contains("更换最低价格商品")]["平均价格"] = df_fentan_2.apply(lambda x: x["总金额"] / x["实际卖出"], axis=1)

                # print("查看修改价格情况:")
                # print(df_fentan_2[df_fentan_2.memo.str.contains("更换最低价格商品")].to_markdown())
                # print(df1[df1.平均价格>=df1.min_price].to_markdown())
                # print(df_fentan_2[df_fentan_2["实际卖出"] == 0].to_markdown())
                print("最后一批替换条码：", df_fentan_2[df_fentan_2.memo.str.contains("更换最低价格商品")]["总金额"].sum())
                # print(df_fentan_2[df_fentan_2["实际卖出"] == 0].to_markdown())
                print(df_fentan_2[df_fentan_2.memo.str.contains("更换最低价格商品")].to_markdown())

                # df_fentan_2[df_fentan_2.平均价格 >= df_fentan_2.min_price]["平均价格"] = df_fentan_2.apply(lambda x: x["总金额"] / x["实际卖出"], axis=1)
                # df1[df1.平均价格>=df1.min_price]["平均价格"]=df1.apply(lambda x:  x["总金额"]/x["实际卖出"] ,axis=1)

                # 如果价格高于最低价格，则删除，并摊销到之前的日期  价格太低
                print("这里容易出错！")
                if df_fentan_2[df_fentan_2.memo.str.contains("价格太低")].shape[0] > 0:
                    som_money = df_fentan_2[df_fentan_2.memo.str.contains("价格太低")]["总金额"].sum()
                    # the_date = df_fentan_2[df_fentan_2.平均价格 < df_fentan_2.min_price]["支付日期"].min()
                    the_date = df_fentan_2[df_fentan_2.memo.str.contains("价格太低")]["发货时间"].min()
                    # df1 = df_fentan_2[df_fentan_2.平均价格 >= df_fentan_2.min_price]["平均价格"]

                    # df["订单付款时间"] = df["订单付款时间"].astype('datetime64[ns]')

                    # 摊销到前一笔
                    print("提前一笔")
                    # print(df1)
                    print(df1.head(5).to_markdown())
                    # 刚好没有咋办？
                    # print(df_fentan_2[df_fentan_2.支付日期 < the_date].to_markdown())
                    # print(df1[df1.支付日期 < the_date].head(3).to_markdown())

                    # pre_max_id = df_fentan_2[df_fentan_2.支付日期 < the_date]["id"].max()
                    pre_max_id = df1[df1.发货时间 < the_date]["id"].max()
                    if pre_max_id == None:
                        pre_max_id = 0
                    # df_fentan_2.loc[df_fentan_2.id == pre_max_id, "总金额"] = df_fentan_2["总金额"] + som_money
                    # df_fentan_2.loc[df_fentan_2.id == pre_max_id, "平均价格"] = df_fentan_2.apply(lambda x: x["总金额"] / x["实际卖出"], axis=1)

                    print("这里多删除了数据，导致数据不平")
                    print("要删除:", pre_max_id, the_date, som_money)

                    # df_fentan_2.to_excel(to_file.replace(".xlsx", "_加薪前.xlsx"), index=False)
                    # df1.to_excel(to_file.replace(".xlsx", "_加薪前.xlsx"), index=False)

                    # df_fentan_2["memo"]=df_fentan_2.apply(lambda x: "给你加薪"  if x.id==pre_max_id else  x["memo"] ,axis=1 )
                    # df_fentan_2["总金额"]=df_fentan_2.apply(lambda x: x["总金额"]+som_money   if x.id==pre_max_id else  x["总金额"] ,axis=1 )
                    # df_fentan_2["平均价格"]=df_fentan_2.apply(lambda x: x["总金额"] / x["实际卖出"]   if x.id==pre_max_id else  x["平均价格"] ,axis=1 )

                    df1["memo"] = df1.apply(lambda x: "给你加薪" if x.id == pre_max_id else x["memo"], axis=1)
                    df1["总金额"] = df1.apply(lambda x: x["总金额"] + som_money if x.id == pre_max_id else x["总金额"],
                                           axis=1)
                    df1["平均价格"] = df1.apply(lambda x: x["总金额"] / x["实际卖出"] if x.id == pre_max_id else x["平均价格"],
                                            axis=1)
                    # 更新实际收款
                    df1["实际收款"] = df1.apply(lambda x: x["总金额"] if x.id == pre_max_id else x["实际收款"], axis=1)

                    # df_fentan_2.to_excel(to_file.replace(".xlsx", "_加薪后.xlsx"), index=False)
                    # df1.to_excel(to_file.replace(".xlsx", "_加薪后.xlsx"), index=False)

                    # 删除 价格过低的数据
                    df_fentan_2 = df_fentan_2[df_fentan_2["实际卖出"] > 0]

                # if df_fentan_2[df_fentan_2["实际卖出"] == 0].shape[0] > 0:
                #     print(filename, "实际卖出==0 2")
                #     print(df_fentan_2[df1["实际卖出"] == 0].head(10).to_markdown())
                #     df_fentan_2.to_excel(to_file.replace(".xlsx", "_error_还有实际卖出为0.xlsx"), index=False)
                #
                #     return

    # df_fentan_2["平均价格"]=df_fentan_2["总金额"]/df_fentan_2["实际卖出"]
    df_fentan_2["回款产品"] = df_fentan_2["实际卖出"]
    df_fentan_2["退款产品"] = 0
    df_fentan_2["总金额_ori"] = 0
    df_fentan_2["销售分类"] = sale_type
    # df_fentan_2["商家编码_ori"] = 0

    if df_fentan_2[df_fentan_2["税率"].isnull()].shape[0] > 0:
        print("替换条码后的结果,空税率")
        print(df_fentan_2[df_fentan_2["税率"].isnull()].to_markdown())

    # 整理数据格式
    df_fentan_2 = df_fentan_2[
        ["主体", "店铺", "平台", "品牌", "类别", "发货时间", "回款日期", "内部参考", "商家编码", "商品名称", "总回款", "总退款", "币种", "外币回款汇总", "外币退款汇总",
         "外币转换港币回款汇总", "外币转换港币退款汇总", "实际收款", "回款产品", "退款产品", "实际卖出", "平均价格", "总金额", "未税金额", "税率", "备注", "销售分类"]]

    # df1.to_excel("d:\debug_分摊_拼接前.xlsx")
    print("第二次分摊后,汇总记录;找到相同日期的记录+替换条码后的记录")
    df_fentan_2.rename(columns={"内部参考": "内部参考号"}, inplace=True)
    print("查看换品后的结果:")
    print(df_fentan_2.head(10).to_markdown())
    df_fentan_2.to_excel("d:\debug_换品后的.xlsx")
    df1 = df1.append(df_fentan_2)
    # df1.to_excel("d:\debug_分摊_拼接后.xlsx")
    # df1.to_excel(to_file.replace(".xlsx", "_test2.xlsx"), index=False)

    return df1


def apportion_paper(df_nomal, df_fentan):
    # 分摊正数，比如 分摊纸巾
    # 第一批最容易分摊的，找得到相同的日期。碰到纸巾进行分摊
    # zongjine = df["实际收款"].sum()
    pre_amnt_sum = df_nomal["实际收款"].sum()
    print("分摊检查(正常):", pre_amnt_sum)
    print("分摊检查(摊销):", df_fentan["实际收款"].sum())

    print("追查零价格2")
    print(df_nomal[df_nomal["平均价格"] == 0].to_markdown())

    sale_type = df_nomal["销售分类"].iloc[0]

    tag = "1111111111"
    # tag="xxxxxxxxxx"
    # df_fentan = df[(df["商家编码"].str.contains("11111"))]  # &(df["支付日期"].str.len>0)
    # 这是正常的数据
    # df_nomal = df[~df["商家编码"].str.contains(tag)]

    # 正常数据，不包含需要分摊的记录。把要分摊的扣掉
    # df_nomal = df[~(df["商家编码"].str.contains(tag)  | df["分类备注"].str.contains("分摊") )]

    if df_nomal.empty:
        print("完全没有正常可销售的产品，请人工摊销！", df_nomal["店铺"].iloc[0], df_nomal["销售分类"].iloc[0])
        return df_nomal

    # 这是包含1111特殊标志需要分摊的数据，或者指定要求分摊的，或者价格低于2元
    # df_fentan = df[(df["商家编码"].str.contains(tag)  | df["分类备注"].str.contains("分摊") )]  # &(df["支付日期"].str.len>0)
    df_fentan_group = df_fentan.groupby(["平台", "店铺", "发货时间", "index_plat_shop_paydate"]).agg(
        {"实际收款": np.sum}).reset_index()
    df_fentan_group.rename(columns={"实际收款": "money"}, inplace=True)

    print("查看分摊是否有数据")
    print(df_fentan_group.head(5).to_markdown())

    # print("剔除纸巾总金额:", df["money"].sum())
    print("所有需要分摊的纸巾:", df_fentan_group["money"].sum())
    # print(df_fentan.head(10).to_markdown())
    # df_fentan.to_excel(to_file.replace(".xlsx", "_所有需要分摊的纸巾.xlsx"), index=False)
    # df_fentan.columns=[["平台","店铺","支付日期","index_plat_shop_paydate","money"]]

    # print("追查 index_plat_shop_paydate 错误:")
    # print(df["index_plat_shop_paydate"].head(5).to_markdown())
    # print(df_fentan["index_plat_shop_paydate"].head(5).to_markdown())

    # 找到相同日期，分摊当前日期
    df_fentan_1 = df_fentan_group[
        (df_fentan_group["index_plat_shop_paydate"].isin(df_nomal["index_plat_shop_paydate"]))].copy()
    # 找不到相同日期的待分摊记录，替换条码（比如这些天只卖纸巾!）
    df_fentan_2 = df_fentan[~df_fentan.index_plat_shop_paydate.isin(df_nomal["index_plat_shop_paydate"])].copy()

    print("分摊预测：")
    # print("分摊前金额:", df["实际收款"].sum())
    print("正常金额:", df_nomal["实际收款"].sum())
    print("分摊1金额(摊销到相同日期):", df_fentan_1["money"].sum())
    print("分摊2金额（更换条码）:", df_fentan_2["实际收款"].sum())
    print("分摊后预计的金额:", df_nomal["实际收款"].sum() + df_fentan_1["money"].sum() + df_fentan_2["实际收款"].sum())

    # df_fentan_1=df_fentan[((df_fentan["index_plat_shop_paydate"].isin(df[~df["商家编码"].str.contains("11111")]["index_plat_shop_paydate"])) & (df["支付日期"].str.len()>0))].copy()
    # df_fentan_1=df_fentan_1[df_fentan_1["支付日期"].str.len()>0]
    # print("第01次分摊结果")
    print(df_fentan_1.head(3).to_markdown())
    # df_fentan_1.to_excel(to_file.replace(".xlsx", "_可以直接分摊到相同日期的纸巾.xlsx"), index=False)

    # 清洗掉11111
    # df = df[~df["商家编码"].str.contains("11111")]
    # print("测试测试测试11")
    # print(df[~df["商家编码"].str.contains("11111")].head(10).to_markdown())
    # print(df[~df["商家编码"].str.contains(tag)].head(10).to_markdown())

    # 找到与摊销相同日期的，不包含标记的正常记录（即剔除当然有摊销的纸巾的正常销售记录）
    df1 = df_nomal.merge(df_fentan_1[["index_plat_shop_paydate", "money"]], how="left",
                         on=["index_plat_shop_paydate"]).reset_index()

    # print("测试测试测试22")
    # print(df1.head(10).to_markdown())

    # print("剔除掉金额为0的纪录")
    # print(df1[~df1["money"].isnull()].head(10).to_markdown())
    df1["总金额_ori"] = df1["实际收款"]
    df1["商家编码_ori"] = df1["商家编码"]
    df1["平均价格_ori"] = df1["平均价格"]
    df1["sum"] = df1.groupby(by=["index_plat_shop_paydate"])["实际收款"].transform("sum")
    df1["cnt"] = df1.groupby(by=["index_plat_shop_paydate"])["实际收款"].transform("count")
    # df["sum"]=df.groupby(["index_plat_shop_paydate"])["实际收款"].sum()
    # print("debug_1")
    # print(df1.head(10).to_markdown())
    # print(df1[~df1["sum"].isnull()].head(10).to_markdown())
    # print(df1[df1["sum"].astype(int)==0].head(10).to_markdown())

    if df1.shape[0] == 0:
        # df.to_excel(to_file.replace(".xlsx", "_error_缺少有效的数据.xlsx"), index=False)
        print("完全没有可销售的产品")
        return df1

    # 计算分摊金额
    df1["share"] = df1.apply(lambda x: x["money"] * x["实际收款"] / x["sum"] if x["sum"] > 0 else x["money"] / x["cnt"],
                             axis=1)
    print("第一次分摊结果")
    # print(df1[~df1["share"].isnull()].head(10).to_markdown())
    # df1.to_excel(r"C:\Users\ns2033\Downloads\1111分摊_1.xlsx")

    print("调整最后的结果")
    # df1.to_excel(to_file.replace(".xlsx", "_分摊前_debug.xlsx"), index=False)
    # df1.to_excel(r"D:\数据处理\摊销\摊销后_debug.xlsx", index=False)

    print(df1[df1["平均价格"] == 0].to_markdown())

    df1["share"].fillna(0, inplace=True)
    df1["实际收款"] = df1["实际收款"] + df1["share"]
    df1["总回款"] = df1["总回款"] + df1["share"]

    # 总金额和实际收款是什么关系？
    df1["总金额"] = df1["实际收款"]
    df1["实际卖出"] = df1.apply(lambda x: 0 if x["平均价格"] == 0 else math.ceil(x["实际收款"] / x["平均价格"]), axis=1)
    df1["平均价格"] = df1["实际收款"] / df1["实际卖出"]
    df1["回款产品"] = df1["实际卖出"]
    df1["销售分类"] = sale_type
    df1["memo"] = ""
    # df1.to_excel(to_file.replace(".xlsx", "_test11.xlsx"), index=False)
    # df1.to_excel(r"C:\Users\ns2033\Downloads\1111分摊_2.xlsx")

    print("第一次分摊检查:")
    print(df1.head(5).to_markdown())
    # print(df1[df1["share"]>0].head(10).to_markdown())
    # df1.to_excel(to_file.replace(".xlsx", "_分摊后.xlsx"), index=False)

    a = df1["总金额_ori"].sum()
    b = df_fentan_1["money"].sum()
    c = df1["总金额"].sum()
    print("原来的金额：", a, "+ 摊销的金额: ", b,
          "=" if abs(df1["总金额_ori"].sum() + df_fentan_1["money"].sum() - df1["总金额"].sum()) < 0.01 else "<" if df1[
                                                                                                                  "总金额_ori"].sum() +
                                                                                                              df_fentan_1[
                                                                                                                  "money"].sum() <
                                                                                                              df1[
                                                                                                                  "总金额"].sum() else ">",
          c)

    df1["销售分类"] = sale_type

    print("实际收款为0的记录3:", df1[df1["实际收款"] == 0].shape[0])
    print(df1[df1["实际收款"] == 0].head(5).to_markdown())

    print("分摊第一步：", sale_type)
    print(df1[df1["销售分类"].str.contains(sale_type)]["实际收款"].sum())

    if df1[df1["实际收款"] < 2].shape[0] > 0:
        print("分摊第一步以后实际收款<2", df1["店铺"].iloc[0], df1["销售分类"].iloc[0])
        print(df1[df1["实际收款"] < 2].head(10).to_markdown())

    # df1.to_excel("d:\debug_分摊第一步.xlsx")
    # 如果没有发挥表
    # if str(type(df1)).find("None") > 0:
    # if not df1:
    #     # if df1 == None:
    #     # print("完全没有可销售的产品，分摊停止！")
    #     return

    print("第二步摊销,没有找到相同的支付日期，金额是正数。把这些纸巾的金额摊销到整个月的其他商品上")
    if df_fentan_2.shape[0] > 0:
        # if True:
        # 第二批  没有找到相同的支付日期，金额是正数
        # sale_type = df["销售分类"].iloc[0]
        # 价格中位数
        df_price = df_nomal.groupby(["商家编码"])["平均价格"].agg("median").reset_index()
        # print(df_price.head(3).to_markdown())
        df_price.columns = ["商家编码", "价格中位数"]
        print("查看价格中位数")
        print(df_price.head(10).to_markdown())

        # tag = "1111111111"
        # tag = "xxxxxxxxxx"
        # df_fentan_2 = df_fentan[((df_fentan["index"].isin(df["index"])) & (df["支付日期"].str.len == 0))]
        # print("第一批替换条码,实际卖出>0")
        # df_fentan_2 = df[((df["商家编码"].str.contains("11111")))]  # &(df["支付日期"].str.len>0)

        # 根据tag标志，找到需要被替换摊销的商品条码和记录行
        # df_fentan_2 = df[((df["商家编码"].str.contains(tag)))]  # &(df["支付日期"].str.len>0)
        # print(df_fentan_2.to_markdown())
        # & (~df["index_plat_shop_paydate"].isin(df["index_plat_shop_paydate"])

        # 跨天分摊
        if df_fentan_2.shape[0] > 0:
            print("无法按照相同【平台+店铺+日期】 匹配到正常的销售记录，需要跨天分摊!仍然采用撒胡椒面的方式")
            # df_fentan_2 = df_fentan_2[~df_fentan_2.index_plat_shop_paydate.isin(df_nomal["index_plat_shop_paydate"])]
            print("没有找到可分摊日期的纸巾:", df_fentan_2["实际收款"].sum())
            print(df_fentan_2.head(5).to_markdown())

            if True:
                print("发现负数了,需要做退货!")
                # df_fushu["index_plat_shop_month"] = df_fushu.apply(lambda x: "{}{}{}".format(x["平台"], x["店铺"], x["月份"]), axis=1)

                # 剩下是正数
                # df1 = df1[~((df1["分类备注"].str.contains("分摊") | df1["分类备注"].str.contains("做退货")) & (df1["实际收款"] < 0))]

                # df1["index_plat_shop_month"] = df1.apply(lambda x: "{}{}{}".format(x["平台"], x["店铺"], x["月份"]), axis=1)

                # df_fushu = df1[  ( df1["分类备注"].str.contains("分摊")   )   ]

                # 算出每个月需要摊销的负金额
                df_fentan_2_group = df_fentan_2.groupby(["index_plat_shop_month"]).agg(
                    实际收款=("实际收款", "sum")).reset_index()
                df_fentan_2_group.columns = ["index_plat_shop_month", "摊销纸巾金额"]
                print("核对需要摊销的纸巾总金额：", df_fentan_2_group["摊销纸巾金额"].sum())
                print(df_fentan_2_group.head(10).to_markdown())
                # if debug:
                #     df_fentan_2_group.to_excel(to_file.replace(".xlsx", "_退货汇总.xlsx"), index=False)

                # 如果当月只卖纸巾可怎么办？
                df_yiwai = df_fentan_2_group[
                    ~df_fentan_2_group["index_plat_shop_month"].isin(df1["index_plat_shop_month"])]
                if df_yiwai.shape[0] > 0:
                    print("太意外了！如果当月只有负数怎么办?累计到最后一个月份!")
                    print(df_yiwai.head(10).to_markdown())

                print("查看要摊销的负数金额")
                print(df_fentan_2_group.head(12).to_markdown())

                # df1["正常销售金额"]=df1.groupby(["平台", "店铺", "月份"])["实际收款"].agg("cumsum")
                df1["正常销售金额"] = df1.groupby(by=["index_plat_shop_month"])["实际收款"].transform("sum")
                df1 = df1.merge(df_fentan_2_group, how="left", on=["index_plat_shop_month"])
                df1["摊销纸巾金额"].fillna(0, inplace=True)  # 默认是负数
                df1["摊销比例"] = df1.apply(lambda x: x["摊销纸巾金额"] / x["正常销售金额"] if x["摊销纸巾金额"] > 0 else 0, axis=1)
                df1["实际收款"] = df1.apply(lambda x: x["实际收款"] + x["实际收款"] * x["摊销比例"], axis=1)

                df1["总金额"] = df1["实际收款"]
                df1["实际卖出"] = df1.apply(lambda x: math.ceil(x["实际收款"] / x["平均价格"]), axis=1)
                df1["平均价格"] = df1.apply(lambda x: 0 if x["实际卖出"] == 0 else x["实际收款"] / x["实际卖出"], axis=1)

                print("查看摊销后的效果")
                print(df1.head(20).to_markdown())

                print("分摊后前 {}，分摊后 {}".format(pre_amnt_sum, df1["实际收款"].sum()))

                # print("删除退货后:")
                # print(df1.head(10).to_markdown())
                # print("需要分摊的负数的总金额：", df_fushu["实际收款"].sum(), df_fushu.shape[0])
                # print("删除退货后的总金额：", df1["实际收款"].sum(), df1.shape[0])
                # print("删除退货后的总金额(加总)：", df1["实际收款"].sum() + df_fushu["实际收款"].sum())

                print("扣减后总金额:", df1["实际收款"].sum())

                del df1["摊销纸巾金额"]
                del df1["正常销售金额"]
                del df1["摊销比例"]

        # df1.to_excel(to_file.replace(".xlsx", "_test1.xlsx"), index=False)

        # 需要采用第二种分摊方法,当天没有找到正常销售的商品
        # if False:
        # 换条码
        change_type = ""
        # change_sku(df1, df_nomal, df_fentan_2, change_type)
        # print(df1.head(10).to_markdown())

        print("分摊2结果：", pre_amnt_sum, "+", df_fentan["实际收款"].sum(),
              "==" if abs(pre_amnt_sum + df_fentan["实际收款"].sum() - df1["实际收款"].sum()) < 0.01 else ">" if pre_amnt_sum +
                                                                                                         df_fentan[
                                                                                                             "实际收款"].sum() >
                                                                                                         df1[
                                                                                                             "实际收款"].sum() else "<",
              df1["实际收款"].sum(), " 差异:", "{:2f}".format(df1["实际收款"].sum() - pre_amnt_sum - df_fentan["实际收款"].sum()))

        df1["销售分类"] = sale_type

    if df1[df1["实际收款"] < 2].shape[0] > 0:
        print("分摊第二步以后实际收款<2", df1["店铺"].iloc[0], df1["销售分类"].iloc[0])
        print(df1[df1["实际收款"] < 2].head(10).to_markdown())

    print("纸巾分摊前 {}，后 {}".format(pre_amnt_sum + df_fentan["实际收款"].sum(), df1["实际收款"].sum()))
    return pd.DataFrame(df1)


def save_purchaseorder(df1, purchase_filename):
    # 生成采购单
    if df1.empty:
        return pd.DataFrame(
            columns=["主体", "供应商", "订单日期", "入库日期", "商家编码", "商品名称", "数量", "税率", "采购未税金额", "采购含税金额", "采购含税单价", "类型"])
    else:
        df2 = df1[["主体", "发货时间", "商家编码", "商品名称", "实际卖出", "采购未税金额", "采购含税金额", "采购含税单价", "销售分类"]].copy()
        df2["供应商"] = "深圳市麦凯莱科技有限公司"
        df2["税率"] = 0.13
        df2.rename(columns={"发货时间": "订单日期", "实际卖出": "数量", "销售分类": "类型"}, inplace=True)
        df2["入库日期"] = df2["订单日期"]
        df2 = df2[["主体", "供应商", "订单日期", "入库日期", "商家编码", "商品名称", "数量", "税率", "采购未税金额", "采购含税金额", "采购含税单价", "类型"]].copy()
        # 同步保存采购订单
        df2.to_excel(purchase_filename, index=False)
    return df2


def amortize(filename, platform, shop, to_file, purchase_filename, saletype, del_fn, add_fn):
    # 摊销
    # 负数也要摊销掉

    # 需要摊销条码特征
    tag = "1111111111"

    # 是否进行数据检查
    _check_data = True

    _path = os.path.dirname(to_file)
    if not os.path.exists(_path):
        os.makedirs(_path)

    df = pd.read_excel(filename)

    # 筛选数据
    print("筛选店铺:", shop)
    df = df[df["平台名称"] == platform]
    df = df[df["店铺名称"] == shop]
    df = df[df["销售分类"] == saletype]

    print(shop, saletype, " 记录数:", df.shape[0])

    # 抽查数据
    # df=df[df["商家编码"].str.contains("6930469054686")]
    df["实际收款"] = df["实际收款"].astype(float)
    # df=df[ abs(df["实际收款"]-56.33)<0.01]

    df["回款日期"] = df["回款日期"].astype(str)
    df["商家编码"] = df["商家编码"].astype(str)


    print("追踪空日期1")
    # print(df[  ( df["回款日期"].str.contains("2022-05-11,2022-05-09")  & df["平台名称"].str.contains("百度") )   ].to_markdown())
    print(df.head(10).to_markdown())
    print(df[df["发货时间"].isnull()].head(10).to_markdown())

    # 如果发货时间为空，取第一个回款日期
    # df["发货时间"] = df.apply(lambda x:  x["回款日期"].split(",")[0] if ( (len(str(x["发货时间"])) < 8) | (str(x["发货时间"])=='nan') | ( (x["发货时间"])==np.nan) | ( x["发货时间"]!=x["发货时间"])   )  else x["发货时间"], axis=1)
    # df["发货时间"] = df.apply(lambda x:  x["回款日期"].split(",").sort()[0] if  x["发货时间"]!=x["发货时间"]  else x["发货时间"], axis=1)

    if df.shape[0] > 0:
        df["发货时间"] = df.apply(lambda x: min(x["回款日期"].split(',')) if x["发货时间"] != x["发货时间"] else x["发货时间"], axis=1)
        # 判断是一个字符串是否为 nan 的方法：  a!=a 返回True
        # | ( math.isnan(x["发货时间"]) )

        print("追踪空日期2")
        print(df[(df["回款日期"].str.contains("2022-05-11,2022-05-09") & df["平台名称"].str.contains("百度"))].to_markdown())

    print("商家编码为空...000")
    print(df[df["商家编码"].str.contains("nan")].to_markdown())

    # 字段清洗
    df.rename(columns={"金额": "总金额", "总回款金额": "总回款", "总退款金额": "总退款", "实际卖出数量": "实际卖出", "平台名称": "平台", "账单公司主体": "主体",
                       "店铺名称": "店铺", "商品类别": "类别", "品牌中文名称": "品牌"}, inplace=True)
    df["发货时间"] = pd.to_datetime(df["发货时间"])
    df["备注"].fillna("", inplace=True)

    # 此处错误!
    df["总金额"] = df["实际收款"]
    df["总金额"].map(lambda x: '{:.2f}'.format(x))
    df["商家编码"] = df["商家编码"].astype(str)
    df["商家编码"] = df["商家编码"].apply(lambda x: x.replace("'", ""))

    if "税率" in df.columns:
        df["税率"] = df["税率"].astype(str)
    else:
        df_tax = pd.read_excel(tax_file)
        company = df.iloc[0, 0]

        df_tax = df_tax[df_tax["主体"].str.contains(company)][["税率"]]
        tax = df_tax.iloc[0, 0]
        df["税率"] = tax
        print("当前公司是:", company, tax)

    df["税率"] = df["税率"].astype(str)

    if not ("未税金额" in df.columns):
        df["未税金额"] = df["总金额"] / (1 + (df["税率"].astype(float)) * 1.00000000)

    # 清理零单价 ，如果数量和价格都是0，则强制数量为1
    # 异常数据，暂时不处理
    if False:
        df["实际卖出"] = df.apply(
            lambda x: 1 if ((x["实际卖出"] == 0) & (x["平均价格"] == 0)) else x["实际卖出"], axis=1)

        # 如果平均价格=0，则重新计算价格
        df["平均价格"] = df.apply(
            lambda x: x["实际收款"] / x["实际卖出"] if x["平均价格"] == 0 else x["平均价格"], axis=1)

    # 只保留2022年数据
    # df = df[df["发货时间"].dt.year == 2022]

    # print("抽查数据222")
    # print(df.head(10).to_markdown())

    print("测试1")
    zongjine = df["实际收款"].sum()
    # print(filename, "总金额：", zongjine, df.shape[0])
    print("删除退货前的总金额：", zongjine, df.shape[0])

    # 删除0金额的
    # df["金额需做分摊"].fillna("",inplace=True)
    df["分类备注"].fillna("", inplace=True)
    # df=df[~df["金额需做分摊"].str.contains("删除")]
    df = df[~df["分类备注"].str.contains("删除")]

    # print("追查零价格00xxx")
    # print(df[df["平均价格"] == 0].head(10).to_markdown())

    # 处理退货
    # print("处理退货")
    print("处理退货前：", saletype, shop)
    print(df[df["销售分类"].str.contains(saletype)]["实际收款"].sum())

    # 处理退货
    df, df_fentan = refund_amount_new(df, to_file)
    print(saletype, shop, " 处理退货后：", df["实际收款"].sum() + df_fentan["实际收款"].sum())
    print(df[df["销售分类"].str.contains(saletype)]["实际收款"].sum(), df_fentan["实际收款"].sum())

    if df.shape[0] == 0:
        print("处理完退货后，完全没有正常可销售的产品，请人工摊销！")

    print("实际收款为0的记录2:", df[df["实际收款"] == 0].shape[0])
    print(df[df["实际收款"] == 0].head(5).to_markdown())

    print("最总总金额")
    print(df[df["销售分类"].str.contains(saletype)]["总金额"].sum())

    print("追查零价格00")
    print(df[df["平均价格"] == 0].to_markdown())

    print(filename, "摊销退货后，扣减前的总金额：", df["总金额"].sum(), df.shape[0])
    print("有{}个摊销".format(len(del_fn)))
    print(df.head(10).to_markdown())

    # 处理补丁,增加和删除的建议
    # df=patch_advise(df,del_fn,add_fn,to_file)

    df["备注"] = df["备注"].astype(str)
    df["分类备注"] = df["分类备注"].astype(str)
    # 价格<=2，把条码替换成1111

    # df["商家编码"]=df.apply(lambda x: "1111111111115"  if  x["平均价格"]<=2 else x["商家编码"] ,axis=1 )
    # # 商务单参与摊销
    if df.shape[0] > 0:
        df["商家编码"] = df.apply(lambda x: "1111111111119" if x["备注"].find("商务单") >= 0 else x["商家编码"], axis=1)
        # 指定分摊的正金额
        # df["商家编码"] = df.apply(lambda x: "1111111111119" if ((x["金额需做分摊"].find("分摊") >= 0) & (x["实际收款"]>0)) else x["商家编码"], axis=1)
        df["商家编码"] = df.apply(
            lambda x: "1111111111119" if ((x["分类备注"].find("分摊") >= 0) & (x["实际收款"] >= 0)) else x["商家编码"], axis=1)
    # df_fushu = df1[(df1["金额需做分摊"].str.contains("分摊") & (df1["实际收款"] > 0))]
    print("还有需要做分摊的吗？")
    print(df[df["分类备注"].str.contains("分摊")].head(10).to_markdown())

    df["发货时间"].fillna("", inplace=True)
    df["id"] = df.index

    if debug:
        df.to_excel(to_file.replace(".xlsx", "_原始文件.xlsx"), index=False)

    # 计算价格中位数
    df_price = df.groupby(["商家编码"])["平均价格"].agg("median").reset_index()
    # print(df_price.head(3).to_markdown())
    df_price.columns = ["商家编码", "价格中位数"]
    print("查看价格中位数")
    print(df_price.head(10).to_markdown())

    print("分摊第一步之前：", saletype)
    print(df[df["销售分类"].str.contains(saletype)]["实际收款"].sum())

    print(df[df["实际收款"] < 2].head(10).to_markdown())
    # df.to_excel("d:\debug_分摊第一步之前_001.xlsx")
    # 第一批最容易分摊的，找得到相同的日期
    if df.shape[0] > 0:
        # 分摊纸巾
        df1 = apportion_paper(df, df_fentan)
    else:
        print("不需要进一步摊销！")
        df1 = df.copy()

    print("纸巾分摊完金额=", df1["实际收款"].sum())
    # df1.to_excel("d:\debug_分摊后_002.xlsx")
    # 第二批  没有找到相同的支付日期
    # df1=fentan_step_2(df,df1)
    # print("跟踪空值")
    # print(zongjine)

    print("step1后立即计算合计金额: {}-->{}".format(zongjine, df1["实际收款"].sum()))

    if df1.shape[0] == 0:
        print("找不到合计数")
    else:
        print(df1["实际收款"].sum())

    print("总金额变化：", zongjine,
          "==" if abs(zongjine - df1["实际收款"].sum()) < 0.01 else ">" if zongjine > df1["实际收款"].sum() else "<",
          df1["实际收款"].sum())

    cnt1 = df1.shape[0]
    df1 = df1.merge(df_price, how="left", on=["商家编码"])
    cnt2 = df1.shape[0]
    print("记录数发生了变化:", cnt1, "->", cnt2, ":", cnt2 - cnt1)

    print("价格纠错前！", df1.shape[0])
    # df1.to_excel("d:\debug_00003.xlsx")
    print(df1[df1["平均价格"] < 2].head(10).to_markdown())

    if df1[df1["实际收款"] < 2].shape[0] > 0:
        print("实际收款<2,test1")
        print(df1[df1["实际收款"] < 2].head(10).to_markdown())

    if df1.shape[0] > 0:
        df1["价格中位数"].fillna(0, inplace=True)
        df1["价格中位数"] = df1.apply(lambda x: x["平均价格"] if x["价格中位数"] == 0 else x["价格中位数"], axis=1)

        df1["平均价格"].fillna(0, inplace=True)
        df1["平均价格"] = df1.apply(lambda x: x["价格中位数"] if ((x["平均价格"] < 1) | (x["平均价格"] > 1000)) else x["平均价格"], axis=1)
        df1["平均价格"] = df1.apply(
            lambda x: 0 if x["实际卖出"] == 0 else x["实际收款"] / x["实际卖出"] if x["平均价格"] == 0 else x["平均价格"], axis=1)

        # df1["平均价格"]=df1["平均价格"].astype("str")
        # df1["平均价格"]=df1["平均价格"].apply(lambda x: 0 if ((x=="") | (x.find("nan")>=0)) else x  )
        # df1["平均价格"] = df1["平均价格"].astype(float)

        # print("抽查数据:")
        # print(df[df["index"]==1076][["index","总回款","回款产品数量","实际卖出","平均价格"]])

        # 如果实际卖出=0，则重新计算，可以给1
        df1["实际卖出"] = df1.apply(lambda x: x["回款产品数量"] if x["实际卖出"] == 0 else x["实际卖出"], axis=1)
        df1["实际卖出"] = df1.apply(
            lambda x: math.ceil(x["实际收款"] / x["价格中位数"]) if ((x["实际卖出"] == 0) & (x["平均价格"] == 0)) else math.ceil(
                x["实际收款"] / x["平均价格"]), axis=1)

        # 解决实际卖出还是0的问题
        # df1["实际卖出"] = df1.apply(lambda x: x["回款产品数量"] if ["实际卖出"] == 0 else x["实际卖出"] ,axis=1)
        # df1["实际卖出"] = df1["实际卖出"].apply(lambda x: 1 if x == 0 else x)

        # print(df1[   ((df1["index"]>=1814) & (df1["index"]<=1815))) ].to_makdown()

        print(df1.sort_values(by=["平均价格"], ascending=True).head(10).to_markdown())
        # df1.to_excel(r"d:\test_debug_1.xlsx")

        # df1["实际卖出"] = df1.apply(lambda x: math.ceil(x["实际收款"] / x["平均价格"]), axis=1)
        df1["平均价格"] = df1["实际收款"] / df1["实际卖出"]
        print("价格纠错后！")
        print(df1[df1["平均价格"] < 2].head(10).to_markdown())

        print("价格纠错纠错:")
        print(df1[df1["平均价格"] < 1].head(10).to_markdown())

        print("分摊第二步：")
        print(df1[df1["销售分类"].str.contains(saletype)]["总金额"].sum())

        print("实际收款为0的记录4:", df1[df1["实际收款"] == 0].shape[0])
        print(df1[df1["实际收款"] == 0].head(5).to_markdown())

    print("总金额变化：", zongjine,
          "==" if abs(zongjine - df1["实际收款"].sum()) < 0.01 else ">" if zongjine > df1["实际收款"].sum() else "<",
          df1["总金额"].sum())

    if abs(zongjine - df1["实际收款"].sum()) >= 0.01:
        # raise error
        print("总金额不相等!")
        # print(100=="abcd")

    print("跟踪结果")
    if debug:
        df1.to_excel(to_file.replace(".xlsx", "_debug.xlsx"), index=False)

    if df1[df1["平均价格"] >= 500].shape[0] > 0:
        # df1[df1["平均价格"] >= 500].to_excel(to_file.replace(".xlsx", "_error_价格偏高请检查.xlsx"), index=False)
        print(df1[df1["平均价格"] >= 500].head(100).to_markdown())

    # 最终清理结果
    del df1["index_plat_shop"]
    del df1["index_plat_shop_paydate"]
    del df1["index_plat_shop_recvdate"]

    if "index" in df1.columns:
        del df1["index"]
    if "money" in df1.columns:
        del df1["money"]
    if "id" in df1.columns:
        del df1["id"]
    if "sum" in df1.columns:
        del df1["sum"]
    if "cnt" in df1.columns:
        del df1["cnt"]
    if "share" in df1.columns:
        del df1["share"]
    if "总金额_ori" in df1.columns:
        del df1["总金额_ori"]
    if "memo" in df1.columns:
        del df1["memo"]
    if "商家编码_ori" in df1.columns:
        del df1["商家编码_ori"]
    if "平均价格_ori" in df1.columns:
        del df1["平均价格_ori"]
    # del df1["level_0"]

    # df1.to_excel(r"d:\test_debug_2.xlsx")

    print("实际收款为0的记录5:", df1[df1["实际收款"] == 0].shape[0])
    print(df1[df1["实际收款"] == 0].head(5).to_markdown())

    if df1[df1["税率"].str.contains("\+")].shape[0] > 0:
        df1.to_excel(to_file.replace(".xlsx", "_error_税率问题.xlsx"), index=False)
        print(filename, "税率有问题")
        return

    df1["总金额"].map(lambda x: '{:.2f}'.format(x))

    # 补充字段
    df1["税率"] = df1["税率"].astype("float64")
    if df1.shape[0] > 0:
        df1["未税金额"] = df1.apply(lambda x: x[["总金额"]] / (1 + x["税率"]), axis=1)
        df1["采购未税金额"] = df1["未税金额"] * 0.7
        df1["采购含税金额"] = df1.apply(
            lambda x: x["采购未税金额"] if x["税率"] == 0.01 else x["采购未税金额"] if x["税率"] == 0.03 else x["采购未税金额"] * (
                        1 + x["税率"]), axis=1)

        # 补齐币种
        money_type = df1["币种"].iloc[0]
        df1["币种"].fillna(money_type, inplace=True)
        df1["币种"] = df1["币种"].astype("str")
        df1["币种"] = df1.apply(lambda x: money_type if (
                    (len(x["币种"]) < 0) | (x["币种"] == "nan") | (x["币种"] == "") | (x["币种"].find("nan") >= 0)) else x[
            "币种"], axis=1)

        print("实际收款为0的记录6:", df1[df1["实际收款"] == 0].shape[0])
        print(df1[df1["实际收款"] == 0].head(5).to_markdown())

        df1["采购含税单价"] = df1.apply(lambda x: x["采购含税金额"] / x["实际卖出"], axis=1)

    # print("发现空白币种：")
    # df1.to_excel(r"d:\test_debug_3.xlsx")
    # df1.to_excel(r"d:\testkkk_零实际卖出.xlsx")

    # df1["类型"]=saletype

    # df1.to_excel("d:\debug_保存前.xlsx")

    print("汇总统计:")
    print(df1.groupby(["平台", "店铺"]).agg(总金额=("总金额", "sum")).reset_index())
    # 保存摊销结果
    # df1.to_excel(to_file, index=False)
    df1.to_excel(to_file.replace(".xls", "_{}_{}.xls".format(platform, shop)))
    # df1.to_excel(to_file.replace(".xls", "_bug.xls"))

    msg = ""
    # 生成采购单
    # if ~df1.empty:
    df2 = save_purchaseorder(df1, purchase_filename.replace(".xls", "_{}_{}.xls".format(platform, shop)))
    # else:
    #     msg="{} {} 完全没有正常可销售的产品，请人工摊销!".format(shop,saletype)

    loss = []
    if _check_data:
        loss_sku, loss_unit, loss_company, loss_shop = check_data(df)

    return df1, df2, loss_sku, loss_unit, loss_company, loss_shop, msg


def check_data(df):
    df_sku = pd.read_excel(sku_file)
    df_warehouse = pd.read_excel(warehouse_file, sheet_name="Sheet2")

    # print("查询仓库000:",warehouse_file)
    # print(df_warehouse.head(10).to_markdown())

    loss_unit = []
    loss_sku = []
    loss_company = []
    loss_shop = []

    df_no_unit = df[((~df["商家编码"].str.upper().isin(df_sku["条码"].str.upper())) & (
        ~df["商家编码"].str.upper().isin(df_sku["内部参考"].str.upper())))]
    if df_no_unit.shape[0] > 0:
        print("没有找到产品单位的有：", df_no_unit["商家编码"].unique())
        loss_unit.extend(df_no_unit["商家编码"].unique())

    df_sku["条码"] = df_sku["条码"].astype(str)
    # df_no_sku = df[~df["商家编码"].isin(df_sku["条码"])]
    df_no_sku = df[((~df["商家编码"].str.upper().isin(df_sku["条码"].str.upper())) & (
        ~df["商家编码"].str.upper().isin(df_sku["内部参考"].str.upper())))]
    if df_no_sku.shape[0] > 0:
        print("没有找到条码的有：", df_no_sku["商家编码"].unique())
        loss_sku.extend(df_no_sku["商家编码"].unique())

    # print("查询仓库")
    # print(df_warehouse.head(10).to_markdown())

    df_no_compid = df[~df["主体"].isin(df_warehouse["公司"])]
    if df_no_compid.shape[0] > 0:
        print("公司主体找不到编码的有：", df_no_compid["主体"].unique())
        loss_company.extend(df_no_compid["主体"].unique())

    df_shop = pd.read_excel(shop_file)
    df["店铺2"] = df["店铺"].apply(lambda x: str.upper(x))
    df_shop["财务店铺名称2"] = df_shop["财务店铺名称"].apply(lambda x: str.upper(x))
    df = df.merge(df_shop[["平台", "财务店铺名称2", "Odoo店铺名称"]], how="left", left_on=["平台", "店铺2"],
                  right_on=["平台", "财务店铺名称2"])

    no_shop = df[df["Odoo店铺名称"].isnull()][["平台", "店铺"]]
    if no_shop.shape[0] > 0:
        print("没有匹配到财务店铺:")
        print(no_shop.to_markdown())
        loss_shop.extend(no_shop["店铺"].unique())

    print("找不到条码的记录数:", len(df_no_unit["商家编码"].unique()) + len(df_no_sku["商家编码"].unique()))
    print("找不到公司主体:", df_no_compid["主体"].unique())
    print("找不到店铺:", no_shop["店铺"].unique())
    # return df_no_unit["商家编码"].unique()
    return loss_sku, loss_unit, loss_company, loss_shop


def tanxiao2(filename, del_fn, add_fn):
    # 自动摊销，生成文件

    # saletype="2C"
    # saletype="2C发出商品"

    df = pd.read_excel(filename)[["平台名称", "销售分类", "店铺名称", "实际收款"]]
    print("分摊前统计汇总:")
    df_pre_tongji = df.groupby(["平台名称", "店铺名称", "销售分类"])["实际收款"].agg("sum").reset_index()
    print(df_pre_tongji.to_markdown())

    loss_sku_list = []
    loss_unit_list = []
    loss_company_list = []
    loss_shop_list = []

    debug_shop = ""
    # _shop = "芭葆兔美容护肤专营店"
    # debug_shop = "樱加美旗舰店"
    for saletype in ["2C", "2C发出商品"]:
        # for saletype in ["2C发出商品"]:
        # for saletype in ["2C"]:
        print(df.head(5).to_markdown())
        # df = df[df["销售分类"] == saletype]

        if debug_shop == '':
            ori_amnt = df[df["销售分类"] == saletype]["实际收款"].sum()
        else:
            ori_amnt = df[((df["销售分类"] == saletype) & (df["店铺名称"] == debug_shop))]["实际收款"].sum()

        print(saletype, "原始实际收款:", ori_amnt)

        df_shop = df[df["销售分类"] == saletype][["平台名称", "店铺名称"]].copy()
        # del df_shop["总金额"]
        # del df_shop["销售分类"]
        df_shop.drop_duplicates(subset=["平台名称", "店铺名称"], inplace=True)
        if debug_shop > '':  # 如果指定店铺信息
            df_shop = df_shop[df_shop["店铺名称"] == debug_shop]

        amnt_list = []
        # 2C / 2C发出商品
        # if True:
        for index, shop in df_shop.iterrows():
            _platform = shop["平台名称"]
            _shop = shop["店铺名称"]

            print("摊销店铺店铺店铺:", _shop, saletype)
            _path = os.path.dirname(filename)
            _path = _path.replace("摊销前", "摊销后")
            # D:\数据处理\摊销\摊销前 替换成 D:\数据处理\摊销\摊销后
            # 目录中增加2c/2c发出商品 ，如果目录不存在，则创建目录
            _path = _path + os.sep + saletype + os.sep + _shop
            if not os.path.exists(_path):
                os.makedirs(_path)

            _filename = os.path.basename(filename)
            oms_so_file = _path + os.sep + _filename.replace(".xls", "_已摊销.xls")
            oms_po_file = _path + os.sep + _filename.replace(".xls", "_采购订单.xls")
            odoo_so_file = _path + os.sep + _filename.replace(".xls", "_odoo销售订单.xls")
            # 销售退货
            odoo_so_return_file = _path + os.sep + _filename.replace(".xls", "_odoo销售退货订单.xls")
            odoo_po_file = _path + os.sep + _filename.replace(".xls", "_odoo采购订单.xls")

            # 开始摊销 ,此处开始进行店铺过滤
            df1, df2, loss_sku, loss_unit, loss_company, loss_shop, msg = \
                amortize(filename, _platform, _shop, oms_so_file, oms_po_file, saletype, del_fn, add_fn)
            print("商家编码为空!")
            print(df1.sort_values(by=["商家编码"]).head(10).to_markdown())
            print(df1[df1["商家编码"].str.len() <= 3].to_markdown())

            loss_sku_list.extend(loss_sku)
            loss_unit_list.extend(loss_unit)
            loss_company_list.extend(loss_company)
            loss_shop_list.extend(loss_shop)

            zongjine = df1["实际收款"].sum()
            print("摊销后实际收款:", zongjine)
            amnt_list.append(zongjine)

            dict = {"平台名称": _platform, "店铺名称": _shop, "销售分类": saletype, "实际收款": zongjine}
            if "df_tanxiaohou" in vars():
                df_tanxiaohou = df_tanxiaohou.append(pd.DataFrame([dict]))
            else:
                df_tanxiaohou = pd.DataFrame([dict])

            # df1["金额需做分摊"]=df1["金额需做分摊"].astype(str)
            df1["分类备注"] = df1["分类备注"].astype(str)
            # 排除掉退货，作为销售订单
            # convert_sales_2022(df1[~df1["金额需做分摊"].str.contains("做退货")],odoo_so_file)
            # shop["店铺名称"],
            # df_sales=convert_sales_2022(df1[~df1["分类备注"].str.contains("做退货")],saletype,odoo_so_file)
            if ~df1.empty:
                df_sales = convert_sales_2022(df1, saletype, odoo_so_file)

            # 销售退货
            # convert_sales_tuihuo_2022(df1[df1["金额需做分摊"].str.contains("做退货")], odoo_so_return_file)
            # shop["店铺名称"],
            # df_return=convert_sales_tuihuo_2022(df1[df1["分类备注"].str.contains("做退货")],saletype, odoo_so_return_file)

            # 采购订单 shop["店铺名称"],
            df_purcharse = convert_purchase_2022(df2, saletype, odoo_po_file)

            if "sum_df1" in vars():
                sum_df1 = sum_df1.append(df1)
                sum_df2 = sum_df2.append(df2)
                if "sum_sales" in vars():
                    print("这里有数据吗？")
                    print(df_sales.head(3).to_markdown())
                    if df_sales.shape[0] > 0:
                        sum_sales = sum_sales.append(df_sales)
                        sum_purcharse = sum_purcharse.append(df_purcharse)
                # if str(type(df_return)).find("None")<0:
                #     if "sum_return" in vars():
                #         sum_return = sum_return.append(df_return)
                #     else:
                #         sum_return = df_return
            else:
                sum_df1 = df1
                sum_df2 = df2
                if "df_sales" in vars():
                    # if ~df_sales.empty:
                    if df_sales.shape[0] > 0:
                        sum_sales = df_sales
                        sum_purcharse = df_purcharse
                # if str(type(df_return)).find("None") < 0:
                #     sum_return = df_return

            # print("新文件名:")
            # print(filename,"-->",oms_so_file)

        # df.to_excel(new_filename)
        amnt = 0
        print(amnt_list)
        for x in amnt_list:
            amnt = amnt + float(x)

        print("摊销前实际收款为:", ori_amnt)
        print("摊销后实际收款为:", amnt_list, "=", amnt)
        print(saletype, " 摊销前后:",
              ": {} 完全相等 {}".format(ori_amnt, amnt) if abs(ori_amnt - amnt) < 0.01 else "{} 不相等 {}".format(ori_amnt,
                                                                                                           amnt))

    print("汇总查询商家编码为空!")
    sum_df1["内部参考号"].fillna("", inplace=True)
    sum_df1["商家编码"].fillna("", inplace=True)
    print(sum_df1.sort_values(by=["商品名称"], ascending=True).head(10)[
              ["发货时间", "内部参考号", "商家编码", "商品名称", "总回款"]].to_markdown())
    # print(sum_df1[sum_df1["商品名称"].str.len() <= 5].to_markdown())

    sum_df1.to_excel(r"d:\tttt.xls")

    _path = os.path.dirname(filename)
    _path = _path.replace("摊销前", "摊销后")
    # D:\数据处理\摊销\摊销前 替换成 D:\数据处理\摊销\摊销后
    # 目录中增加2c/2c发出商品 ，如果目录不存在，则创建目录
    # _path = _path + os.sep + saletype
    if not os.path.exists(_path):
        os.makedirs(_path)
    _filename = os.path.basename(filename)

    oms_so_file = _path + os.sep + _filename.replace(".xls", "_已摊销.xls")
    oms_po_file = _path + os.sep + _filename.replace(".xls", "_采购订单.xls")
    odoo_so_file = _path + os.sep + _filename.replace(".xls", "_odoo销售订单.xls")
    # 销售退货
    # odoo_so_return_file = _path + os.sep + _filename.replace(".xls", "_odoo销售退货订单.xls")
    odoo_po_file = _path + os.sep + _filename.replace(".xls", "_odoo采购订单.xls")

    sum_df1.to_excel(oms_so_file, index=False)
    sum_df2.to_excel(oms_po_file, index=False)
    sum_sales.to_excel(odoo_so_file, index=False)
    sum_purcharse.to_excel(odoo_po_file, index=False)
    # if "sum_return" in vars():
    #     if str(type(sum_return)).find("None") < 0:
    #         sum_return.to_excel(odoo_so_return_file)

    print("摊销完最后结果:")
    print(df_tanxiaohou.to_markdown())

    # 验收标准
    # df_fn=pd.read_excel(r"D:\数据处理\摊销\摊销前\2022年1-6月店铺收入汇总表7.27.xlsx")
    df_fn = pd.read_excel(r"Z:\it审计处理需求\odoo导入\2022年\2022年1-6月店铺收入汇总表7.27.xlsx")
    df_fn.rename(columns={"验收标准：22年1-6月收入金额（含税）": "2C", "22年发出商品金额": "2C发出商品"}, inplace=True)
    df_fn = df_fn[["平台名称", "店铺名称", "2C", "2C发出商品"]]
    df_fn1 = df_fn[["平台名称", "店铺名称", "2C"]].copy()
    df_fn1.rename(columns={"2C": "实际收款"}, inplace=True)
    df_fn1["销售分类"] = "2C"

    df_fn2 = df_fn[["平台名称", "店铺名称", "2C发出商品"]].copy()
    df_fn2.rename(columns={"2C发出商品": "实际收款"}, inplace=True)
    df_fn2["销售分类"] = "2C发出商品"

    df_fn3 = df_fn1.append(df_fn2)

    df_fn3 = df_fn3.groupby(["平台名称", "店铺名称", "销售分类"])["实际收款"].agg("sum").reset_index()

    print("找不到条码的有：", loss_sku_list)
    print("找不到单位的有：", loss_unit_list)
    print("找不到公司主体的有：", loss_company_list)
    print("找不到店铺的有：", loss_shop_list)
    print(msg)

    print("比对摊销结果:")
    df_result = df_pre_tongji.merge(df_tanxiaohou, how="left", on=["平台名称", "店铺名称", "销售分类"]).merge(df_fn3, how="left",
                                                                                                  on=["平台名称", "店铺名称",
                                                                                                      "销售分类"])
    df_result.rename(columns={"实际收款_x": "摊销前金额", "实际收款_y": "摊销后金额", "实际收款": "财务实际收款"}, inplace=True)
    # df_result["差异"]=df_result["摊销后金额"]-df_result["财务总金额"]
    df_result["差异"] = df_result.apply(lambda x: "{:.2f}".format(x["摊销后金额"] - x["财务实际收款"]), axis=1)
    print(df_result.to_markdown())

    print("自动拷贝文件")
    # \\file.maclove.com\IT审计\it审计处理需求\odoo导入\2022年\账套转格式\摊销后
    # \\file.maclove.com\审计对接财务共享\已摊销\茱莉珂丝（深圳）化妆品有限公司

    # Z:\it审计处理需求\odoo导入\2022年\账套转格式\摊销前\盈养泉（深圳）化妆品有限公司2C.xlsx

    # mycopyfile(oms_so_file,oms_so_file.replace(r"IT审计\it审计处理需求\odoo导入\2022年\账套转格式\摊销后","审计对接财务共享\已摊销\{}".format(_filename.replace("2C",""))) )
    # mycopyfile(oms_po_file,oms_po_file.replace(r"IT审计\it审计处理需求\odoo导入\2022年\账套转格式\摊销后","审计对接财务共享\已摊销\{}".format(_filename.replace("2C",""))) )
    # mycopyfile(odoo_so_file,odoo_so_file.replace(r"IT审计\it审计处理需求\odoo导入\2022年\账套转格式\摊销后","审计对接财务共享\已摊销\{}".format(_filename.replace("2C",""))) )
    # mycopyfile(odoo_so_return_file,odoo_so_return_file.replace(r"IT审计\it审计处理需求\odoo导入\2022年\账套转格式\摊销后","审计对接财务共享\已摊销\{}".format(_filename.replace("2C",""))) )
    # mycopyfile(odoo_po_file,odoo_po_file.replace(r"IT审计\it审计处理需求\odoo导入\2022年\账套转格式\摊销后","审计对接财务共享\已摊销\{}".format(_filename.replace("2C",""))) )

    print("文件拷贝结束！")

    # 不分类型
    # print(df_pre_tongji.groupby(["店铺名称"])["总金额"].agg("sum").reset_index().merge(df_tanxiaohou.groupby(["店铺名称"])["总金额"].agg("sum").reset_index(),how="left",on=["店铺名称"]).to_markdown())


def mycopyfile(srcfile, dstpath):  # 复制函数
    if not os.path.isfile(srcfile):
        # os.makedirs(srcfile)
        print("%s not exist!" % (srcfile))
    else:
        # if True:
        fpath, fname = os.path.split(srcfile)  # 分离文件名和路径
        if not os.path.exists(dstpath):
            os.makedirs(dstpath)  # 创建路径
        shutil.copy(srcfile, dstpath + fname)  # 复制文件
        print("copy %s -> %s" % (srcfile, dstpath + fname))


def read_company(company):
    ToC_money = 0
    ToC_notax = 0
    ToC_money = 0
    ToC_pur_havtax = 0
    ToC_pur_notax = 0

    Fachu_money = 0
    Fachu_notax = 0
    Fachu_pur_havtax = 0
    Fachu_pur_notax = 0

    if os.path.exists(output_dir_2C + os.sep + "2C{}.xlsx".format(company)):
        df_2c = pd.read_excel(output_dir_2C + os.sep + "2C{}.xlsx".format(company))
        # df_2c=pd.read_excel( r"C:\Users\ns2033\Downloads\摊销\需处理的\2C\2C{}.xlsx".format(company)   )
        ToC_money = df_2c["总金额"].sum()
        ToC_notax = df_2c["未税金额"].sum()
        ToC_pur_havtax = df_2c["采购未税金额"].sum()
        ToC_pur_notax = df_2c["采购含税金额"].sum()

    if os.path.exists(output_dir_2C发出商品 + os.sep + "2C发出商品{}.xlsx".format(company)):
        df_fachu = pd.read_excel(output_dir_2C发出商品 + os.sep + "2C发出商品{}.xlsx".format(company))
        # df_fachu=pd.read_excel(r"C:\Users\ns2033\Downloads\摊销\自动处理结果\需处理的\2C发出商品{}.xlsx".format(company) )
        Fachu_money = df_fachu["总金额"].sum()
        Fachu_notax = df_fachu["未税金额"].sum()
        Fachu_pur_havtax = df_fachu["采购未税金额"].sum()
        Fachu_pur_notax = df_fachu["采购含税金额"].sum()

    dict = {"主体": company, "ToC_money": ToC_money, "ToC_notax": ToC_notax, "ToC_pur_havtax": ToC_pur_havtax,
            "ToC_pur_notax": ToC_pur_notax,
            "Fachu_money": Fachu_money, "Fachu_notax": Fachu_notax, "Fachu_pur_havtax": Fachu_pur_havtax,
            "Fachu_pur_notax": Fachu_pur_notax
            }

    # print(dict)
    # df= pd.DataFrame.from_dict(dict).transpose()
    # df= pd.DataFrame.from_dict(dict,orient='index').reset_index()
    df = pd.DataFrame([dict])
    # df.rename(columns={'index': 'item', 0: 'value'}, inplace=True)
    # df= pd.DataFrame.from_dict(dict,orient='index')
    # pd.pivot_table(df, index=[u'主体'])
    # print(df.stack())
    # df.columns=["key","value"]
    # df=df.reset_index(drop=True)
    # print(df)
    # df=df.unstack()
    print("转换格式")
    print(df)
    # return pd.DataFrame.from_dict(dict,orient='index')
    # df = df.T
    # return pd.DataFrame.from_dict(dict).T
    return df


def read_shop(company):
    if os.path.exists(output_dir_2C + os.sep + "2C{}.xlsx".format(company)):
        df_2c = pd.read_excel(output_dir_2C + os.sep + "2C{}.xlsx".format(company))
        df_2c = df_2c.groupby(["主体", "平台", "店铺"])["总金额"].sum().reset_index()
        df_2c.columns = ["主体", "平台", "店铺", "2C总金额"]
        df_2c["店铺"] = df_2c["店铺"].apply(lambda x: str.upper(x))
        shop_list1 = df_2c[["主体", "平台", "店铺"]]


        if os.path.exists(output_dir_2C发出商品 + os.sep + "2C发出商品{}.xlsx".format(company)):
            df_fachu = pd.read_excel(output_dir_2C发出商品 + os.sep + "2C发出商品{}.xlsx".format(company))
            # df_fachu=pd.read_excel(r"C:\Users\ns2033\Downloads\摊销\自动处理结果\需处理的\2C发出商品{}.xlsx".format(company) )
            # Fachu_money = df_fachu["总金额"].sum()
            df_fachu = df_fachu.groupby(["主体", "平台", "店铺"])["总金额"].sum().reset_index()
            df_fachu.columns = ["主体", "平台", "店铺", "发出总金额"]

            # print("这里是发出商品")
            df_fachu["店铺"] = df_fachu["店铺"].apply(lambda x: str.upper(x))

            shop_list2 = df_fachu[["主体", "平台", "店铺"]]

            shop_list = shop_list1.append(shop_list2)
            shop_list = shop_list.drop_duplicates()


            df_2c = shop_list.merge(df_2c, how="left", on=["主体", "平台", "店铺"])
            # print("跟踪001")
            # print(df_2c.to_markdown())

            df_2c = df_2c.merge(df_fachu, how="left", on=["主体", "平台", "店铺"])
            # print("跟踪002")
            # print(df_2c.to_markdown())

            df_2c.fillna(0, inplace=True)
        else:
            print("2C发出商品{} 不存在".format(company))
            df_2c["发出总金额"] = 0

        return df_2c
    else:
        print("2C{} 不存在".format(company))
        # return pd.DataFrame({"主体":"","平台":"","店铺":"","2C总金额":"","发出总金额":""})
        return pd.DataFrame(data=None, columns=["主体", "平台", "店铺", "2C总金额", "发出总金额"])


def compress_attaches(files, out_name):
    f = zipfile.ZipFile(out_name, 'w', zipfile.ZIP_DEFLATED)
    for file in files:
        f.write(file, file.split(os.sep)[-1])
    f.close()


def check_shop(df_fn, df_it):
    # company
    # df_fn = df_fn[df_fn["主体"].str.contains(company, na=False)]

    df_it["店铺2"] = df_it["店铺"].apply(lambda x: str.upper(x).strip())
    df_fn["财务店铺名称2"] = df_fn["财务店铺名称"].apply(lambda x: str.upper(x).strip())

    print("测试it")
    print(df_it.to_markdown())

    print("测试fn")
    print(df_fn.to_markdown())

    df_fn = df_fn.merge(df_it, how="left", left_on=["主体", "平台", "财务店铺名称2"], right_on=["主体", "平台", "店铺2"])
    df_fn.fillna(0, inplace=True)

    df_fn["减20发出商品后"] = df_fn["减20发出商品后"].astype("float64")
    df_fn["加2021年发出商品"] = df_fn["加2021年发出商品"].astype("float64")
    df_fn["2C总金额"] = df_fn["2C总金额"].astype("float64")
    df_fn["发出总金额"] = df_fn["发出总金额"].astype("float64")

    print("debug:")
    print(df_fn.to_markdown())

    df_fn["2C差异"] = df_fn.apply(
        lambda x: "OK" if abs(x["减20发出商品后"] - x["2C总金额"]) < 0.01 else x["减20发出商品后"] - x["2C总金额"], axis=1)
    df_fn["发出商品差异"] = df_fn.apply(
        lambda x: "OK" if abs(x["加2021年发出商品"] - x["发出总金额"]) < 0.01 else x["加2021年发出商品"] - x["发出总金额"], axis=1)
    del df_fn["店铺"]

    print("检查IT")
    print(df_it.to_markdown())

    print("结果")
    print(df_fn.to_markdown())

    df_fn["2C差异"] = df_fn["2C差异"].astype(str)
    df_fn["发出商品差异"] = df_fn["发出商品差异"].astype(str)

    kouchu_2C = []
    kouchu_Fachu = []

    zengjia_2C = []
    zengjia_Fachu = []

    company_list = df_fn["主体"].unique()
    # files=[]
    for company in company_list:
        files = [output_dir_2C + os.sep + "2C{}.xlsx".format(company),
                 output_dir_2C发出商品 + os.sep + "2C发出商品{}.xlsx".format(company),
                 output_dir_2C + os.sep + "2C{}_采购订单.xlsx".format(company),
                 output_dir_2C发出商品 + os.sep + "2C发出商品{}_采购订单.xlsx".format(company)]

        # 在上级目录生成压缩包文件
        compress_attaches(files, output_dir_2C + os.sep + '..' + os.sep + company + '_摊销后底稿.zip')
        # Zip_files(files,output_dir_2C + os.sep+'..'+os.sep+company+'.zip')

    for index, row in df_fn.iterrows():
        if row["2C差异"] != "OK":
            if row["2C差异"].find("-") >= 0:
                company = row["主体"]
                pingtai = row["平台"]
                shop = row["财务店铺名称"]
                jine = round(float("".join(row["2C差异"]).replace("-", "")), 2)

                print(company, pingtai, shop, jine)

                df_2c = pd.read_excel(output_dir_2C + os.sep + "2C{}.xlsx".format(company))
                # df_2c=pd.read_excel( r"C:\Users\ns2033\Downloads\摊销\需处理的\2C\2C{}.xlsx".format(company)   )
                df_2c = df_2c.groupby(["主体", "平台", "店铺", "商家编码", "平均价格"])["总金额"].sum().reset_index()
                df_2c.columns = ["主体", "平台", "店铺", "商家编码", "平均价格", "总金额"]
                df_2c["店铺"] = df_2c["店铺"].astype(str)
                df_2c["店铺2"] = df_2c["店铺"].apply(lambda x: str.upper(x))
                print("抽查畅销产品")
                # &(df_2c["平均价格"]<=jine)
                df_2c_2 = df_2c[
                    ((df_2c["平台"].str.contains(pingtai)) & (df_2c["店铺2"].str.contains(str.upper(shop))))].sort_values(
                    by=["总金额"], ascending=False)
                if df_2c_2.shape[0] > 0:
                    print(df_2c_2.head(3).to_markdown())
                    sku = df_2c_2.iloc[0]["商家编码"]
                    shop = df_2c_2.iloc[0]["店铺"]
                    print("找到产品了:", sku)
                    # tiaozheng.append(["'{}'".format(pingtai),"'{}'".format(shop),"'{}'".format(sku),jine])
                    kouchu_2C.append([pingtai, shop, sku, jine])
                else:
                    print("没有找到该店铺商品")
            else:
                company = row["主体"]
                pingtai = row["平台"]
                shop = row["财务店铺名称"]
                jine = round(float("".join(row["2C差异"])), 2)

                print(company, pingtai, shop, jine)
                zengjia_2C.append([pingtai, shop, jine])

        if row["发出商品差异"] != "OK":
            if row["发出商品差异"].find("-") >= 0:
                company = row["主体"]
                pingtai = row["平台"]
                shop = row["财务店铺名称"]
                jine = round(float("".join(row["发出商品差异"]).replace("-", "")), 2)

                print(company, pingtai, shop, jine)

                if os.path.exists(output_dir_2C发出商品 + os.sep + "2C发出商品{}.xlsx".format(company)):
                    df_2c = pd.read_excel(output_dir_2C发出商品 + os.sep + "2C发出商品{}.xlsx".format(company))
                    # df_2c=pd.read_excel( r"C:\Users\ns2033\Downloads\摊销\需处理的\2C\2C{}.xlsx".format(company)   )
                    df_2c = df_2c.groupby(["主体", "平台", "店铺", "商家编码", "平均价格"])["总金额"].sum().reset_index()
                    df_2c.columns = ["主体", "平台", "店铺", "商家编码", "平均价格", "总金额"]
                    df_2c["店铺"] = df_2c["店铺"].astype(str)
                    df_2c["店铺2"] = df_2c["店铺"].apply(lambda x: str.upper(x))
                    print("抽查畅销产品")
                    # &(df_2c["平均价格"]<=jine)
                    df_2c_2 = df_2c[((df_2c["平台"].str.contains(pingtai)) & (
                        df_2c["店铺2"].str.contains(str.upper(shop))))].sort_values(by=["总金额"], ascending=False)
                    if df_2c_2.shape[0] > 0:
                        print(df_2c_2.head(3).to_markdown())
                        sku = df_2c_2.iloc[0]["商家编码"]
                        shop = df_2c_2.iloc[0]["店铺"]
                        print("找到产品了:", sku)
                        # tiaozheng.append(["'{}'".format(pingtai),"'{}'".format(shop),"'{}'".format(sku),jine])
                        kouchu_Fachu.append([pingtai, shop, sku, jine])
                    else:
                        print("没有找到该店铺商品")
            else:
                company = row["主体"]
                pingtai = row["平台"]
                shop = row["财务店铺名称"]
                jine = round(float("".join(row["发出商品差异"])), 2)

                print(company, pingtai, shop, jine)
                zengjia_Fachu.append([pingtai, shop, jine])

    print("调账建议：")

    print("2C_扣减建议：")
    print(kouchu_2C)

    print("2C_增加建议：")
    print(zengjia_2C)

    print("2C发出商品_扣减建议：")
    print(kouchu_Fachu)

    print("2C发出商品_增加建议：")
    print(zengjia_Fachu)

    del df_fn["财务店铺名称2"]
    del df_fn["店铺2"]

    df_fn.to_excel(output_dir + r"\摊销比对结果_按主体+店铺.xlsx")

    print("检查完毕!")

    return df_fn


def check_2C(company):
    filename = fn_file
    df_fn = pd.read_excel(filename, sheet_name="2021年未税收入和成本汇总（含发出商品）")
    print(df_fn.head(10).to_markdown())
    df_fn = df_fn[df_fn["主体"].str.contains(company, na=False)][
        ["主体", "减发出商品后含税", "减发出商品后未税收入合计", "加2021年订单2022年回款含税", "加2021年订单2022年回款未税", "加2021年订单2022年回款成本", "2021年线上含税",
         "2021年线上未税", "2021年线上成本"]]
    print("财务结果")
    print(df_fn.head(10).to_markdown())

    df_2c = pd.read_excel(output_dir_2C + os.sep + "2C{}.xlsx".format(company))
    # df_2c=pd.read_excel( r"C:\Users\ns2033\Downloads\摊销\需处理的\2C\2C{}.xlsx".format(company)   )
    ToC_money = df_2c["总金额"].sum()
    ToC_notax = df_2c["未税金额"].sum()
    ToC_pur_havtax = df_2c["采购未税金额"].sum()
    ToC_pur_notax = df_2c["采购含税金额"].sum()

    if os.path.exists(output_dir_2C发出商品 + os.sep + "2C发出商品{}.xlsx".format(company)):
        df_fachu = pd.read_excel(output_dir_2C发出商品 + os.sep + "2C发出商品{}.xlsx".format(company))
        # df_fachu=pd.read_excel(r"C:\Users\ns2033\Downloads\摊销\自动处理结果\需处理的\2C发出商品{}.xlsx".format(company) )
        Fachu_money = df_fachu["总金额"].sum()
        Fachu_notax = df_fachu["未税金额"].sum()
        Fachu_pur_havtax = df_fachu["采购未税金额"].sum()
        Fachu_pur_notax = df_fachu["采购含税金额"].sum()
    else:
        Fachu_money = 0
        Fachu_notax = 0
        Fachu_pur_havtax = 0
        Fachu_pur_notax = 0

    print("减发出商品后含税:",
          "检查合格" if abs(df_fn.iloc[0]["减发出商品后含税"] - ToC_money) < 0.01 else "{}{}{}".format(df_fn.iloc[0]["减发出商品后含税"],
                                                                                           "<>", ToC_money))
    print("减发出商品后未税收入合计:", "检查合格" if abs(df_fn["减发出商品后未税收入合计"].sum() - ToC_notax) < 0.01 else "{}{}{}".format(
        df_fn.iloc[0]["减发出商品后未税收入合计"], "<>", ToC_notax))
    print("加2021年订单2022年回款未税:",
          "检查合格" if abs(df_fn["加2021年订单2022年回款未税"].sum() - Fachu_notax) < 0.01 else "{}{}{}".format(
              df_fn.iloc[0]["加2021年订单2022年回款未税"], "<>", Fachu_notax))
    print("加2021年订单2022年回款成本:",
          "检查合格" if abs(df_fn["加2021年订单2022年回款成本"].sum() - Fachu_pur_havtax) < 0.01 else "{}{}{}".format(
              df_fn.iloc[0]["加2021年订单2022年回款成本"], "<>", Fachu_pur_havtax))
    print("2021年线上含税:", "检查合格" if abs(df_fn["2021年线上含税"].sum() - (ToC_money + Fachu_money)) < 0.01 else "{}{}{}".format(
        df_fn.iloc[0]["2021年线上含税"], "<>", ToC_money + Fachu_money))
    print("2021年线上未税:", "检查合格" if abs(df_fn["2021年线上未税"].sum() - (ToC_notax + Fachu_notax)) < 0.01 else "{}{}{}".format(
        df_fn.iloc[0]["2021年线上未税"], "<>", ToC_notax + Fachu_notax))
    print("2021年线上成本:",
          "检查合格" if abs(df_fn["2021年线上成本"].sum() - (ToC_pur_notax + Fachu_pur_notax)) < 0.01 else "{}{}{}".format(
              df_fn.iloc[0]["2021年线上成本"], "<>", ToC_pur_notax + Fachu_pur_notax))

    dict = {'减发出商品后含税': "检查合格" if abs(df_fn.iloc[0]["减发出商品后含税"] - ToC_money) < 0.01 else df_fn.iloc[0][
                                                                                             "减发出商品后含税"] - ToC_money,
            '减发出商品后未税收入合计': "检查合格" if abs(df_fn["减发出商品后未税收入合计"].sum() - ToC_notax) < 0.01 else df_fn[
                                                                                                   "减发出商品后未税收入合计"].sum() - ToC_notax,
            '加2021年订单2022年回款未税': "检查合格" if abs(df_fn["加2021年订单2022年回款未税"].sum() - Fachu_notax) < 0.01 else df_fn[
                                                                                                               "加2021年订单2022年回款未税"].sum() - Fachu_notax,
            '加2021年订单2022年回款成本': "检查合格" if abs(df_fn["加2021年订单2022年回款成本"].sum() - Fachu_pur_havtax) < 0.01 else df_fn[
                                                                                                                    "加2021年订单2022年回款成本"].sum() - Fachu_pur_havtax,
            '2021年线上含税': "检查合格" if abs(df_fn["2021年线上含税"].sum() - (ToC_money + Fachu_money)) < 0.01 else df_fn[
                                                                                                             "2021年线上含税"].sum() - (
                                                                                                                     ToC_money + Fachu_money),
            '2021年线上未税': "检查合格" if abs(df_fn["2021年线上未税"].sum() - (ToC_notax + Fachu_notax)) < 0.01 else df_fn[
                                                                                                             "2021年线上未税"].sum() - (
                                                                                                                     ToC_notax + Fachu_notax),
            '2021年线上成本': "检查合格" if abs(df_fn["2021年线上未税"].sum() - (ToC_pur_notax + Fachu_pur_notax)) < 0.01 else df_fn[
                                                                                                                     "2021年线上成本"].sum() - (
                                                                                                                             ToC_pur_notax + Fachu_pur_notax)
            }

    print(dict)

    # df_2c=df_2c.transform({"总金额":np.sum,"未税金额":np.sum,"含税金额":np.sum,"采购未税金额":np.sum})
    # df_2c["saletype"]="2C"
    # df_fachu=df_fachu.transform({"总金额":np.sum,"未税金额":np.sum,"采购含税金额":np.sum,"含税金额":np.sum,"采购未税金额":np.sum})
    # df_fachu["saletype"] = "发出商品"
    # df_sum=pd.concat(df_2c,df_fachu)

    print("检查完毕!")


def cal_sum_by_company():
    # 读取所有excel中的汇总数
    df = read_company("芭葆兔（深圳）日用品有限公司")
    df = df.append(read_company("一片珍芯（深圳）化妆品有限公司"))
    df = df.append(read_company("不是很酷（深圳）服装有限公司"))
    df = df.append(read_company("不酷（深圳）商贸有限公司"))
    df = df.append(read_company("冰川女神（深圳）化妆品有限公司"))
    df = df.append(read_company("勃狄（深圳）化妆品有限公司"))
    df = df.append(read_company("可瘾（广州）化妆品有限公司"))
    df = df.append(read_company("可隐（深圳）化妆品有限公司"))
    # df = df.append(read_company("吉维儿（深圳）化妆品有限公司"))
    df = df.append(read_company("商魂信息（深圳）有限公司"))
    df = df.append(read_company("喝啥（深圳）食品有限公司"))
    df = df.append(read_company("多瑞（深圳）日用品有限公司"))
    df = df.append(read_company("妈咪港湾（深圳）化妆品有限公司"))
    df = df.append(read_company("宅星人（深圳）食品有限公司"))
    df = df.append(read_company("宝贝港湾（深圳）化妆品有限公司"))
    df = df.append(read_company("宝贝配方师（深圳）日用品有限公司"))
    df = df.append(read_company("宝贝魔术师（深圳）日用品有限公司"))
    df = df.append(read_company("尚隐（深圳）计生用品有限公司"))
    df = df.append(read_company("平湖宏炽贸易有限公司"))
    df = df.append(read_company("平湖鑫桂贸易有限公司"))
    df = df.append(read_company("平湖鲁文国际贸易有限公司"))
    df = df.append(read_company("广州市尚西国际贸易有限公司"))
    df = df.append(read_company("广州斌闻贸易有限公司"))
    df = df.append(read_company("广州驰骄服饰有限公司"))
    df = df.append(read_company("惠优购（深圳）日用品有限公司"))
    df = df.append(read_company("戏酱（深圳）食品有限公司"))
    df = df.append(read_company("控师（深圳）化妆品有限公司"))
    df = df.append(read_company("播地艾（广州）化妆品有限公司"))
    df = df.append(read_company("无极爽（深圳）日用品有限公司"))
    df = df.append(read_company("末隐师（广州）化妆品有限公司"))
    df = df.append(read_company("柚选（深圳）化妆品有限公司"))
    df = df.append(read_company("植之璨（深圳）化妆品有限公司"))
    df = df.append(read_company("樱语（深圳）日用品有限公司"))

    df = df.append(read_company("橙意满满（深圳）化妆品有限公司"))
    df = df.append(read_company("泉州市航星贸易有限公司"))
    df = df.append(read_company("泡研（深圳）化妆品有限公司"))
    df = df.append(read_company("浙江魔湾电子有限公司"))
    df = df.append(read_company("深圳大前海物流有限公司"))
    df = df.append(read_company("深圳市二十四小时七天商贸有限公司"))
    df = df.append(read_company("深圳市卖家优选实业有限公司"))
    df = df.append(read_company("深圳市卖家联合商贸有限公司"))
    df = df.append(read_company("深圳市博滴日用品有限公司"))
    df = df.append(read_company("深圳市白皮书文化传媒有限公司"))
    df = df.append(read_company("深圳市精酿商贸有限公司"))
    df = df.append(read_company("深圳市艾法商贸有限公司"))
    df = df.append(read_company("深圳市配颜师生物科技有限公司"))
    df = df.append(read_company("深圳市魔湾游戏科技有限公司"))
    df = df.append(read_company("深圳市麦凯莱科技有限公司"))
    df = df.append(read_company("深圳樱岚护肤品有限公司"))
    df = df.append(read_company("深圳睿旗科技有限公司"))
    df = df.append(read_company("深圳造白化妆品有限公司"))
    df = df.append(read_company("深圳魔湾电子有限公司"))
    df = df.append(read_company("燃威（深圳）食品有限公司"))
    df = df.append(read_company("珍芯漾肤（深圳）化妆品有限公司"))
    df = df.append(read_company("白卿（深圳）化妆品有限公司"))
    df = df.append(read_company("盈养泉（深圳）化妆品有限公司"))
    df = df.append(read_company("秀美颜（广州）化妆品有限公司"))
    df = df.append(read_company("秀美颜（深圳）化妆品有限公司"))
    df = df.append(read_company("肌密泉（深圳）化妆品有限公司"))
    df = df.append(read_company("肌沫（深圳）化妆品有限公司"))
    df = df.append(read_company("肯妮诗（深圳）化妆品有限公司"))
    df = df.append(read_company("芭葆兔（深圳）日用品有限公司"))
    df = df.append(read_company("若蘅（深圳）化妆品有限公司"))
    df = df.append(read_company("茉小桃（深圳）化妆品有限公司"))
    df = df.append(read_company("茱莉珂丝（深圳）化妆品有限公司"))
    df = df.append(read_company("萌洁齿（深圳）日用品有限公司"))
    df = df.append(read_company("萦丝茧（深圳）化妆品有限公司"))
    df = df.append(read_company("补舍（深圳）食品有限公司"))
    df = df.append(read_company("谷口（深圳）化妆品有限公司"))
    df = df.append(read_company("贝贝港湾（深圳）化妆品有限公司"))
    df = df.append(read_company("造味（深圳）食品有限公司"))
    df = df.append(read_company("造白（广州）化妆品有限公司"))
    df = df.append(read_company("配颜师（嘉兴）生物科技有限公司"))
    df = df.append(read_company("配颜师（深圳）化妆品有限公司"))
    df = df.append(read_company("铲喜官（深圳）日用品有限公司"))
    df = df.append(read_company("魔妆（深圳）化妆品有限公司"))

    #
    #
    # print(df.to_markdown())
    df.to_excel(output_dir + r"\摊销后汇总表(按主体).xlsx")
    df_it = df.copy()
    # #
    #
    # return df

    filename = fn_file
    df_fn = pd.read_excel(filename, sheet_name="2021年未税收入和成本汇总（含发出商品）")
    # print(df_fn.head(10).to_markdown())
    df_fn = df_fn[
        ["主体", "减发出商品后含税", "减发出商品后未税收入合计", "加2021年订单2022年回款含税", "加2021年订单2022年回款未税", "加2021年订单2022年回款成本", "2021年线上含税",
         "2021年线上未税", "2021年线上成本"]]
    print("财务结果")
    print(df_fn.head(10).to_markdown())

    # df_it = pd.read_excel(r"C:\Users\ns2033\Downloads\摊销后汇总表(按主体).xlsx")
    print(df_it.to_markdown())
    df = df_fn.merge(df_it, how="left", on=["主体"])

    df["减发出商品后含税"] = df["减发出商品后含税"].astype("float64")
    df["减发出商品后未税收入合计"] = df["减发出商品后未税收入合计"].astype("float64")
    df["加2021年订单2022年回款含税"] = df["加2021年订单2022年回款含税"].astype("float64")
    df["加2021年订单2022年回款未税"] = df["加2021年订单2022年回款未税"].astype("float64")
    df["加2021年订单2022年回款成本"] = df["加2021年订单2022年回款成本"].astype("float64")
    df["2021年线上含税"] = df["2021年线上含税"].astype("float64")
    df["2021年线上未税"] = df["2021年线上未税"].astype("float64")
    df["2021年线上成本"] = df["2021年线上成本"].astype("float64")

    df["ToC_money"] = df["ToC_money"].astype("float64")
    df["ToC_notax"] = df["ToC_notax"].astype("float64")
    df["ToC_pur_havtax"] = df["ToC_pur_havtax"].astype("float64")
    df["ToC_pur_notax"] = df["ToC_pur_notax"].astype("float64")

    df["Fachu_money"] = df["Fachu_money"].astype("float64")
    df["Fachu_notax"] = df["Fachu_notax"].astype("float64")
    df["Fachu_pur_havtax"] = df["Fachu_pur_havtax"].astype("float64")
    df["Fachu_pur_notax"] = df["Fachu_pur_notax"].astype("float64")

    df["减发出商品后含税_check"] = df.apply(
        lambda x: "检查合格" if abs(x["减发出商品后含税"] - x["ToC_money"]) < 0.01 else x["减发出商品后含税"] - x["ToC_money"], axis=1)
    df["减发出商品后未税收入合计_check"] = df.apply(
        lambda x: "检查合格" if abs(x["减发出商品后未税收入合计"] - x["ToC_notax"]) < 0.01 else x["减发出商品后未税收入合计"] - x["ToC_notax"],
        axis=1)
    df["加2021年订单2022年回款未税_check"] = df.apply(
        lambda x: "检查合格" if abs(x["加2021年订单2022年回款未税"] - x["Fachu_notax"]) < 0.01 else x["加2021年订单2022年回款未税"] - x[
            "Fachu_notax"], axis=1)
    df["加2021年订单2022年回款成本_check"] = df.apply(
        lambda x: "检查合格" if abs(x["加2021年订单2022年回款成本"] - x["Fachu_pur_havtax"]) < 0.01 else x["加2021年订单2022年回款成本"] - x[
            "Fachu_pur_havtax"], axis=1)
    df["2021年线上含税_check"] = df.apply(
        lambda x: "检查合格" if abs(x["2021年线上含税"] - (x["ToC_money"] + x["Fachu_money"])) < 0.01 else x["2021年线上含税"] - (
                    x["ToC_money"] + x["Fachu_money"]), axis=1)
    df["2021年线上未税_check"] = df.apply(
        lambda x: "检查合格" if abs(x["2021年线上未税"] - (x["ToC_notax"] + x["Fachu_notax"])) < 0.01 else x["2021年线上未税"] - (
                    x["ToC_notax"] + x["Fachu_notax"]), axis=1)
    df["2021年线上成本_check"] = df.apply(
        lambda x: "检查合格" if abs(x["2021年线上未税"] - (x["ToC_pur_notax"] + x["Fachu_pur_notax"])) < 0.01 else x[
                                                                                                              "2021年线上成本"] - (
                                                                                                                      x[
                                                                                                                          "ToC_pur_notax"] +
                                                                                                                      x[
                                                                                                                          "Fachu_pur_notax"]),
        axis=1)

    df.rename(columns={"ToC_money": "2C总金额"}, inplace=True)
    df.rename(columns={"ToC_notax": "2C未税金额"}, inplace=True)
    df.rename(columns={"ToC_pur_havtax": "2C采购含税金额"}, inplace=True)
    df.rename(columns={"ToC_pur_notax": "2C采购未税金额"}, inplace=True)

    df.rename(columns={"Fachu_money": "2C发出商品总金额"}, inplace=True)
    df.rename(columns={"Fachu_notax": "2C发出商品未税金额"}, inplace=True)
    df.rename(columns={"Fachu_pur_havtax": "2C发出商品采购含税金额"}, inplace=True)
    df.rename(columns={"Fachu_pur_notax": "2C发出商品采购未税金额"}, inplace=True)

    print(df.to_markdown())
    df.to_excel(output_dir + r"\摊销后汇总表(按主体)_核对结果.xlsx")


def check_sum_by_shop(company):
    if company == '':
        df = read_shop("芭葆兔（深圳）日用品有限公司")
        df = df.append(read_shop("一片珍芯（深圳）化妆品有限公司"))
        df = df.append(read_shop("不是很酷（深圳）服装有限公司"))
        df = df.append(read_shop("不酷（深圳）商贸有限公司"))
        df = df.append(read_shop("冰川女神（深圳）化妆品有限公司"))
        df = df.append(read_shop("勃狄（深圳）化妆品有限公司"))
        df = df.append(read_shop("可瘾（广州）化妆品有限公司"))
        df = df.append(read_shop("可隐（深圳）化妆品有限公司"))
        # df = df.append(read_shop("吉维儿（深圳）化妆品有限公司"))
        df = df.append(read_shop("商魂信息（深圳）有限公司"))
        df = df.append(read_shop("喝啥（深圳）食品有限公司"))
        df = df.append(read_shop("多瑞（深圳）日用品有限公司"))
        df = df.append(read_shop("妈咪港湾（深圳）化妆品有限公司"))
        df = df.append(read_shop("宅星人（深圳）食品有限公司"))
        df = df.append(read_shop("宝贝港湾（深圳）化妆品有限公司"))
        df = df.append(read_shop("宝贝配方师（深圳）日用品有限公司"))
        df = df.append(read_shop("宝贝魔术师（深圳）日用品有限公司"))
        df = df.append(read_shop("尚隐（深圳）计生用品有限公司"))
        df = df.append(read_shop("平湖宏炽贸易有限公司"))
        df = df.append(read_shop("平湖鑫桂贸易有限公司"))
        df = df.append(read_shop("平湖鲁文国际贸易有限公司"))
        df = df.append(read_shop("广州市尚西国际贸易有限公司"))
        df = df.append(read_shop("广州斌闻贸易有限公司"))
        df = df.append(read_shop("广州驰骄服饰有限公司"))
        df = df.append(read_shop("惠优购（深圳）日用品有限公司"))
        df = df.append(read_shop("戏酱（深圳）食品有限公司"))
        df = df.append(read_shop("控师（深圳）化妆品有限公司"))
        df = df.append(read_shop("播地艾（广州）化妆品有限公司"))
        df = df.append(read_shop("无极爽（深圳）日用品有限公司"))
        df = df.append(read_shop("末隐师（广州）化妆品有限公司"))
        df = df.append(read_shop("柚选（深圳）化妆品有限公司"))
        df = df.append(read_shop("植之璨（深圳）化妆品有限公司"))
        df = df.append(read_shop("樱语（深圳）日用品有限公司"))

        df = df.append(read_shop("橙意满满（深圳）化妆品有限公司"))
        df = df.append(read_shop("泉州市航星贸易有限公司"))
        df = df.append(read_shop("泡研（深圳）化妆品有限公司"))
        df = df.append(read_shop("浙江魔湾电子有限公司"))
        df = df.append(read_shop("深圳大前海物流有限公司"))
        df = df.append(read_shop("深圳市二十四小时七天商贸有限公司"))
        df = df.append(read_shop("深圳市卖家优选实业有限公司"))
        df = df.append(read_shop("深圳市卖家联合商贸有限公司"))
        df = df.append(read_shop("深圳市博滴日用品有限公司"))
        df = df.append(read_shop("深圳市白皮书文化传媒有限公司"))
        df = df.append(read_shop("深圳市精酿商贸有限公司"))
        df = df.append(read_shop("深圳市艾法商贸有限公司"))
        df = df.append(read_shop("深圳市配颜师生物科技有限公司"))
        df = df.append(read_shop("深圳市魔湾游戏科技有限公司"))
        df = df.append(read_shop("深圳市麦凯莱科技有限公司"))
        df = df.append(read_shop("深圳樱岚护肤品有限公司"))
        df = df.append(read_shop("深圳睿旗科技有限公司"))
        df = df.append(read_shop("深圳造白化妆品有限公司"))
        df = df.append(read_shop("深圳魔湾电子有限公司"))
        df = df.append(read_shop("燃威（深圳）食品有限公司"))
        df = df.append(read_shop("珍芯漾肤（深圳）化妆品有限公司"))
        df = df.append(read_shop("白卿（深圳）化妆品有限公司"))
        df = df.append(read_shop("盈养泉（深圳）化妆品有限公司"))
        df = df.append(read_shop("秀美颜（广州）化妆品有限公司"))
        df = df.append(read_shop("秀美颜（深圳）化妆品有限公司"))
        df = df.append(read_shop("肌密泉（深圳）化妆品有限公司"))
        df = df.append(read_shop("肌沫（深圳）化妆品有限公司"))
        df = df.append(read_shop("肯妮诗（深圳）化妆品有限公司"))
        df = df.append(read_shop("芭葆兔（深圳）日用品有限公司"))
        df = df.append(read_shop("若蘅（深圳）化妆品有限公司"))
        df = df.append(read_shop("茉小桃（深圳）化妆品有限公司"))
        df = df.append(read_shop("茱莉珂丝（深圳）化妆品有限公司"))
        df = df.append(read_shop("萌洁齿（深圳）日用品有限公司"))
        df = df.append(read_shop("萦丝茧（深圳）化妆品有限公司"))
        df = df.append(read_shop("补舍（深圳）食品有限公司"))
        df = df.append(read_shop("谷口（深圳）化妆品有限公司"))
        df = df.append(read_shop("贝贝港湾（深圳）化妆品有限公司"))
        df = df.append(read_shop("造味（深圳）食品有限公司"))
        df = df.append(read_shop("造白（广州）化妆品有限公司"))
        df = df.append(read_shop("配颜师（嘉兴）生物科技有限公司"))
        df = df.append(read_shop("配颜师（深圳）化妆品有限公司"))
        df = df.append(read_shop("铲喜官（深圳）日用品有限公司"))
        df = df.append(read_shop("魔妆（深圳）化妆品有限公司"))
    else:
        df = read_shop(company)

    if company > '':
        df = df[df["主体"].str.contains(company, na=False)]
        print("查询 df_it ")
        print(df.to_markdown())
        # filename = r"C:\Users\ns2033\Downloads\2021年订单回款-明细表（20220610154403-2021稿）（发出商品20220614093726-2022稿）-2022.6.16提供核对.xlsx"
        filename = fn_file
        df_fn = pd.read_excel(filename, sheet_name="2021年72家公司合并")
        # 主体	平台	财务店铺名称 减20发出商品后  加2021年发出商品
        print(df_fn.head(10).to_markdown())
        if company > '':
            df_fn = df_fn[df_fn["主体"].str.contains(company, na=False)]

        df_fn = df_fn[["主体", "平台", "财务店铺名称", "减20发出商品后", "加2021年发出商品"]]
        # df_fn["财务店铺名称"]=df_fn["财务店铺名称"].apply(lambda x: str.upper(x))

        df_fn.fillna(0, inplace=True)

        print("财务结果")
        print(df_fn.to_markdown())

        check_shop(df_fn, df)
    else:
        pass
        # 第二种方案
        # df=check_shop("芭葆兔（深圳）日用品有限公司",df_fn)
        # # df.append(check_shop("芭葆兔（深圳）日用品有限公司"),df_fn)
        # df.append(check_shop("一片珍芯（深圳）化妆品有限公司",df_fn))
        # df.append(check_shop("不是很酷（深圳）服装有限公司",df_fn) )
        # df.append(check_shop("不酷（深圳）商贸有限公司",df_fn))
        # df.append(check_shop("冰川女神（深圳）化妆品有限公司",df_fn))
        # df.append(check_shop("勃狄（深圳）化妆品有限公司",df_fn))
        # df.append(check_shop("可瘾（广州）化妆品有限公司",df_fn))
        # df.append(check_shop("可隐（深圳）化妆品有限公司",df_fn))
        # df.append(check_shop("吉维儿（深圳）化妆品有限公司",df_fn))
        # df.append(check_shop("商魂信息（深圳）有限公司",df_fn))
        # df.append(check_shop("喝啥（深圳）食品有限公司",df_fn))
        # df.append(check_shop("多瑞（深圳）日用品有限公司",df_fn))
        # df.append(check_shop("妈咪港湾（深圳）化妆品有限公司",df_fn))
        # df.append(check_shop("宅星人（深圳）食品有限公司",df_fn))
        # df.append(check_shop("宝贝港湾（深圳）化妆品有限公司",df_fn))
        # df.append(check_shop("宝贝配方师（深圳）日用品有限公司",df_fn))
        # df.append(check_shop("宝贝魔术师（深圳）日用品有限公司",df_fn))
        # df.append(check_shop("尚隐（深圳）计生用品有限公司",df_fn))
        # df.append(check_shop("平湖宏炽贸易有限公司",df_fn))
        # df.append(check_shop("平湖鑫桂贸易有限公司",df_fn))
        # df.append(check_shop("平湖鲁文国际贸易有限公司",df_fn))
        # df.append(check_shop("广州市尚西国际贸易有限公司",df_fn))
        # df.append(check_shop("广州斌闻贸易有限公司",df_fn))
        # df.append(check_shop("广州驰骄服饰有限公司",df_fn))
        # df.append(check_shop("惠优购（深圳）日用品有限公司",df_fn))
        # df.append(check_shop("戏酱（深圳）食品有限公司",df_fn))
        # df.append(check_shop("控师（深圳）化妆品有限公司",df_fn))
        # df.append(check_shop("播地艾（广州）化妆品有限公司",df_fn))
        # df.append(check_shop("无极爽（深圳）日用品有限公司",df_fn))
        # df.append(check_shop("末隐师（广州）化妆品有限公司",df_fn))
        # df.append(check_shop("柚选（深圳）化妆品有限公司",df_fn))
        # df.append(check_shop("植之璨（深圳）化妆品有限公司",df_fn))
        # df.append(check_shop("樱语（深圳）日用品有限公司",df_fn))
        #
        # df.append(check_shop("橙意满满（深圳）化妆品有限公司",df_fn))
        # df.append(check_shop("泉州市航星贸易有限公司",df_fn))
        # df.append(check_shop("泡研（深圳）化妆品有限公司",df_fn))
        # df.append(check_shop("浙江魔湾电子有限公司",df_fn))
        # df.append(check_shop("深圳大前海物流有限公司",df_fn))
        # df.append(check_shop("深圳市二十四小时七天商贸有限公司",df_fn))
        # df.append(check_shop("深圳市卖家优选实业有限公司",df_fn))
        # df.append(check_shop("深圳市卖家联合商贸有限公司",df_fn))
        # df.append(check_shop("深圳市博滴日用品有限公司",df_fn))
        # df.append(check_shop("深圳市白皮书文化传媒有限公司",df_fn))
        # df.append(check_shop("深圳市精酿商贸有限公司",df_fn))
        # df.append(check_shop("深圳市艾法商贸有限公司",df_fn))
        # df.append(check_shop("深圳市配颜师生物科技有限公司",df_fn))
        # df.append(check_shop("深圳市魔湾游戏科技有限公司",df_fn))
        # df.append(check_shop("深圳市麦凯莱科技有限公司",df_fn))
        # df.append(check_shop("深圳樱岚护肤品有限公司",df_fn))
        # df.append(check_shop("深圳睿旗科技有限公司",df_fn))
        # df.append(check_shop("深圳造白化妆品有限公司",df_fn))
        # df.append(check_shop("深圳魔湾电子有限公司",df_fn))
        # df.append(check_shop("燃威（深圳）食品有限公司",df_fn))
        # df.append(check_shop("珍芯漾肤（深圳）化妆品有限公司",df_fn))
        # df.append(check_shop("白卿（深圳）化妆品有限公司",df_fn))
        # df.append(check_shop("盈养泉（深圳）化妆品有限公司",df_fn))
        # df.append(check_shop("秀美颜（广州）化妆品有限公司",df_fn))
        # df.append(check_shop("秀美颜（深圳）化妆品有限公司",df_fn))
        # df.append(check_shop("肌密泉（深圳）化妆品有限公司",df_fn))
        # df.append(check_shop("肌沫（深圳）化妆品有限公司",df_fn))
        # df.append(check_shop("肯妮诗（深圳）化妆品有限公司",df_fn))
        # df.append(check_shop("芭葆兔（深圳）日用品有限公司",df_fn))
        # df.append(check_shop("若蘅（深圳）化妆品有限公司",df_fn))
        # df.append(check_shop("茉小桃（深圳）化妆品有限公司",df_fn))
        # df.append(check_shop("茱莉珂丝（深圳）化妆品有限公司",df_fn))
        # df.append(check_shop("萌洁齿（深圳）日用品有限公司",df_fn))
        # df.append(check_shop("萦丝茧（深圳）化妆品有限公司",df_fn))
        # df.append(check_shop("补舍（深圳）食品有限公司",df_fn))
        # df.append(check_shop("谷口（深圳）化妆品有限公司",df_fn))
        # df.append(check_shop("贝贝港湾（深圳）化妆品有限公司",df_fn))
        # df.append(check_shop("造味（深圳）食品有限公司",df_fn))
        # df.append(check_shop("造白（广州）化妆品有限公司",df_fn))
        # df.append(check_shop("配颜师（嘉兴）生物科技有限公司",df_fn))
        # df.append(check_shop("配颜师（深圳）化妆品有限公司",df_fn))
        # df.append(check_shop("铲喜官（深圳）日用品有限公司",df_fn))
        # df.append(check_shop("魔妆（深圳）化妆品有限公司",df_fn))

        # df.to_excel( r"C:\Users\ns2033\Downloads\摊销与财务的比对结果.xlsx")


def convert_2C(company, del_fn, add_fn):
    if company > '':
        tanxiao(work_dir_2C + os.sep + "2C{}.xlsx".format(company), "2C", del_fn, add_fn)


def convert_Fachu(company, del_fn, add_fn):
    pass
    # 发出商品
    # #删除
    # del_fn = []
    # # del_fn = [['快手','博滴播地艾专卖店','6973007000950',3.08]]
    # # 增加
    # add_fn=[]
    # add_fn = [['抖音','播地艾个护专营店',3.08]]
    # del_fn =[['抖音', '尚西个护专营店', '4897112700166', 1428.95]]
    if company > '':
        tanxiao(work_dir_2C发出商品 + os.sep + "2C发出商品{}.xlsx".format(company), "2C发出商品", del_fn, add_fn)


# 列出所有文件
def list_all_files(rootdir, filekey_list):
    if len(filekey_list) > 0:
        filekey_list = filekey_list.replace(",", " ")
        filekey = filekey_list.split(" ")
    else:
        filekey = ''

    # print("key=",filekey)
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
                        # print(filename, "===" ,"".join(path.split("\\")[-1:]))
                        if filename.find(key) >= 0:  # 只做文件名的过滤
                            _files.append(path)
                else:
                    _files.append(path)

    # print(_files)
    # 返回一个文件列表 list
    return _files


def convert_sales_tuihuo_2022(df, sale_type, newfile):
    # 源头
    # 主体	主体	2C\2C发出商品	订单日期	订单日期	店铺名		固定	固定			内部编码	单位	实际卖出	平均价格	税率

    df = df[df["销售分类"] == sale_type]
    # df = df[df["店铺"] == shop]

    if df.shape[0] == 0:
        print("没有销售退货，无需转换！")
        return

    # 读取参考表格
    df_shop = pd.read_excel(shop_file)
    df_sku = pd.read_excel(sku_file)
    df_warehouse = pd.read_excel(warehouse_file, sheet_name="Sheet2")

    # df_sku = df_sku.groupby(["条码"]).agg({"内部参考": max, "显示名称": max, "计量单位/显示名称": max}).reset_index()
    df_sku = df_sku.groupby(["条码"]).agg({"内部参考": max, "名称": max, "计量单位/单位": max}).reset_index()
    df_sku.columns = ["条码", "内部参考", "名称", "计量单位/显示名称"]
    df_sku["名称"] = df_sku.apply(lambda x: x["名称"].replace("[{}]".format(x["内部参考"]), "").strip(), axis=1)

    # 原始文件
    # df = pd.read_excel(filename)
    # print("读取文件:",filename)
    # print(df.to_markdown())
    print(df.head(10).to_markdown())

    cnt = df.shape[0]
    print("原表行数:", df.shape[0])

    # sale_type=df["类型"].iloc[0]
    print("销售分类:", sale_type)

    df_sales = df.copy()
    # sale_type = "2C发出商品" if filename.find("2C发出商品") >= 0 else "2C"
    df_sales.rename({"主体": "配送仓库", "店铺": "客户"}, inplace=True)

    # df_product = pd.read_excel(product_file)
    df_sales["商家编码"] = df_sales["商家编码"].astype(str)
    # df_product["条码"] = df_product["条码"].astype(str)

    # df_sku = df_sku.groupby(["条码"]).agg(名称=("显示名称", "max")).reset_index()
    # df_sku.columns = ["条码", "名称"]

    df_sales = df_sales.merge(df_sku[["条码", "内部参考", "名称"]], how="left", left_on=["商家编码"], right_on=["条码"])

    if (df_sales[df_sales["名称"].isnull()].shape[0] > 0):
        df_sales[df_sales["名称"].isnull()].to_excel(newfile.replace(".xls", "_销售条码异常.xls"))
    else:
        print("条码检查合格")

    # del df_sales["条码"]
    print("抽查数据")
    # print(df_sales.head(5).to_markdown())

    df_sales["进销存标识"] = sale_type
    df_sales["单据日期"] = df_sales["发货时间"]
    # df_sales["价格表"] = ""
    # df_sales["跟单员"] = "陆俊秀"
    # df_sales["业务团队"] = "潘勤"
    df_sales["源单据"] = ""
    df_sales["客户参考"] = ""
    df_sales["退货负责人"] = ""
    df_sales["退货订单行/说明"] = df_sales.apply(lambda x: "[{}]{}".format(x["内部参考"], x["商品名称"]), axis=1)

    df_sales["店铺2"] = df_sales["店铺"].apply(lambda x: str.upper(x))
    # df_shop["OMS店铺名称2"]=df_shop["OMS店铺名称"].apply(lambda x: str.upper(x))
    df_shop["财务店铺名称2"] = df_shop["财务店铺名称"].apply(lambda x: str.upper(x))

    df_sales = df_sales.merge(df_shop[["平台", "财务店铺名称2", "Odoo店铺名称"]], how="left", left_on=["平台", "店铺2"],
                              right_on=["平台", "财务店铺名称2"])
    print("转换1", "成功" if cnt == df.shape[0] else "失败:{}-{}".format(cnt, df_sales.shape[0]))

    print("sku_x")
    print(df_sales.head(10).to_markdown())
    df_sales = df_sales.merge(df_sku[["条码", "计量单位/显示名称"]], how="left", left_on=["条码"],
                              right_on=["条码"])
    print("转换2", "成功" if cnt == df.shape[0] else "失败:{}-{}".format(cnt, df_sales.shape[0]))

    print("sku_y")
    print(df_sales.head(10).to_markdown())
    print(df_warehouse.head(10).to_markdown())

    # df_warehouse.rename(columns={"公司": "仓库公司"}, inplace=True)
    df_sales = df_sales.merge(df_warehouse[["公司", "仓库/显示名称"]], how="left", left_on=["主体"],
                              right_on=["公司"])
    # df_sales = df_sales.merge(df_warehouse[["仓库公司", "仓库/显示名称"]], how="left", left_on=["主体"], right_on=["仓库公司"])

    # 重命名
    # "商家条码":"sku",
    df_sales.rename(
        columns={"Odoo店铺名称": "客户", "内部参考": "退货订单行/产品", "计量单位/显示名称": "退货订单行/计量单位", "实际卖出": "退货订单行/预退数量",
                 "平均价格": "退货订单行/单价",
                 "仓库/显示名称": "仓库"},
        inplace=True)

    print("缺少客户名称的记录:")
    print(df_sales[df_sales["客户"].isnull()].to_markdown())

    print("缺少配送仓库的记录:")
    print(df_sales[df_sales["仓库"].isnull()].to_markdown())

    # 订单行/产品
    # df_sales=df_sales.merge(df_sku[["条码","内部参考"]],how="left",left_on=["商家编码"],right_on=["条码"])

    # print(df_sales[df_sales["税率"].isnull()])
    # df_sales["税率"].fillna(99, inplace=True)

    df_sales["退货订单行/税率"] = df_sales["税率"].apply(lambda x: "税收{}%（含）".format(int(x * 100)))
    df_sales["退货单号"] = ""
    df_sales["note"] = ""
    df_sales["公司"] = df_sales["主体"]
    # df_sales["配送仓库"] = df_warehouse["仓库/显示名称"]
    df_sales["退货负责人"] = ""
    # df_sales["配送仓库"] = df_sales['仓库/显示名称']

    print("结果行数:", df_sales.shape[0])

    print("转换 ", "成功" if cnt == df_sales.shape[0] else "失败")

    print("查看结果:")
    print(df_sales.head(10).to_markdown())
    df_sales = df_sales[
        ["进销存标识", "公司", "仓库", "退货单号", "单据日期", "客户", "note", "条码", "退货订单行/产品", "退货订单行/说明", "退货订单行/计量单位",
         "退货订单行/预退数量", "退货订单行/单价", "退货订单行/税率", "退货负责人"]]

    # mubiao = r"D:\work\OMS理帐\2C发出商品"
    df_sales.to_excel(newfile, index=False)

    print("报告:")
    print("缺少客户名称的记录:", df_sales[df_sales["客户"].isnull()].shape[0])
    print("缺少配送仓库的记录:", df_sales[df_sales["仓库"].isnull()].shape[0])

    return df_sales

    # 目标
    #  公司	配送仓库	进销存标识	订单日期	承诺日期	客户	价格表	跟单员	业务团队	源单据	客户参考	订单行/产品	订单行/计量单位	订单行/订购数量	订单行/单价	订单行/税率
    # 	公司	仓库	退货单号	单据日期	客户	note	条码	退货订单行/产品	退货订单行/说明	退货订单行/计量单位	退货订单行/预退数量	退货订单行/单价	退货订单行/税率	退货负责人


def convert_sales_2022(df, sale_type, newfile):
    # 源头
    # 主体	主体	2C\2C发出商品	订单日期	订单日期	店铺名		固定	固定			内部编码	单位	实际卖出	平均价格	税率

    loss_shop = []

    # 2C 和 2C发出商品 分开
    df = df[df["销售分类"] == sale_type]
    # df = df[df["店铺"] == shop]

    # 读取参考表格
    df_shop = pd.read_excel(shop_file)
    df_sku = pd.read_excel(sku_file)
    df_warehouse = pd.read_excel(warehouse_file, sheet_name="Sheet2")

    # df_sku = df_sku.groupby(["条码"]).agg({"内部参考": max, "显示名称": max, "计量单位/显示名称": max}).reset_index()
    df_sku = df_sku.groupby(["内部参考"]).agg({"条码": max, "名称": max, "计量单位/单位": max}).reset_index()
    df_sku.columns = ["内部参考", "条码", "名称", "计量单位/显示名称"]
    df_sku["名称"] = df_sku.apply(lambda x: x["名称"].replace("[{}]".format(x["内部参考"]), "").strip(), axis=1)

    # 原始文件
    # df = pd.read_excel(filename)
    # print("读取文件:",filename)
    # print(df.to_markdown())
    print(df.head(10).to_markdown())

    cnt = df.shape[0]
    print("原表行数:", df.shape[0])

    # sale_type=df["类型"].iloc[0]
    print("销售分类:", sale_type)

    df_sales = df.copy()
    # sale_type = "2C发出商品" if filename.find("2C发出商品") >= 0 else "2C"
    df_sales.rename({"主体": "配送仓库", "店铺": "客户"}, inplace=True)

    # df_product = pd.read_excel(product_file)
    df_sales["商家编码"] = df_sales["商家编码"].astype(str)
    # df_product["条码"] = df_product["条码"].astype(str)

    # df_sku = df_sku.groupby(["条码"]).agg(名称=("显示名称", "max")).reset_index()
    # df_sku.columns = ["条码", "名称"]

    df_sales = df_sales.merge(df_sku[["条码", "名称", "计量单位/显示名称"]], how="left", left_on=["商家编码"], right_on=["条码"])

    if (df_sales[df_sales["名称"].isnull()].shape[0] > 0):
        df_sales[df_sales["名称"].isnull()].to_excel(newfile.replace(".xls", "_销售条码异常.xls"))
    else:
        print("条码检查合格")

    # 先按照条码匹配，然后按照内部参考号匹配
    del df_sales["条码"]
    if df_sales.shape[0] > 0:
        df_sales = df_sales.merge(df_sku[["内部参考", "名称", "计量单位/显示名称"]], how="left", left_on=["商家编码"], right_on=["内部参考"])
        df_sales["名称_x"] = df_sales.apply(lambda x: x["名称_y"] if len(str(x["名称_x"])) <= 0 else x["名称_x"], axis=1)
        del df_sales["名称_y"]
        df_sales.rename(columns={"名称_x": "名称"}, inplace=True)

    if df_sales.shape[0] > 0:
        df_sales["计量单位/显示名称_x"] = df_sales.apply(
            lambda x: x["计量单位/显示名称_y"] if len(str(x["计量单位/显示名称_x"])) <= 0 else x["计量单位/显示名称_x"], axis=1)
        del df_sales["计量单位/显示名称_y"]
        df_sales.rename(columns={"计量单位/显示名称_x": "计量单位/显示名称"}, inplace=True)

    print("抽查销售订单：")
    print(df_sales.head(5).to_markdown())

    df_sales["进销存标识"] = sale_type
    df_sales["承诺日期"] = df_sales["发货时间"]
    df_sales["订单日期"] = df_sales["发货时间"]
    df_sales["价格表"] = ""
    df_sales["跟单员"] = "陆俊秀"
    df_sales["业务团队"] = "潘勤"
    df_sales["源单据"] = ""
    df_sales["客户参考"] = ""

    # df_sales.rename(columns={"店铺名称":"店铺","平台名称":"平台","账单公司主体":"主体","实际卖出数量":"实际卖出"},inplace=True)

    df_sales["店铺2"] = df_sales["店铺"].apply(lambda x: str.upper(x))

    # if "店铺" in df_sales.columns:
    #     df_sales["店铺2"]=df_sales["店铺"].apply(lambda x: str.upper(x))
    # elif    "店铺名称" in df_sales.columns:
    #     df_sales["店铺2"]=df_sales["店铺名称"].apply(lambda x: str.upper(x))

    # df_shop["OMS店铺名称2"]=df_shop["OMS店铺名称"].apply(lambda x: str.upper(x))
    df_shop["财务店铺名称2"] = df_shop["财务店铺名称"].apply(lambda x: str.upper(x))

    df_sales = df_sales.merge(df_shop[["平台", "财务店铺名称2", "Odoo店铺名称"]], how="left", left_on=["平台", "店铺2"],
                              right_on=["平台", "财务店铺名称2"])
    print("转换1", "成功" if cnt == df.shape[0] else "失败:{}-{}".format(cnt, df_sales.shape[0]))

    # print("sku_x")
    # print(df_sales.head(10).to_markdown())
    # df_sales = df_sales.merge(df_sku[["条码", "内部参考", "计量单位/显示名称"]], how="left", left_on=["商家编码"],
    #                           right_on=["条码"])
    # print("转换2", "成功" if cnt == df.shape[0] else "失败:{}-{}".format(cnt, df_sales.shape[0]))
    #
    # print("sku_y")
    # print(df_sales.head(10).to_markdown())
    print(df_warehouse.head(10).to_markdown())

    # df_warehouse.rename(columns={"公司": "仓库公司"}, inplace=True)
    df_sales = df_sales.merge(df_warehouse[["公司", "仓库/显示名称"]], how="left", left_on=["主体"],
                              right_on=["公司"])
    # df_sales = df_sales.merge(df_warehouse[["仓库公司", "仓库/显示名称"]], how="left", left_on=["主体"], right_on=["仓库公司"])

    print("查看为什么漏字段：")
    print(df_sales.head(5).to_markdown())
    # df_sales["订单行/产品"]=df_sales["内部参考"]
    # 重命名
    df_sales.rename(
        columns={"Odoo店铺名称": "客户", "内部参考": "订单行/产品", "内部参考号": "订单行/产品", "计量单位/显示名称": "订单行/计量单位", "实际卖出": "订单行/订购数量",
                 "平均价格": "订单行/单价",
                 "仓库/显示名称": "配送仓库"},
        inplace=True)

    if df_sales[df_sales["客户"].isnull()].shape[0] > 0:
        print("缺少客户名称的记录:")
        print(df_sales[df_sales["客户"].isnull()].to_markdown())

    if df_sales[df_sales["配送仓库"].isnull()].shape[0] > 0:
        print("缺少配送仓库的记录:")
        print(df_sales[df_sales["配送仓库"].isnull()].to_markdown())

    # 订单行/产品
    # df_sales=df_sales.merge(df_sku[["条码","内部参考"]],how="left",left_on=["商家编码"],right_on=["条码"])

    # print(df_sales[df_sales["税率"].isnull()])
    # df_sales["税率"].fillna(99, inplace=True)

    df_sales["订单行/税率"] = df_sales["税率"].apply(lambda x: "税收{}%（含）".format(int(x * 100)))
    df_sales["订单行/关税税额"] = ""
    df_sales["订单行/报关单号"] = ""
    df_sales["订单行/汇率"] = ""
    df_sales["公司"] = df_sales["主体"]


    print("结果行数:", df_sales.shape[0])

    print("转换 ", "成功" if cnt == df_sales.shape[0] else "失败")

    print("查看结果:")
    print(df_sales.head(10).to_markdown())
    df_sales = df_sales[
        ["公司", "配送仓库", "进销存标识", "订单日期", "承诺日期", "客户", "价格表", "跟单员", "业务团队", "源单据", "客户参考", "订单行/产品", "订单行/计量单位",
         "订单行/订购数量", "订单行/单价", "订单行/税率", "订单行/关税税额", "订单行/报关单号", "订单行/汇率"]]

    # mubiao = r"D:\work\OMS理帐\2C发出商品"
    df_sales.to_excel(newfile, index=False)

    print("报告:")
    print("缺少客户名称的记录:", df_sales[df_sales["客户"].isnull()].shape[0])
    print("缺少配送仓库的记录:", df_sales[df_sales["配送仓库"].isnull()].shape[0])

    return df_sales

    # 目标
    #  公司	配送仓库	进销存标识	订单日期	承诺日期	客户	价格表	跟单员	业务团队	源单据	客户参考	订单行/产品	订单行/计量单位	订单行/订购数量	订单行/单价	订单行/税率


def convert_purchase(filename, targe_dir, newfile):
    # 源头
    # 主体	主体	2C\2C发出商品	订单日期	订单日期	店铺名		固定	固定			内部编码	单位	实际卖出	平均价格	税率

    # 读取参考表格

    df_shop = pd.read_excel(shop_file)
    df_sku = pd.read_excel(sku_file)
    df_warehouse = pd.read_excel(warehouse_file)

    # df_sku = df_sku.groupby(["条码"]).agg({"内部参考": max, "显示名称": max, "计量单位/显示名称": max}).reset_index()
    df_sku = df_sku.groupby(["条码"]).agg({"内部参考": max, "名称": max, "计量单位/单位": max}).reset_index()
    df_sku.columns = ["条码", "内部参考", "名称", "计量单位"]
    df_sku["条码"] = df_sku["条码"].astype(str)
    df_sku["内部参考"] = df_sku["内部参考"].astype(str)
    df_sku["名称"] = df_sku.apply(lambda x: x["名称"].replace("[{}]".format(x["内部参考"]), "").strip(), axis=1)
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

    # df_product = pd.read_excel(product_file)
    df_purcharse["商家编码"] = df_purcharse["商家编码"].astype(str)
    # df_product["条码"] = df_product["条码"].astype(str)

    # df_product = df_product.groupby(["条码"]).agg(名称=("名称", "max")).reset_index()
    # df_product.columns = ["条码", "名称"]

    df_purcharse = df_purcharse.merge(df_sku[["条码", "名称"]], how="left", left_on=["商家编码"], right_on=["条码"])

    if df_purcharse[df_purcharse["条码"].isnull()].shape[0] > 0:
        print("条码为空")
        print(df_purcharse[df_purcharse["条码"].isnull()].head(10).to_markdown())

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
                                      left_on=["商家编码"],
                                      right_on=["条码"])
    print("转换2", "成功" if cnt == df.shape[0] else "失败:{}-{}".format(cnt, df_purcharse.shape[0]))

    # 重命名
    df_purcharse.rename(
        columns={"内部参考": "订单行/产品", "计量单位/显示名称": "订单行/计量单位", "实际卖出数量": "订单行/订购数量"},
        inplace=True)

    print("缺少 订单行/产品 的记录:")
    print(df_purcharse[df_purcharse["订单行/产品"].isnull()].head(10).to_markdown())

    print("缺少 交货到/数据库 ID 的记录:")
    print(df_purcharse[df_purcharse["交货到/数据库 ID"].isnull()].head(10).to_markdown())

    # 订单行/产品
    df_purcharse["订单行/税率"] = df_purcharse["税率"].apply(lambda x: "税收{}%（含）".format(int(x * 100)))
    df_purcharse["订单行/关税税额"] = ""
    df_purcharse["订单行/报关单号"] = ""
    df_purcharse["订单行/汇率"] = ""

    print("结果行数:", df_purcharse.shape[0])

    print("转换 ", "成功" if cnt == df_purcharse.shape[0] else "失败")

    print("查看结果:")
    print(df_purcharse.head(10).to_markdown())

    df_purcharse = df_purcharse[
        ["公司", "交货到/数据库 ID", "进销存标识", "供应商", "date_order", "采购员", "订单行/计划日期", "币种", "订单行/产品", "订单行/计量单位", "订单行/订购数量",
         "订单行/单价", "订单行/税率",
         "订单行/关税税额", "订单行/报关单号", "订单行/汇率", "源单据"]]

    # mubiao = r"D:\work\OMS理帐\2C发出商品"
    print("{}\{}".format(targe_dir, newfile))
    df_purcharse.to_excel("{}\{}".format(targe_dir, newfile), index=False)

    print("报告:")
    print("缺少 订单行/产品 的记录:", df_purcharse[df_purcharse["订单行/产品"].isnull()].shape[0])
    print("缺少 交货到/数据库 ID 的记录:", df_purcharse[df_purcharse["交货到/数据库 ID"].isnull()].shape[0])


def convert_purchase_2022(df, sale_type, newfile):
    # 源头
    # 主体	主体	2C\2C发出商品	订单日期	订单日期	店铺名		固定	固定			内部编码	单位	实际卖出	平均价格	税率

    # df = df[df["销售分类"] == sale_type]
    df = df[df["类型"] == sale_type]
    # df = df[df["店铺"] == shop]
    # 读取参考表格

    df_shop = pd.read_excel(shop_file)
    df_sku = pd.read_excel(sku_file)
    df_warehouse = pd.read_excel(warehouse_file, sheet_name="Sheet2")

    # df_sku = df_sku.groupby(["条码"]).agg({"内部参考": max, "显示名称": max, "计量单位/显示名称": max}).reset_index()
    df_sku = df_sku.groupby(["条码"]).agg({"内部参考": max, "名称": max, "计量单位/单位": max}).reset_index()
    df_sku.columns = ["条码", "内部参考", "名称", "计量单位/显示名称"]
    df_sku["条码"] = df_sku["条码"].astype(str)
    df_sku["内部参考"] = df_sku["内部参考"].astype(str)
    df_sku["名称"] = df_sku.apply(lambda x: x["名称"].replace("[{}]".format(x["内部参考"]), "").strip(), axis=1)
    # 公司	交货到/数据库 ID

    # 原始文件
    # filename = r"C:\Users\sjit27\Desktop\文件\212C铲喜官（深圳）日用品有限公司\212C铲喜官（深圳）日用品有限公司\2C铲喜官（深圳）日用品有限公司销售订单.xlsx"
    # df = pd.read_excel(filename)
    # print(df.head(10).to_markdown())

    cnt = df.shape[0]
    print("原表行数:", df.shape[0])

    df_purcharse = df.copy()
    # sale_type = "2C发出商品" if filename.find("2C发出商品") >= 0 else "2C"
    # sale_type = df["类型"].iloc[0]
    # sale_type = df["销售分类"].iloc[0]
    print("销售分类:", sale_type)

    # df_purcharse["公司"] = df_purcharse["供应商"]

    # df_product = pd.read_excel(product_file)
    df_purcharse["商家编码"] = df_purcharse["商家编码"].astype(str)
    # df_product["条码"] = df_product["条码"].astype(str)

    # df_product = df_product.groupby(["条码"]).agg(名称=("名称", "max")).reset_index()
    # df_product.columns = ["条码", "名称"]

    df_purcharse = df_purcharse.merge(df_sku[["条码", "名称"]], how="left", left_on=["商家编码"], right_on=["条码"])

    print("条码为空")
    print(df_purcharse[df_purcharse["条码"].isnull()].head(10).to_markdown())

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
                                      left_on=["商家编码"],
                                      right_on=["条码"])
    print("转换2", "成功" if cnt == df.shape[0] else "失败:{}-{}".format(cnt, df_purcharse.shape[0]))

    # 重命名
    df_purcharse.rename(
        columns={"内部参考": "订单行/产品", "计量单位/显示名称": "订单行/计量单位", "实际卖出数量": "订单行/订购数量"},
        inplace=True)

    print("缺少 订单行/产品 的记录:")
    print(df_purcharse[df_purcharse["订单行/产品"].isnull()].head(10).to_markdown())

    print("缺少 交货到/数据库 ID 的记录:")
    print(df_purcharse[df_purcharse["交货到/数据库 ID"].isnull()].head(10).to_markdown())

    # 订单行/产品
    df_purcharse["订单行/税率"] = df_purcharse["税率"].apply(lambda x: "税收{}%（含）".format(int(x * 100)))
    df_purcharse["订单行/关税税额"] = ""
    df_purcharse["订单行/报关单号"] = ""
    df_purcharse["订单行/汇率"] = ""

    print("结果行数:", df_purcharse.shape[0])

    print("转换 ", "成功" if cnt == df_purcharse.shape[0] else "失败")

    print("查看结果:")
    print(df_purcharse.head(10).to_markdown())

    df_purcharse = df_purcharse[
        ["公司", "交货到/数据库 ID", "进销存标识", "供应商", "date_order", "采购员", "订单行/计划日期", "币种", "订单行/产品", "订单行/计量单位", "订单行/订购数量",
         "订单行/单价", "订单行/税率",
         "订单行/关税税额", "订单行/报关单号", "订单行/汇率", "源单据"]]

    # mubiao = r"D:\work\OMS理帐\2C发出商品"
    # print("{}\{}".format( newfile))
    df_purcharse.to_excel(newfile, index=False)

    print("报告:")
    print("缺少 订单行/产品 的记录:", df_purcharse[df_purcharse["订单行/产品"].isnull()].shape[0])
    print("缺少 交货到/数据库 ID 的记录:", df_purcharse[df_purcharse["交货到/数据库 ID"].isnull()].shape[0])

    return df_purcharse


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
            convert_sales(_file, targe_dir, new_filename.replace(".xlsx", "_转换后.xlsx"))


def Convert_odoo_Salesorder(filename):
    # 转odoo格式
    odoo_so_file = filename.replace(".xls", "_odoo销售订单.xls")
    # for saletype in ["2C","2C发出商品"]:

    saletype = "2C"
    df1 = pd.read_excel(filename)
    df_sales_1 = convert_sales_2022(df1, saletype, odoo_so_file)

    saletype = "2C"
    df2 = pd.read_excel(filename)
    df_sales_2 = convert_sales_2022(df2, saletype, odoo_so_file)

    df3 = df1.append(df2)

    df3.to_excel(odoo_so_file)


if __name__ == "__main__":
    print("开始:", time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()))

    # 调减
    del_fn = []
    # del_fn = [['抖音', '泡研化妆品专营店', '6957712933499', 14881.09], ['抖音', '微笑实验室泡研专卖店', '6930469054686', 199.0]]
    # del_fn = [['抖音', '驰骄日用品专营店', '4897112700265', 2415.66]]
    # del_fn = [['抖音', '尚西个护专营店', '4897112700265', 1428.95], ['抖音', '尚西专营店', '69577129334994', 7172.98]]
    # del_fn = [['抖音', '博滴专卖店', '6973007000066', 44.98]]
    # del_fn = [["京东", "Dentyl Active旗舰店", "5011784002062", 1], ["抖音", "卖家联合个护专营店", "6973007000882", 1323],
    #           ["拼多多", "博滴旗舰店", "6957712933499", 1148.95], ["拼多多", "樱语旗舰店", "6970464040130", 29.6]]
    # del_fn = [['抖音', '播地艾个护专营店', '4897112700609', 1519.0], ['快手', '博滴播地艾专卖店', '6957712933499', 0.03]]

    # 调增
    add_fn = []
    # add_fn=[["快手","播地艾个护专营店",1],["快手","播地艾美妆店",28]]

    # df=tanxiao( "D:\数据处理\mega20.xlsx", "2C发出商品", del_fn, add_fn)
    # df=tanxiao( r"C:\Users\ns2033\Downloads\mega20摊销金额标黄色的.xlsx", "2C发出商品", del_fn, add_fn)
    # df.to_excel("D:\数据处理\mega20摊销金额标黄色的_摊销后.xlsx")

    # tanxiao2( r"D:\数据处理\摊销\摊销前\22仙台（深圳）日用品有限公司2C\22仙台（深圳）日用品有限公司2C销售订单.xlsx", "2C", del_fn, add_fn)
    # tanxiao2(r"D:\数据处理\摊销\摊销前\22伊豆（深圳）化妆品有限公司2C\22伊豆（深圳）化妆品有限公司2C销售订单.xlsx", "2C", del_fn, add_fn)
    # tanxiao2(r"D:\数据处理\摊销\摊销前\22俐思（深圳）化妆品有限公司2C\22俐思（深圳）化妆品有限公司2C销售订单.xlsx", "2C", del_fn, add_fn)   # 找不到产品 4897112700098-1
    # tanxiao2(r"D:\数据处理\摊销\摊销前\22克诺克（深圳）化妆品有限公司2C\22克诺克（深圳）化妆品有限公司2C销售订单.xlsx", "2C", del_fn, add_fn)
    # tanxiao2(r"D:\数据处理\摊销\摊销前\22兰卡斯（深圳）化妆品有限公司2C\22兰卡斯（深圳）化妆品有限公司2C销售订单.xlsx", "2C", del_fn, add_fn)  # '1111111111115'
    # tanxiao2(r"D:\数据处理\摊销\摊销前\22列日（深圳）化妆品有限公司2C\22列日（深圳）化妆品有限公司2C销售订单.xlsx", "2C", del_fn, add_fn)  #
    # tanxiao2(r"D:\数据处理\摊销\摊销前\22加古川（深圳）日用品有限公司2C\22加古川（深圳）日用品有限公司2C销售订单.xlsx", "2C", del_fn, add_fn)  # 完全没有可销售的产品
    # tanxiao2(r"D:\数据处理\摊销\摊销前\22吉维儿（深圳）化妆品有限公司2C\22吉维儿（深圳）化妆品有限公司2C销售订单.xlsx", "2C", del_fn, add_fn)  # 69412770114021
    # tanxiao2(r"D:\数据处理\摊销\摊销前\22和歌山（深圳）日用品有限公司2C\22和歌山（深圳）日用品有限公司2C销售订单.xlsx", "2C", del_fn, add_fn)
    # tanxiao2(r"D:\数据处理\摊销\摊销前\22天下星途（深圳）文化传媒有限公司2C\22天下星途（深圳）文化传媒有限公司2C销售订单.xlsx", "2C", del_fn, add_fn)
    # tanxiao2(r"D:\数据处理\摊销\摊销前\22宫崎（深圳）日用品有限公司2C\22宫崎（深圳）日用品有限公司2C销售订单.xlsx", "2C", del_fn, add_fn)
    # tanxiao2(r"D:\数据处理\摊销\摊销前\22斯帕尔（深圳）化妆品有限公司2C\22斯帕尔（深圳）化妆品有限公司2C销售订单.xlsx", "2C", del_fn, add_fn)   # 完全没有可销售的产品
    # tanxiao2(r"D:\数据处理\摊销\摊销前\22果然净肤（深圳）化妆品有限公司2C\22果然净肤（深圳）化妆品有限公司2C销售订单.xlsx", "2C", del_fn, add_fn)  #  693046905504
    # tanxiao2(r"D:\数据处理\摊销\摊销前\22根特（深圳）化妆品有限公司2C\根特（深圳）化妆品有限公司2C销售订单.xlsx", "2C", del_fn, add_fn)          # MOR1545-11
    # tanxiao2(r"D:\数据处理\摊销\摊销前\22欧脉尔（深圳）化妆品有限公司2C\22欧脉尔（深圳）化妆品有限公司2C销售订单.xlsx", "2C", del_fn, add_fn)  # MOR1484-13
    # tanxiao2(r"D:\数据处理\摊销\摊销前\22沃金顿（深圳）化妆品有限公司2C\22沃金顿（深圳）化妆品有限公司2C销售订单.xlsx", "2C", del_fn, add_fn)  #
    # tanxiao2(r"D:\数据处理\摊销\摊销前\22米兰诺（深圳）化妆品有限公司2C\22米兰诺（深圳）化妆品有限公司2C销售订单.xlsx", "2C", del_fn, add_fn)  # 完全没有可销售的产品
    # tanxiao2(r"D:\数据处理\摊销\摊销前\22约克（深圳）化妆品有限公司2C\22约克（深圳）化妆品有限公司2C销售订单.xlsx", "2C", del_fn, add_fn)  # 完全没有可销售的产品
    # tanxiao2(r"D:\数据处理\摊销\摊销前\22罗伊斯（深圳）日用品有限公司2C\22罗伊斯（深圳）日用品有限公司2C销售订单.xlsx", "2C", del_fn, add_fn) # 完全没有可销售的产品
    # tanxiao2(r"D:\数据处理\摊销\摊销前\22美珂森（深圳）化妆品有限公司2C\22美珂森（深圳）化妆品有限公司2C销售订单.xlsx", "2C", del_fn, add_fn)   # 完全没有可销售的产品
    # tanxiao2(r"D:\数据处理\摊销\摊销前\22美珂森（深圳）日用品有限公司2C\22美珂森（深圳）日用品有限公司2C销售订单.xlsx", "2C", del_fn, add_fn)
    # tanxiao2(r"D:\数据处理\摊销\摊销前\22肌密泉（深圳）化妆品有限公司2C\22肌密泉（深圳）化妆品有限公司2C销售订单.xlsx", "2C", del_fn, add_fn)
    # tanxiao2(r"D:\数据处理\摊销\摊销前\22肌沫（深圳）化妆品有限公司2C\22肌沫（深圳）化妆品有限公司2C销售订单.xlsx", "2C", del_fn, add_fn)
    # tanxiao2(r"D:\数据处理\摊销\摊销前\22谷口（深圳）化妆品有限公司2C\22谷口（深圳）化妆品有限公司2C销售订单.xlsx", "2C", del_fn, add_fn)  #  MOR1125-11
    # tanxiao2(r"D:\数据处理\摊销\摊销前\22长滩（深圳）化妆品有限公司2C\22长滩（深圳）化妆品有限公司2C销售订单.xlsx", "2C", del_fn, add_fn)
    # tanxiao2(r"D:\数据处理\摊销\摊销前\22鹿儿（深圳）化妆品有限公司2C\22鹿儿（深圳）化妆品有限公司2C销售订单.xlsx", "2C", del_fn, add_fn)  # 完全没有可销售的产品

    # tanxiao2(r"D:\数据处理\摊销\摊销前\分摊1.xlsx", "2C", del_fn, add_fn)  # 完全没有可销售的产品

    # tanxiao2(r"D:\数据处理\摊销\摊销前\吉维儿.xlsx",  del_fn, add_fn)  # 完全没有可销售的产品
    # tanxiao2(r"Z:\it审计处理需求\odoo导入\2022年\账套转格式\摊销前\芭葆兔（深圳）日用品有限公司2.xlsx",  del_fn, add_fn)  # 完全没有可销售的产品
    # tanxiao2(r"Z:\it审计处理需求\odoo导入\2022年\账套转格式\摊销前\白卿（深圳）化妆品有限公司2C.xlsx",  del_fn, add_fn)  # 完全没有可销售的产品
    # tanxiao2(r"Z:\it审计处理需求\odoo导入\2022年\账套转格式\摊销前\宝贝港湾（深圳）化妆品有限公司2C.xlsx",  del_fn, add_fn)  # 完全没有可销售的产品
    # tanxiao2(r"Z:\it审计处理需求\odoo导入\2022年\账套转格式\摊销前\宝贝魔术师（深圳）日用品有限公司2C.xlsx",  del_fn, add_fn)  # 完全没有可销售的产品  有问题
    # tanxiao2(r"Z:\it审计处理需求\odoo导入\2022年\账套转格式\摊销前\宝贝配方师（深圳）日用品有限公司2C.xlsx",  del_fn, add_fn)  # 完全没有可销售的产品  有问题
    # tanxiao2(r"Z:\it审计处理需求\odoo导入\2022年\账套转格式\摊销前\茱莉珂丝（深圳）化妆品有限公司2C.xlsx",  del_fn, add_fn)  # 完全没有可销售的产品  有问题
    # tanxiao2(r"Z:\it审计处理需求\odoo导入\2022年\账套转格式\摊销前\植之璨（深圳）化妆品有限公司2C.xlsx",  del_fn, add_fn)  # 完全没有可销售的产品  有问题，报错
    # tanxiao2(r"Z:\it审计处理需求\odoo导入\2022年\账套转格式\摊销前\珍芯漾肤（深圳）化妆品有限公司2C.xlsx",  del_fn, add_fn)  # 完全没有可销售的产品  有问题，报错
    # tanxiao2(r"Z:\it审计处理需求\odoo导入\2022年\账套转格式\摊销前\浙江魔湾电子有限公司2C.xlsx",  del_fn, add_fn)  # 完全没有可销售的产品  有问题，报错
    # tanxiao2(r"Z:\it审计处理需求\odoo导入\2022年\账套转格式\摊销前\造味（深圳）食品有限公司2C.xlsx",  del_fn, add_fn)  # 完全没有可销售的产品  有问题， -4.03
    # tanxiao2(r"Z:\it审计处理需求\odoo导入\2022年\账套转格式\摊销前\约克（深圳）化妆品有限公司2C.xlsx",  del_fn, add_fn)  # 完全没有可销售的产品  有问题，  -1.01
    tanxiao2(r"D:\审计2022\摊销前\深圳秀美妍化妆品有限公司2C.xlsx", del_fn, add_fn)  # 完全没有可销售的产品  有问题，  MOR1545-11  建品
    # tanxiao2(r"Z:\it审计处理需求\odoo导入\2022年\账套转格式\摊销前\萦丝茧（深圳）化妆品有限公司2C.xlsx",  del_fn, add_fn)  # 完全没有可销售的产品   dr0267-11
    # tanxiao2(r"Z:\it审计处理需求\odoo导入\2022年\账套转格式\摊销前\盈养泉（深圳）化妆品有限公司2C.xlsx",  del_fn, add_fn)  #  报错！
    # tanxiao2(r"Z:\it审计处理需求\odoo导入\2022年\账套转格式\摊销前\樱语（深圳）日用品有限公司2C.xlsx",  del_fn, add_fn)  #  6932162333129,'MOR1125-11' '6930469054594-NEW'
    # tanxiao2(r"Z:\it审计处理需求\odoo导入\2022年\账套转格式\摊销前\伊豆（深圳）化妆品有限公司2C.xlsx",  del_fn, add_fn)  # MOR1125-11
    # tanxiao2(r"Z:\it审计处理需求\odoo导入\2022年\账套转格式\摊销前\一片珍芯（深圳）化妆品有限公司2C.xlsx",  del_fn, add_fn)  #MOR1545-11
    # tanxiao2(r"Z:\it审计处理需求\odoo导入\2022年\账套转格式\摊销前\秀美颜（深圳）化妆品有限公司2C.xlsx",  del_fn, add_fn)  # OK
    # tanxiao2(r"Z:\it审计处理需求\odoo导入\2022年\账套转格式\摊销前\仙台（深圳）日用品有限公司2C.xlsx",  del_fn, add_fn)  # 4897112700098-1
    # tanxiao2(r"Z:\it审计处理需求\odoo导入\2022年\账套转格式\摊销前\戏酱（深圳）食品有限公司2C.xlsx",  del_fn, add_fn)  #  报错！
    # tanxiao2(r"Z:\it审计处理需求\odoo导入\2022年\账套转格式\摊销前\无极爽（深圳）日用品有限公司2C.xlsx",  del_fn, add_fn)  #  ok
    # tanxiao2(r"Z:\it审计处理需求\odoo导入\2022年\账套转格式\摊销前\沃金顿（深圳）化妆品有限公司2C.xlsx",  del_fn, add_fn)  #  平  6932162333129
    # tanxiao2(r"Z:\it审计处理需求\odoo导入\2022年\账套转格式\摊销前\茱莉珂丝（深圳）化妆品有限公司2C.xlsx",  del_fn, add_fn)
    # tanxiao2(r"Z:\it审计处理需求\odoo导入\2022年\账套转格式\摊销前\盈养泉（深圳）化妆品有限公司2C.xlsx",  del_fn, add_fn)  # 发货日期为空，金额为空，数据异常
    # tanxiao2(r"Z:\it审计处理需求\odoo导入\2022年\账套转格式\摊销前\深圳樱岚护肤品有限公司2C.xlsx",  del_fn, add_fn)  #
    # tanxiao2(r"Z:\it审计处理需求\odoo导入\2022年\账套转格式\摊销前\尚隐（深圳）计生用品有限公司2C.xlsx",  del_fn, add_fn)  #
    # tanxiao2(r"Z:\it审计处理需求\odoo导入\2022年\账套转格式\摊销前\柚选（深圳）化妆品有限公司2C.xlsx",  del_fn, add_fn)  # MOR1545-11
    # tanxiao2(r"Z:\it审计处理需求\odoo导入\2022年\账套转格式\摊销前\深圳市魔湾游戏科技有限公司2C.xlsx",  del_fn, add_fn)
    # tanxiao2(r"Z:\it审计处理需求\odoo导入\2022年\账套转格式\摊销前\浙江魔湾电子有限公司2C.xlsx",  del_fn, add_fn)   # 123456-生姜泵头
    # tanxiao2(r"Z:\it审计处理需求\odoo导入\2022年\账套转格式\摊销前\深圳造白化妆品有限公司2C.xlsx",  del_fn, add_fn)
    # tanxiao2(r"D:\work\摊销前\@已处理-宝贝港湾（深圳）化妆品有限公司2C.xlsx",  del_fn, add_fn)
    # tanxiao2(r"Z:\it审计处理需求\odoo导入\2022年\账套转格式\摊销前\樱语（深圳）日用品有限公司2C.xlsx",  del_fn, add_fn)  #6930469055522-A   6737664107544 6930469054594-NEW  6737664107544

    # tanxiao2(r"D:\数据处理\摊销\摊销前\吉维儿.xlsx", "2C发出商品", del_fn, add_fn)  # 完全没有可销售的产品

    # Convert_odoo_Salesorder(r"D:\work\@已处理-宝贝港湾（深圳）化妆品有限公司2C.xlsx")

    # company='泡研（深圳）化妆品有限公司'
    # company='广州市尚西国际贸易有限公司'
    # company='深圳市麦凯莱科技有限公司'
    # company='深圳睿旗科技有限公司'
    # company='泉州市航星贸易有限公司'
    # company='广州市尚西国际贸易有限公司'
    # company='深圳睿旗科技有限公司'

    # 转2C
    # convert_2C(company,[['抖音', '泡研化妆品专营店', '6957712933499', 14881.09], ['抖音', '微笑实验室泡研专卖店', '6930469054686', 199.0]],[])
    # convert_2C(company,[],[])
    # convert_2C(company,[['抖音', '航星个护专营店', '6957712933499', 6429.5], ['淘宝', '航星洗护店（tb862892834）', '6937331518536', 3.09], ['天猫', 'cain旗舰店', '6973412410160', 109.0]],[["抖音","航星美妆专营",1.03],["抖音","航星玩具专营店",2.059]])
    # convert_2C(company,[['抖音', '睿旗个护家清专营店', '6973007000035', 792.0], ['天猫', 'unix旗舰店', '8802636250635', 842.09]],[])
    # convert_2C(company,[['抖音', 'BodyAid旗舰店', '6957712933499', 16376.99], ['抖音', 'MOREI家清旗舰店', '4897112700265', 659.0], ['抖音', '麦凯莱个护家清专营店', '6957712933499', 4057.93],
    #                     ['抖音', '麦凯莱美妆专营店', '6957712933499', 7937.95], ['抖音', '麦凯莱专营店', '6973630270089', 4243.9], ['枫叶小店', 'Mega严选', '4897112700562', 3823.17], ['枫叶小店', '麦凯莱好物精选', '4897112700081', 2893.04],
    #                     ['枫叶小店', '麦凯莱精选好物', '4897112700562', 637.75], ['枫叶小店', '麦凯莱严选', '4897112700081', 5007.03], ['京东', 'LoShi旗舰店', '4936201100699', 230.45], ['京东', '博滴官方旗舰店', '6973007000950', 495.3],
    #                     ['京东', '麦凯莱美妆专营店', '6941277011570', 185.75], ['考拉', 'BodyAid博滴旗舰店', '6957712933499', 301.52], ['考拉', '麦凯莱个护专营店', '4897112700081', 200.91], ['马到', '麦凯莱臻选店', '6957712933499', 2605.0],
    #                     ['天猫', 'loshi旗舰店', '4936201054824', 247.48], ['天猫', 'smilelab旗舰店', '7350060860117', 2263.73], ['微盟', '依娜心选好物', '6941277011570', 1057.9], ['有赞', '麦凯莱严选/国际好物', '6957712933499', 18846.13],
    #                     ['有赞', '卖家联合全球购', '6957712933499', 4005.93], ['做梦吧', '博滴官方旗舰店', '6937331518505', 1001.0], ['做梦吧', '麦凯莱好物精选', '4897112700562', 426.0]],[["阿里巴巴","卖家联合loshi总代店",20386.89999],
    #                                                                                                                                                            ["小红书","BodyAid旗舰店",0.49999999997089617]])  # 深圳市麦凯莱科技有限公司
    # convert_2C("广州市尚西国际贸易有限公司",[['抖音', '尚西个护专营店', '4897112700265', 1428.95], ['抖音', '尚西专营店', '6957712933499', 7172.98]],[['百度','尚西总账号',697.008549988037]])
    # convert_2C(company,[],[])

    # filename = r"C:\Users\ns2033\Downloads\2020年2C销售订单.xlsx"
    # print("文件名:")
    # print(filename)
    # new_filename =  r"C:\Users\ns2033\Downloads\2020年2C销售订单_摊销.xlsx"
    # # saletype="2C"
    # df = read_oms(filename, new_filename, "2C", [], [])
    # print("结束")

    # # 转2C发出商品
    # convert_Fachu(company,[['考拉', 'BodyAid博滴旗舰店', '6957712933499', 301.52]],[["考拉","BodyAid博滴旗舰店",304.403429999998]])
    # convert_Fachu(company,[['考拉', 'BodyAid博滴旗舰店', '6957712933499', 301.52]],[["考拉","BodyAid博滴旗舰店",304.403429999998]])  # 深圳市麦凯莱科技有限公司
    # convert_Fachu(company,[],[])  # 深圳市麦凯莱科技有限公司

    # convert_Fachu(company,[],[])

    # convert_Fachu(company,[['拼多多', '航星化妆品专营', '6973007001551', 4.12]],[["抖音","航星美妆专营",1.03],["抖音","航星玩具专营店",3.09]])  # 深圳市麦凯莱科技有限公司
    # convert_Fachu("广州市尚西国际贸易有限公司",[['天猫','allnaturaladvice旗舰店','6917591070500',1.02]],[['抖音','尚西个护专营店',1.02]])
    # # 检查汇总数
    # check_sum_by_shop("广州市尚西国际贸易有限公司")
    # check_sum_by_shop(company)

    # check_2C("若蘅（深圳）化妆品有限公司")
    # check_2C("造白（广州）化妆品有限公司")
    # check_2C("深圳造白化妆品有限公司")
    # check_2C("芭葆兔（深圳）日用品有限公司")

    # check_sum_by_shop("若蘅（深圳）化妆品有限公司")
    # check_sum_by_shop("深圳市卖家优选实业有限公司")
    # check_sum_by_shop("深圳市卖家联合商贸有限公司")
    # check_sum_by_shop("造白（广州）化妆品有限公司")
    # check_sum_by_shop("造白（广州）化妆品有限公司")

    # 按主体汇总
    # cal_sum_by_company()

    # 转odoo格式
    # 目标目录

    # targe_dir = r"D:\数据处理\odoo数据处理"
    # B_convert_2C=1
    # if B_convert_2C==1:
    #     source_dir = r"C:\Users\ns2033\Downloads\摊销\自动处理结果\2C"
    #     # source_dir = r"D:\数据处理\odoo数据处理\陈航\转换前\212C广州驰骄服饰有限公司-转换导入格式前"
    #     convert_all_sales(source_dir, targe_dir, '2C{}.xlsx'.format(company))
    #     # convert_all_sales(source_dir, targe_dir, '2C{}销售订单.xlsx'.format(company))
    #     convert_all_purchase(source_dir, targe_dir, '2C{}_采购订单.xlsx'.format(company))
    #     # convert_all_purchase(source_dir, targe_dir, '2C{}采购订单.xlsx'.format(company))
    #
    # B_convert_2C_Delay = 1
    # if B_convert_2C_Delay==1:
    #     source_dir = r"C:\Users\ns2033\Downloads\摊销\自动处理结果\2C发出商品"
    #     # source_dir = r"D:\数据处理\odoo数据处理\陈航\转换前\212C广州驰骄服饰有限公司-转换导入格式前"
    #     convert_all_sales(source_dir, targe_dir, '2C发出商品{}.xlsx'.format(company))
    #     # convert_all_sales(source_dir, targe_dir, '2C发出商品{}销售订单.xlsx'.format(company))
    #     convert_all_purchase(source_dir, targe_dir, '2C发出商品{}_采购订单.xlsx'.format(company))
    #     # convert_all_purchase(source_dir, targe_dir, '2C发出商品{}采购订单.xlsx'.format(company))
    #
    #     # files=[]
    #
    # files = [targe_dir + os.sep +  '2C{}_转换后.xlsx'.format(company),
    #          targe_dir + os.sep + "2C发出商品{}_转换后.xlsx".format(company),
    #          targe_dir + os.sep + '2C{}_采购订单_转换后.xlsx'.format(company),
    #          targe_dir + os.sep + "2C发出商品{}_采购订单_转换后.xlsx".format(company)]
    #
    # # 在上级目录生成压缩包文件
    # compress_attaches(files, targe_dir + os.sep  + company + '_转odoo后.zip')

    print("结束:", time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()))
    print("ok")
