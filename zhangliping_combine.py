'''
Desc:合并表并取出需要的数据
Author:Jacky.Xu
Date:2022-08-17
'''
import pandas as pd
import numpy as np
import time
import datetime


def zhangliping3(file_caiwu, file_all):
    # df_result = get_df_order(df_order, df_huikuan)
    # df_result.head(20).to_markdown()
    df_file_caiwu = pd.read_excel(file_caiwu, sheet_name='Sheet1')
    df_all = pd.read_excel(file_all, sheet_name='Sheet1')

    print(df_file_caiwu.head(5).to_markdown())
    print(df_all.head(5).to_markdown())

    # df_all = df_all.groupby(["订单号"]).agg({"订单号": "count"}).reset_index()

    # df_all = df_all[["订单号"]].reset_index()
    # df_all.drop_duplicates(inplace=True)

    df_file_caiwu["商户订单号"] = df_file_caiwu["商户订单号"].astype(str)
    df_all["订单号"] = df_all["订单号"].astype(str)

    df_file_caiwu["商户订单号"] = df_file_caiwu["商户订单号"].str.strip()
    df_all["订单号"] = df_all["订单号"].str.strip()

    # df_2019.rename(columns={"订单号": "商户订单号"}, inplace=True)

    # df_2019["m_2019"] = 1

    # df_result = df_result.merge(df_2019[["商户订单号", "m_2019"]], how="left", on="商户订单号")

    df_all_sum = df_all.merge(df_file_caiwu[["商户订单号", "商品名称", "收入金额（+元）"]], how="left", left_on="订单号", right_on="商户订单号")

    df_all_sum["代收记录"] ="代收跨境电商费用"

    print(df_all_sum.head(5).to_markdown())
    print(df_file_caiwu.head(5).to_markdown())

    lst_order_no = df_all_sum['订单号'].drop_duplicates().tolist()
    df_chayi = df_file_caiwu[(~df_file_caiwu['商户订单号'].isin(lst_order_no))]

    # df_result.to_excel("D:\work\zhangliping\麦凯莱代收跨境电商款\处理后的表格_徐.xlsx")
    df_all_sum.to_excel("D:\work\zhangliping\麦凯莱代收跨境电商款\处理后的表格_2019_2020_2021.xlsx")
    df_chayi.to_excel(r"D:\work\zhangliping\麦凯莱代收跨境电商款\差异表.xlsx")


if __name__ == "__main__":
    print("开始时间:", time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()))
    start_time = datetime.datetime.now()
    file_caiwu = r"D:\work\zhangliping\麦凯莱代收跨境电商款\财务明细汇总\代收跨境电商款.xlsx"
    file_all = r"D:\work\zhangliping\麦凯莱代收跨境电商款\年度汇总\0_合并表格_0_20220817114635.xlsx"

    zhangliping3(file_caiwu, file_all)
    print("结束时间:", time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()))
    end_time = datetime.datetime.now()
    print("完成,共用时 " + str((end_time - start_time).seconds) + " 秒")
