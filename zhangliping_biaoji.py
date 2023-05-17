'''
Desc:合并表并取出需要的数据
Author:Jacky.Xu
Date:2022-08-17
'''
import pandas as pd
import numpy as np
import time
import datetime


def zhangliping2(file_daishou, df_2019, df_2020, df_2021):
    df_daishou = pd.read_excel(file_daishou, sheet_name='Sheet1')
    df_2019 = pd.read_excel(df_2019, sheet_name='Sheet1')
    df_2020 = pd.read_excel(df_2020, sheet_name='Sheet1')
    df_2021 = pd.read_excel(df_2021, sheet_name='Sheet1')
    print(df_daishou.head(5).to_markdown())
    print(df_2019.head(5).to_markdown())
    print(df_2020.head(5).to_markdown())
    print(df_2021.head(5).to_markdown())

    df_daishou["商户订单号"] = df_daishou["商户订单号"].astype(str)
    df_2019["订单号"] = df_2019["订单号"].astype(str)
    df_2020["订单号"] = df_2020["订单号"].astype(str)
    df_2021["订单号"] = df_2021["订单号"].astype(str)

    df_daishou["商户订单号"] = df_daishou["商户订单号"].str.strip()
    df_2019["订单号"] = df_2019["订单号"].str.strip()
    df_2020["订单号"] = df_2020["订单号"].str.strip()
    df_2021["订单号"] = df_2021["订单号"].str.strip()

    df_2019.rename(columns={"订单号": "商户订单号"}, inplace=True)
    df_2020.rename(columns={"订单号": "商户订单号"}, inplace=True)
    df_2021.rename(columns={"订单号": "商户订单号"}, inplace=True)

    df_2019 = df_2019.merge(df_daishou[["商户订单号", "商品名称"]], how="left", on="商户订单号")
    df_2020 = df_2020.merge(df_daishou[["商户订单号", "商品名称"]], how="left", on="商户订单号")
    df_2021 = df_2021.merge(df_daishou[["商户订单号", "商品名称"]], how="left", on="商户订单号")

    print(df_2019.head(5).to_markdown())
    print(df_2020.head(5).to_markdown())
    print(df_2021.head(5).to_markdown())

    df_2019["代收说明"] = df_2019.apply(lambda x: "麦凯莱代收跨境电商款" if (x["商品名称"] == x["商品名称"]) else "",
                                    axis=1)  # 当值为nan时，用自己不等于自己来判断，返回True
    df_2020["代收说明"] = df_2020.apply(lambda x: "麦凯莱代收跨境电商款" if (x["商品名称"] == x["商品名称"]) else "",
                                    axis=1)  # 当值为nan时，用自己不等于自己来判断，返回True
    df_2021["代收说明"] = df_2021.apply(lambda x: "麦凯莱代收跨境电商款" if (x["商品名称"] == x["商品名称"]) else "",
                                    axis=1)  # 当值为nan时，用自己不等于自己来判断，返回True
    print(df_2019[df_2019["代收说明"].str.contains('麦凯莱代收跨境电商款', na=False)].head(10).to_markdown())
    print(df_2020[df_2020["代收说明"].str.contains('麦凯莱代收跨境电商款', na=False)].head(10).to_markdown())
    print(df_2021[df_2021["代收说明"].str.contains('麦凯莱代收跨境电商款', na=False)].head(10).to_markdown())
    # print(df_2019[df_2019["商品名称"]=='nan'].head(10).to_markdown())
    # print("debug_3")
    # print(df_2019[df_2019["商品名称"]!=df_2019["商品名称"]].head(10).to_markdown())

    # del df_2019["商品名称"]
    # del df_2020["商品名称"]
    # del df_2021["商品名称"]

    print(df_2019.head(5).to_markdown())
    print(df_2020.head(5).to_markdown())
    print(df_2021.head(5).to_markdown())

    # df_result["匹配上"] = df_result.apply(
    #     lambda x: 1 if ((x["m_2019"] > 0) | (x["m_2020"] > 0) | (x["m_2021"] > 0)) else 0, axis=1)

    df_2019.to_excel("D:\work\zhangliping\麦凯莱代收跨境电商款\处理后的表格_2019.xlsx")
    print("df_2019 转换完成！")
    df_2020.to_excel("D:\work\zhangliping\麦凯莱代收跨境电商款\处理后的表格_2020.xlsx")
    print("df_2020 转换完成！")
    df_2021.to_excel("D:\work\zhangliping\麦凯莱代收跨境电商款\处理后的表格_2021.xlsx")
    print("df_2021 转换完成！")


if __name__ == "__main__":
    print("开始时间:", time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()))
    start_time = datetime.datetime.now()

    file_daishou = r"D:\work\zhangliping\麦凯莱代收跨境电商款\财务明细汇总\代收跨境电商款.xlsx"
    file_2019 = r"D:\work\zhangliping\麦凯莱代收跨境电商款\年度汇总\2019.xlsx"
    file_2020 = r"D:\work\zhangliping\麦凯莱代收跨境电商款\年度汇总\2020.xlsx"
    file_2021 = r"D:\work\zhangliping\麦凯莱代收跨境电商款\年度汇总\2020.xlsx"

    zhangliping2(file_daishou, file_2019, file_2020, file_2021)
    print("结束时间:", time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()))
    end_time = datetime.datetime.now()
    print("程序执行完毕,共用时 " + str((end_time - start_time).seconds) + " 秒")
