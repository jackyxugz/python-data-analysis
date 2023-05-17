import pandas as pd
import numpy as np


def read_excel(file):
    df = pd.read_excel(file,dtype=str)
    print(df.head().to_markdown())
    print(len(df))
    df = df[df["平台"].str.contains("百度")]
    print(len(df))
    df = df[df["月份"].str.contains("1月|2月|3月|4月|5月|6月")]
    print(df.head().to_markdown())
    print(len(df))
    df["店铺名称"] = df["店铺名称"].apply(lambda x:get_shop(x))
    df["回款金额"] = df["回款金额"].astype(float)
    df["退款金额"] = df["退款金额"].astype(float)
    df["海外回款金额"] = df["海外回款金额"].astype(float)
    df["海外退款金额"] = df["海外退款金额"].astype(float)
    df["收入金额"] = df["收入金额"].astype(float)
    df["收入金额（不含税）"] = df["收入金额（不含税）"].astype(float)

    group_df = df.groupby(by=["年","月份","主体","店铺名称","平台"]).agg({"回款金额":"sum","退款金额":"sum","海外回款金额":"sum","海外退款金额":"sum","收入金额":"sum","收入金额（不含税）":"sum"})
    group_df = pd.DataFrame(group_df).reset_index()
    print(group_df.head().to_markdown())

    group_df.to_excel(r"D:\沙井\财务账单\2021\百度\财务口径重构\oms百度汇总账单.xlsx",index=False)



def get_shop(shop):
    if shop.find("SF麦凯莱科技")>=0:
        return "麦凯莱总账号"
    elif shop.find("SF麦凯莱")>=0:
        return "鑫桂总账号"
    elif shop.find("深分特殊化妆品-麦凯莱")>=0:
        return "麦凯莱总账号"
    elif shop.find("SF萌")>=0:
        return "宏炽总账号"
    elif shop.find("广分电商-尚西10")>=0:
        return "宏炽总账号"
    elif ((shop.find("广分电商-尚西")>=0)&(shop.find("尚西1")<0)):
        return "尚西总账号（BDCC-SX123）"
    else:
        return shop



if __name__ == "__main__":

    file = r"D:\沙井\财务账单\2021\百度\财务口径重构\账单对比表-2021_0217.xlsx"
    read_excel(file)