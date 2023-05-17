# __coding=utf8__
# /** 作者：zengyanghui **/
import pandas as pd
import numpy as np
from collections import Counter
from pprint import pprint

# df = pd.read_excel('/Volumes/IT审计处理需求/2020/typedata/抖音/订单/麦凯莱个护专营店（抖音小店）订单202012-原表.xlsx')
# df = pd.read_excel('/Volumes/IT审计处理需求/2021/typedata/拼多多/订单/拼多多博滴旗舰店202103 - 原表.xlsx')
df = pd.read_pickle("data/抖音订单_合并表2021.10.pkl")

# print(df)
sku_no = Counter(df["商家编码"])
# 通过调用most_common()方法，能够获取到
# 排序以后的结果
sku_no_sort = sku_no.most_common()
# 以下列表解析的结果是遍历结果并
# 排除掉val <= 1的结果，并返回key
# [item[0] for item in order_no_sort if item[1] > 1]

# sku_no = Counter(df["商家编码"])
# 通过调用most_common()方法，能够获取到
# 排序以后的结果
# sku_no_sort = sku_no.most_common()
# 以下列表解析的结果是遍历结果并
# 排除掉val <= 1的结果，并返回key
# [item[0] for item in order_no_sort if item[1] > 1]
print("订单号，重复数量")

# print("订单号，重复数量")
# for item in order_no_sort:
#     if item[1] > 1:
#         print(item)
# 总行数不计入表头
sku_no_sort = pd.DataFrame(sku_no_sort)

print(f"总行数：{len(sku_no_sort)}")
print(sku_no_sort)
sku_no_sort.to_excel("data/sku_list.xlsx")

# pprint(order_no_sort)