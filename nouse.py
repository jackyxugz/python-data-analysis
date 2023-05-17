import os
import pandas as pd
import numpy  as np
filename = r"D:\22年进销存明细_麦凯莱_0813_01_hsc.xlsx"

df = pd.read_excel(filename)
# print(df.to_markdown)

df1 = df.copy()
df1 = df1.groupby(["产品参考"]).agg({"本期入库数量": sum, "本期出库数量": sum})
df1 = df1[((df1.本期入库数量 == 0) & (df1.本期出库数量 == 0))].reset_index("产品参考")

df1["处理结果"] ="没有动过"

result = df.merge(df1[["产品参考", "处理结果"]], how="left", on="产品参考")
result.to_excel("处理后文件.xlsx")

if df.shape[0] != result.shape[0]:
    print("合并前后数据记录不一致，请核对！")
else:
    print("ok")