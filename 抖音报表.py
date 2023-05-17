import pandas as pd
import numpy as np
import time
import os
import os.path

_array = np.zeros((10000, 10000), dtype="float16")


# def list_all_files(rootdir, filekey_list):
#     if len(filekey_list) > 0:
#         filekey_list = filekey_list.replace(",", " ")
#         filekey = filekey_list.split(" ")
#     else:
#         filekey = ''
#
#     _files = []
#     list = os.listdir(rootdir)  # 列出文件夹下所有的目录与文件
#     for i in range(0, len(list)):
#         path = os.path.join(rootdir, list[i])
#         if os.path.isdir(path):
#             _files.extend(list_all_files(path, filekey_list))
#         if os.path.isfile(path):
#             if ((path.find("~") < 0) and (path.find(".DS_Store") < 0)):  # 带~符号表示临时文件，不读取
#                 if len(filekey) > 0:
#                     for key in filekey:
#                         # print(path)
#                         filename = "".join(path.split("\\")[-1:])
#                         # print("文件名:",filename)
#
#                         key = key.replace("！", "!")
#
#                         if key.find("!") >= 0:
#                             # print("反向选择:",key)
#                             if filename.find(key.replace("!", "")) >= 0:  # 此文件不要读取
#                                 # print("{} 不应该包含 {}，所以剔除:".format(filename,key ))
#                                 pass
#                         elif filename.find(key) > 0:  # 只做文件名的过滤
#                             _files.append(path)
#
#                 else:
#                     _files.append(path)
#
#     # print(_files)
#     return _files
#
#
# def get_all_files(rootdir, filekey):
#     filelist = list_all_files(rootdir, filekey)
#     # print(filelist)
#     mySeries = pd.Series(filelist)
#     df = pd.DataFrame(mySeries)
#     df.columns = ["filename"]
#     # df["filename"] = df["filename"].apply(lambda x: x.strip().replace(" ", ""))
#
#     df=df[~df["filename"].str.contains("汇总")]
#
#     return df
#
#
# def read_all_excel(rootdir, filekey):
#     df_files = get_all_files(rootdir, filekey)
#     for index, file in df_files.iterrows():
#         if 'df' in locals().keys():  # 如果变量已经存在
#             dd = read_excel(file["filename"])
#             dd["filename"] = file["filename"]
#             print("进度表：{}/{}  文件{}，行数{}".format(index + 1, df_files.shape[0], file["filename"], dd.shape[0]))
#
#             df = df.append(dd)
#
#         else:
#             df = read_excel(file["filename"])
#             df["filename"] = file["filename"]
#             # print(file["filename"],  df.shape[0])
#             print("进度表：{}/{}  文件{}，行数{}".format(index + 1, df_files.shape[0], file["filename"], df.shape[0]))
#
#     return df
#
#
# def combine_excel():
#     print('订单数据校对逻辑:')
#     print('1.财务订单数据需要放在财务数据文件夹下，例如/校对数据/财务数据/...')
#     print('2.导出订单数据需要放在导出数据文件夹下，例如/校对数据/导出数据/...')
#     print("请输入财务订单和导出订单所在的文件夹：")
#     # filedir=""
#     # filedir = input()
#     # myTuple = shell.SHBrowseForFolder(0, None, "", 64)
#     try:
#         # path = shell.SHGetPathFromIDList(myTuple[0])
#         filedir = input()
#     except:
#         print("你没有输入任何目录 :(")
#         sys.exit()
#         return
#
#     # filedir=path.decode('ansi')
#     print("你选择的路径是：", filedir)
#
#     global default_dir
#     default_dir = filedir
#
#     if len(filedir) == 0:
#         print("你没有输入任何目录 :(")
#         sys.exit()
#         return
#
#     print('请输入要合并的文件名称或者后缀匹配符:（比如都必须包含 序时账 三个字，那么就请输入  "序时账" ， 不输入就表示所有excel都要合并！）')
#     filekey = input()
#
#     if len(filedir) == 0:
#         print("你没有输入任何关键词 :(")
#         filekey = ''
#         # sys.exit()
#         # return
#
#     print("你希望在'{}'目录下找到所有的包含“{}”文件，然后合并。".format(filedir, filekey))
#
#     # table = read_all_excel(filedir, filekey)
#     table = read_all_excel(filedir, filekey)
#     del table["filename"]
#     table.to_excel(default_dir + os.sep +"处理后的订单.xlsx", index=False)
#
#     return table


def get_order(file):
    # 1. 读取订单
    # f1 = r"D:\数据\DY202110\测试\csv\202110抖音订单_合并表_TEST.csv"
    print(file)
    if file.find("csv") >= 0:
        try:
            orders = pd.read_csv(file, encoding='gb18030', usecols=['主订单编号', '子订单编号', '选购商品', '商品数量', '订单状态', \
                                                                    '支付完成时间', '订单类型', '商家编码', '商品单价', '订单应付金额', \
                                                                    '快递单号', '快递公司', '收件人', '收件地址', '收件人手机号',
                                                                    'filename'])
        except Exception as e:
            orders = pd.read_csv(file, usecols=['主订单编号', '子订单编号', '选购商品', '商品数量', '订单状态', \
                                                '支付完成时间', '订单类型', '商家编码', '商品单价', '订单应付金额', \
                                                '快递单号', '快递公司', '收件人', '收件地址', '收件人手机号', 'filename'])
    elif file.find("xls") >= 0:
        orders = pd.read_excel(file, dtype=str)
    else:
        orders = pd.read_pickle(file)
    print(orders.head(1).to_markdown())
    df_orders = pd.DataFrame(orders)
    # df1 = pd.DataFrame()  # 订单
    # df_orders["商家编码"] = df_orders["商家编码"].apply(lambda x: x.replace("*2", "").replace("*3", "").replace("*4", "").replace("*6", ""))
    df_orders["商家编码"] = df_orders["商家编码"].astype(str)
    df_orders["商家编码"] = df_orders["商家编码"].apply(lambda x: x.replace(" ", "").replace("\n", "+").strip())
    df1 = df_orders

    df1["数据来源"] = df_orders['filename']
    # df1["数据来源"] = ""
    df1["平台"] = "抖音"
    df1['订单编号'] = df_orders['子订单编号']
    df1['主订单编号'] = df_orders['主订单编号']
    df1['店铺'] = df_orders["filename"]
    # df1['店铺'] = ""
    df1['店铺'] = df1['店铺'].apply(lambda x: x.replace("D:\沙井\财务账单\\2022\抖音\\2022.01\原文件\\", "").replace(
        "D:\南山\抖音报表\\2022.01\账单\\", "").replace("/Users/maclove/Downloads/抖音报表/抖音/2021.11/订单/", ""))
    orders["iid"] = orders.index
    orders["出现序号"] = orders.groupby(orders["子订单编号"])["iid"].rank(method='dense')
    df1["出现序号"] = orders["出现序号"]
    df_count = pd.DataFrame(orders["子订单编号"].value_counts()).reset_index()
    df_count.columns = ["子订单编号", "总序号"]
    df1 = orders.merge(df_count, how="left", on="子订单编号")
    df_orders['商品数量'] = df_orders['商品数量'].astype(float)
    df_orders['订单应付金额'] = df_orders['订单应付金额'].astype(float)

    df1['商品名称'] = df_orders['选购商品']
    df1['购买数量'] = df_orders['商品数量']
    df1['订单状态'] = df_orders['订单状态']
    df1['订单时间'] = df_orders['订单提交时间']
    df1['支付方式'] = df_orders['订单类型']
    df1['商家编码'] = df_orders['商家编码']
    df1['销售单价'] = df_orders.apply(lambda x: x['商品单价'] if x['商品单价'].isnumeric() else x['订单应付金额'] / x["商品数量"], axis=1)
    df1['销售金额'] = df_orders['订单应付金额']
    df1['物流单号'] = df_orders['快递单号']
    df1['物流公司'] = df_orders['快递公司']
    df1['收货人姓名'] = df_orders['收件人']
    # df1['收货地址'] = df_orders['收件地址']
    if "收件地址" in df_orders.columns:
        df1['收货地址'] = df_orders.apply(lambda x: x['省'] + x['市'] + x['区'] if pd.isnull(x["收件地址"]) else x["收件地址"], axis=1)
    else:
        df1['收货地址'] = df_orders.apply(lambda x: x['省'] + x['市'] + x['区'] if x["详细地址"].find("*") >= 0 else x["详细地址"],
                                      axis=1)
    df1['联系手机'] = df_orders['收件人手机号']
    del df1["iid"]
    del df1["filename"]

    df1 = df1[
        ["数据来源", "平台", "订单编号", "主订单编号", "店铺", "出现序号", "总序号", "商品名称", "购买数量", "订单状态", "订单时间", "支付方式", "商家编码", "销售单价",
         "销售金额", "物流单号", "物流公司", "收货人姓名", "收货地址", "联系手机"]]

    print(df1.head(1).to_markdown())
    return df1


def get_bill(file):
    # 2. 读取账单
    # f2 = r"D:\数据\DY202110\测试\csv\202110抖音账单_合并表_TEST.csv"
    if file.find("csv") >= 0:
        try:
            # bills = pd.read_csv(file, encoding='gb18030', usecols = ["子订单号",'订单退款(元)','结算状态','结算账户', '订单净收益(元)', '结算时间', '运费(元)'])
            bills = pd.read_csv(file, encoding='gb18030')
        except Exception as e:
            # bills = pd.read_csv(file, usecols = ["子订单号",'订单退款(元)','结算状态','结算账户', '订单净收益(元)', '结算时间', '运费(元)'])
            bills = pd.read_csv(file)
    elif file.find("xls") >= 0:
        bills = pd.read_excel(file, dtype=str)
    else:
        bills = pd.read_pickle(file)
    # print(bills.head(1).to_markdown())
    df_bills = pd.DataFrame(bills)
    df_bills['子订单号'] = df_bills['子订单号'].astype(str)
    df_bills['子订单号'] = df_bills['子订单号'].apply(lambda x: x.replace("'", "").replace('"', '').replace('=', ''))
    # df_bills = df_bills[df_bills["订单号"].str.contains("4852990449509144467")]
    # df_bills.rename(columns={"订单号":"主订单编号"},inplace=True)
    # print(df_bills.head(1).to_markdown())
    df_bills["订单净收益(元)"] = df_bills["订单净收益(元)"].astype(float)
    df_bills["订单退款(元)"] = df_bills["订单退款(元)"].astype(float)
    df_bills["收入"] = df_bills.apply(lambda x: x["订单净收益(元)"] if x["订单净收益(元)"] > 0 else 0, axis=1)
    df_bills["退款"] = df_bills.apply(lambda x: x["订单退款(元)"] if x["订单退款(元)"] < 0 else 0, axis=1)
    # print(df_bills.head(1).to_markdown())
    # 4852990449509144467
    # print(df_bills[df_bills["订单号"].str.contains("4852990449509144467")].to_markdown())
    if "运费实付(元)" in df_bills.columns:
        df_bills.rename(columns={"运费实付(元)": "运费(元)"}, inplace=True)
    print(df_bills.head(1).to_markdown())
    df2 = df_bills.groupby(by=["子订单号", "结算账户"]).agg({"收入": "sum", "退款": "sum", "运费(元)": "sum"})
    df2 = pd.DataFrame(df2).reset_index()
    # print(df2.head(1).to_markdown())
    # print(df2[df2["子订单号"].str.contains("4852990449509144467")].to_markdown())
    # df2 = pd.DataFrame()  # 空账单
    df2['回款日期'] = bills["结算时间"].loc[bills["结算状态"] == "已结算"]
    df2['订单编号'] = df2['子订单号']
    df2['退款金额'] = df2['退款']
    # df2['是否回款'] = ''
    df2.loc[df2.收入 > 0, '是否回款'] = "是"
    df2.loc[df2.收入 == 0, '是否回款'] = "否"
    # df2['是否回款'] = df_bills['是否回款']
    df2['结算方式'] = df2['结算账户']
    df2['回款金额'] = df2['收入']
    # df2['回款日期'] = bills["结算时间"].loc[bills["结算状态"]=="已结算"]
    df2['快递运费'] = df2['运费(元)']
    # df2['结算状态'] = df_bills['结算状态']

    print(df2.head(1).to_markdown())
    return df2


def merge_table(file1, file2):
    # 3. 合并表单：订单表+账单表
    df1 = get_order(file1)
    df2 = get_bill(file2)

    df1['订单编号'] = df1['订单编号'].astype(str)
    df2['订单编号'] = df2['订单编号'].astype(str)

    df3 = pd.merge(df1, df2, how="left", on="订单编号")
    print(df3.head(1).to_markdown())
    return df3


def get_costs(file):
    # 4. 计算成本金额
    # 成本金额 = 成本单价 x 购买数量
    # 成本单价
    # f3 = r"D:\数据\DY202110\测试\抖音成本.xlsx"
    costs = pd.read_excel(file)
    df_costs = pd.DataFrame(costs)
    # print(df_costs.tail(20).to_markdown())
    df_costs.rename(columns={"条码": "商家编码", "线上成本": "成本单价"}, inplace=True)
    df_costs["成本单价"] = df_costs["成本单价"].astype(float)
    df_costs = df_costs[["商家编码", "成本单价"]]
    df_costs["商家编码"] = df_costs["商家编码"].astype(str)
    # df_costs["商家编码"] = df_costs["商家编码"].apply(lambda x:x.upper())
    df_costs["商家编码"] = df_costs["商家编码"].apply(
        lambda x: x.replace(" ", "").replace("'", "").replace('"', '').strip().upper())
    # print("成本表：")
    # print(df_costs.head(1).to_markdown())
    # print(df_costs[df_costs["商家编码"].str.contains("6972818648566")])
    return df_costs


def merge_cost(file1, file2, file3):
    df1 = merge_table(file1, file2)
    df3 = get_costs(file3)
    # df4 = pd.merge(df4,df_costs[["商家编码","线上成本"],how="left",on="商家编码")  # df4,之前是订单和账单合并表
    print(df1.head(1).to_markdown())
    print(df3.head(1).to_markdown())

    df3 = pd.merge(df1, df3, how="left", on="商家编码")

    # 成本金额
    # df3["成本单价"] = df3["成本单价"].replace("￥","")
    df3["成本单价"] = df3["成本单价"].astype(float)
    df3["成本单价"].fillna(0, inplace=True)
    df3["购买数量"] = df3["购买数量"].astype(float)
    df3["成本金额"] = df3["成本单价"] * df3["购买数量"]

    print(df3.head(1).to_markdown())

    # 成本占比
    # for i in df3["总序号"]:
    #     if i == 1:
    #         cost = 1
    #         df3["成本占比"].apply(cost)
    #     else:
    #         cost = 1/i
    #         df3["成本占比"].apply(cost)
    # df3["总序号"] = df3["总序号"].astype(float)
    df3.loc[df3.总序号 == 1, "成本占比"] = 1
    df3.loc[df3.总序号 > 1, "成本占比"] = 1 / (df3.总序号 > 1)
    # df3['成本占比'] = df3['成本金额']/sum_cost

    df3 = df3[
        ["数据来源", "平台", "订单编号", "主订单编号", "店铺", "出现序号", "总序号", "商品名称", "购买数量", "订单状态", "订单时间", "支付方式", "商家编码", "销售单价",
         "销售金额", "成本占比", "成本单价", "成本金额", "物流单号", "物流公司", "收货人姓名", "收货地址", "联系手机", "退款金额",
         "是否回款", "结算方式", "回款金额", "回款日期", "快递运费"]]
    print(df3.head(1).to_markdown())
    # df3.to_csv("/Users/maclove/Downloads/抖音报表",index=False)
    # df3.to_pickle("data/抖音处理后报表.pkl")
    return df3

    # index = 0
    # print("第{}个表格,记录数:{}".format(index, df3.shape[0]))
    # print(df3.head(1).to_markdown())
    # # df.to_excel(r"work/合并表格_test.xlsx")
    # print("账单总行数：")
    # print(df3.shape[0])
    # for i in range(0, int(df3.shape[0] / 800000) + 1):
    #     print("存储分页：{}  from:{} to:{}".format(i, i * 800000, (i + 1) * 800000))
    #     # df.iloc[i*500000:(i+1)*500000].to_csv(default_dir + "\{}_合并表格_{}.csv".format(index,i))
    #     df3.iloc[i * 800000:(i + 1) * 800000].to_csv("data/抖音处理后的报表{}.csv".format(i),index=False)

    # return df3


def split_table(file1, file2, file3):
    # df = pd.read_pickle(file4)
    # df=df[df["订单编号"].str.contains("4852990449509144467")]
    df = merge_table(file1, file2)
    print(f"原数据行数：{len(df)}")
    # 商家编码字段预处理
    df["商家编码"] = df["商家编码"].apply(lambda x: x.upper())
    df["商家编码"] = df["商家编码"].apply(
        lambda x: x.replace("*", "X").replace("+", "-").replace("-NEW", "NEW").replace("-new", "NEW"))
    df["商家编码"] = df["商家编码"].apply(
        lambda x: x.replace("X10", "*10").replace("X11", "*11").replace("X12", "*12").replace("X13", "*13").replace(
            "X14", "*14").replace("X15", "*15").replace("X16", "*16").replace("X17", "*17").replace("X18",
                                                                                                    "*18").replace(
            "X19", "*19"))
    df["商家编码"] = df["商家编码"].apply(lambda x: x.replace("X1", ""))
    df["商家编码"] = df["商家编码"].apply(
        lambda x: x.replace("*10", "X10").replace("*11", "X11").replace("*12", "X12").replace("*13", "X13").replace(
            "*14", "X14").replace("*15", "X15").replace("*16", "X16").replace("*17", "X17").replace("*18",
                                                                                                    "X18").replace(
            "*19", "X19"))
    # df["商家编码"] = df["商家编码"].apply(lambda x:x.replace("*10","X10"))

    # 单个编码数据拆分
    df1 = df[((~df["商家编码"].str.contains("X")) & (~df["商家编码"].str.contains("-")))]
    print(f"单编码数据行数：{len(df1)}")
    # print(df1.head(1).to_markdown())

    # 多个编码，非组合商品
    df2 = df[((df["商家编码"].str.contains("X")) & (~df["商家编码"].str.contains("-")))]
    print(f"多编码数据行数：{len(df2)}")
    # print(df2.head(1).to_markdown())
    # print(df2[df2["订单编号"].str.contains("4852990449509144467")].to_markdown())

    # 组合商品
    df3 = df[df["商家编码"].str.contains("-")]
    print(f"组合商品数据行数：{len(df3)}")
    # print(df3.head(1).to_markdown())
    print("定位1")
    # print(df3[df3["商家编码"].str.contains("5X3")].to_markdown())
    # print(df3[df3["订单编号"].str.contains("4852990449509144467")].to_markdown())

    # 组合商品拆分多编码
    df3_split = df3["商家编码"].str.split("-", expand=True)
    df3_split = df3_split.stack()
    df3_split = df3_split.reset_index()
    df3_split = df3_split.set_index("level_0")
    df3_split.columns = ["level_1", "商家编码"]
    df3_new = df3.drop(["商家编码"], axis=1).join(df3_split)
    # df3_new["商家编码"] = df3_new["商家编码"].apply(lambda x:x.replace("X1",""))
    # print(df3_new.head(1).to_markdown())
    df3_new = df3_new[
        ["数据来源", "平台", "订单编号", "主订单编号", "店铺", "出现序号", "总序号", "商品名称", "购买数量", "订单状态", "订单时间", "支付方式", "商家编码", "销售单价",
         "销售金额", "物流单号", "物流公司", "收货人姓名", "收货地址", "联系手机", "退款金额",
         "是否回款", "结算方式", "回款金额", "回款日期", "快递运费"]]
    print(f"组合商品数据拆分后行数：{len(df3_new)}")
    # print(df3_new.head(1).to_markdown())
    print("定位2")
    # print(df3_new[df3_new["订单编号"].str.contains("4852990449509144467")].to_markdown())

    # 组合商品拆分后再次区分单编码和多编码
    # df3_new["商家编码"] = df3_new["商家编码"].apply(lambda x: x.replace("X1", ""))
    df3_1 = df3_new[((~df3_new["商家编码"].str.contains("X")) & (~df3_new["商家编码"].str.contains("-")))]
    print(f"组合拆分单编码数据行数：{len(df3_1)}")
    df3_2 = df3_new[((df3_new["商家编码"].str.contains("X")) & (~df3_new["商家编码"].str.contains("-")))]
    print(f"组合拆分多编码数据行数：{len(df3_2)}")
    # print(df3_2[df3_2["商家编码"].str.contains("5X3")].to_markdown())
    # 合并单编码
    dfs1 = [df1, df3_1]
    df1_new = pd.concat(dfs1)
    print(f"单编码数据行数：{len(df1_new)}")
    # print(df1_new.head(1).to_markdown())
    # 合并多编码
    # dfs2 = [df2, df3_2]
    # df2_new = pd.concat(dfs2)
    # print(f"多编码数据行数：{len(df2_new)}")
    # print(df2_new.head(1).to_markdown())
    # df2_new.to_pickle("data/多编码抖音数据.pkl")
    # print(df2_new[df2_new["主订单编号"].str.contains("4848566289581260031", na=False)].to_markdown())
    # print("定位3")
    # print(df2_new[df2_new["订单编号"].str.contains("4852990449509144467")].to_markdown())

    # 多编码拆分单编码
    df2_1 = df2[["订单编号", "商家编码"]]
    print(f"多编码数据行数：{len(df2_1)}")
    # df2_1["商家编码"] = df2_1["商家编码"].apply(lambda x: x.replace("X1", ""))
    # print(df2[df2["商家编码"].str.contains("2X2")].to_markdown())
    df2_1["cnt"] = df2_1["商家编码"].apply(lambda x: x[x.find("X") + 1:] if x.find("X") > 0 else 1)
    df2_1["cnt"].replace("2X2", "4", inplace=True)
    df2_1["cnt"].replace("2X3", "6", inplace=True)
    df2_1["cnt"] = df2_1["cnt"].apply(lambda x: 4 if x.find("2X2") >= 0 else x)
    df2_1["cnt"] = df2_1["cnt"].apply(lambda x: 6 if x.find("2X3") >= 0 else x)
    df2_1["商家编码"] = df2_1["商家编码"].apply(lambda x: x[:x.find("X")] if x.find("X") > 0 else x)
    # print(df2_1[df2_1["cnt"].str.contains("2X2")].to_markdown())
    df2_1["cnt"] = df2_1["cnt"].astype(int)
    # print("定位3.1")
    # print(df2_1[df2_1["订单编号"].str.contains("4852990449509144467")].to_markdown())
    # df_qty = pd.read_excel("data/sku_qty.xlsx")
    df_qty = pd.read_excel("data/sku_qty.xlsx")
    df2_result = df2_1.merge(df_qty, how="left", left_on="cnt", right_on="cnt")
    # print(df2_result.head(1).to_markdown())
    # print("定位3.2")
    # print(df2_result[df2_result["订单编号"].str.contains("4852990449509144467")].to_markdown())
    df2_result["订单编号"] = df2_result["订单编号"].str.replace("\s+", "")
    df2["订单编号"] = df2["订单编号"].str.replace("\s+", "")

    df2_result = df2_result.set_index("订单编号")
    df2 = df2.set_index("订单编号")
    print(f"多编码数据拆分前行数：{len(df3)}")
    print("定位1")
    # df2 = df2.drop(["商家编码"], axis=1)
    print(f"df2行数：{len(df2)}\n{df2.head(1).to_markdown()}")
    print(f"df2_result行数：{len(df2_result)}\n{df2_result.head(1).to_markdown()}")
    print("定位2")
    df2.to_excel("data/df2.xlsx")
    df2_result.to_excel("data/df2_result.xlsx")
    df1_1 = df2.drop(["商家编码"], axis=1).join(df2_result)

    # df2_result1 = df2_result.iloc[0:500000]
    # df2_result2 = df2_result.iloc[500000:]
    # print(f"df2_result行数：{len(df2_result1)}")
    # print(f"df2_result行数：{len(df2_result2)}")
    # df1_1 = pd.merge(df2,df2_result,how="left",on="订单编号")
    # dfx = df2.join(df2_result1)
    # dfy = df2.join(df2_result2)
    # df1_1 = pd.concat([dfx,dfy])

    print("定位3")
    df1_1 = df1_1.reset_index()
    df1_1 = df1_1[
        ["数据来源", "平台", "订单编号", "主订单编号", "店铺", "出现序号", "总序号", "商品名称", "购买数量", "订单状态", "订单时间", "支付方式", "商家编码", "销售单价",
         "销售金额", "物流单号", "物流公司", "收货人姓名", "收货地址", "联系手机", "退款金额",
         "是否回款", "结算方式", "回款金额", "回款日期", "快递运费"]]
    print(f"多编码数据拆分后行数：{len(df1_1)}")
    print(df1_1.head(1).to_markdown())
    # print(df1_1[df1_1["主订单编号"].str.contains("4848566289581260031",na=False)].to_markdown())
    print("定位4")
    # print(df1_1[df1_1["订单编号"].str.contains("4852990449509144467")].to_markdown())

    # 组合拆分的多编码拆分单编码
    df2_2 = df3_2[["订单编号", "商家编码"]]
    print(f"多编码数据行数：{len(df2_2)}")
    # df2_2["商家编码"] = df2_2["商家编码"].apply(lambda x: x.replace( "X1", ""))
    df2_2["cnt"] = df2_2["商家编码"].apply(lambda x: x[x.find("X") + 1:] if x.find("X") > 0 else 1)
    df2_2["商家编码"] = df2_2["商家编码"].apply(lambda x: x[:x.find("X")] if x.find("X") > 0 else x)
    print("定位4.1")
    # print(df2_2[df2_2["cnt"].str.contains("5X3")].to_markdown())
    df2_2["cnt"].replace("5X3", "15", inplace=True)
    df2_2["cnt"].replace("2X2", "4", inplace=True)
    df2_2["cnt"] = df2_2["cnt"].apply(lambda x: 15 if x.find("5X3") >= 0 else x)
    df2_2["cnt"] = df2_2["cnt"].apply(lambda x: 4 if x.find("2X2") >= 0 else x)
    # print(df2_2[df2_2["cnt"].str.contains("5X3")].to_markdown())
    df2_2["cnt"] = df2_2["cnt"].astype(int)
    # print(df2_2[df2_2["订单编号"].str.contains("4852990449509144467")].to_markdown())
    df_qty = pd.read_excel("data/sku_qty.xlsx")
    df2_result = df2_2.merge(df_qty, how="left", left_on="cnt", right_on="cnt")
    # print(df2_result.head(1).to_markdown())
    print("定位4.2")
    # print(df2_result[df2_result["订单编号"].str.contains("4852990449509144467")].to_markdown())
    df2_result = df2_result.set_index("订单编号")
    df3_3 = df3.set_index("订单编号")
    print(f"多编码数据拆分前行数：{len(df3)}")
    df1_2 = df3_3.drop(["商家编码"], axis=1).join(df2_result)
    df1_2 = df1_2.reset_index()
    df1_2 = df1_2[
        ["数据来源", "平台", "订单编号", "主订单编号", "店铺", "出现序号", "总序号", "商品名称", "购买数量", "订单状态", "订单时间", "支付方式", "商家编码", "销售单价",
         "销售金额", "物流单号", "物流公司", "收货人姓名", "收货地址", "联系手机", "退款金额",
         "是否回款", "结算方式", "回款金额", "回款日期", "快递运费"]]
    print(f"多编码数据拆分后行数：{len(df1_2)}")
    # print(df1_2.head(1).to_markdown())
    # print(df1_1[df1_1["主订单编号"].str.contains("4848566289581260031",na=False)].to_markdown())
    print("定位5")
    # print(df1_2[df1_2["订单编号"].str.contains("4852990449509144467")].to_markdown())

    # 合并单编码数据
    dfs3 = [df1_new, df1_1, df1_2]
    df_new = pd.concat(dfs3)
    # print(f"合并后单编码数据行数：{len(df_new)}")
    # print(df_new.head(1).to_markdown())
    print("定位6")
    # print(df_new[df_new["订单编号"].str.contains("4852990449509144467")].to_markdown())
    # df_new.to_pickle("data/单编码抖音数据.pkl")

    return df_new


def cost_merge(file1, file2, file3):
    # df1 = pd.read_pickle("data/单编码抖音数据.pkl")
    df1 = split_table(file1, file2, file3)
    df3 = get_costs(file3)

    # print(df1.head(1).to_markdown())
    # print("定位1")
    # print(df1[df1["订单编号"].str.contains("4852990449509144467")])
    del df1["出现序号"]
    del df1["总序号"]
    df1["商家编码"] = df1["商家编码"].astype(str)
    df1["商家编码"] = df1["商家编码"].apply(lambda x: x.replace(" ", "").replace("'", "").replace('"', '').strip())
    df1.dropna(subset=["商家编码"], inplace=True)
    df1["商家编码"].fillna("xx99", inplace=True)
    df1 = df1[~df1["商家编码"].str.contains("xx99|nan")]

    # print("定位2")
    # print(df1[df1["订单编号"].str.contains("4852990449509144467")])
    # print(df1.head(1).to_markdown())
    df = pd.merge(df1, df3, how="left", on="商家编码")
    # df["成本单价"] = df["成本单价"].replace("￥","")
    print(df.head(1).to_markdown())
    df["成本单价"] = df["成本单价"].astype(float)
    df["成本单价"].fillna(0, inplace=True)
    df["购买数量"] = df["购买数量"].astype(int)
    df["成本金额"] = df["成本单价"] * df["购买数量"]

    df_group = df.groupby(["主订单编号"]).agg({"成本金额": "sum"})
    df_group.rename(columns={"成本金额": "成本总金额"}, inplace=True)
    # print(df_group.head(1).to_markdown())
    df = df.merge(df_group, how="left", on="主订单编号")
    # print(df.head(1).to_markdown())
    # df["成本占比"] = df.apply(lambda x: x["成本金额"]/x["成本总金额"] if abs(x["成本总金额"] - x["成本金额"])<0.001 else 1,axis=1)
    df["成本占比"] = df.apply(lambda x: x["成本金额"] / x["成本总金额"] if x["成本总金额"] != x["成本金额"] else 1, axis=1)
    # print(df.loc[df["总序号"] > 1].head(1).to_markdown())
    # print("定位3")
    # print(df[df["订单编号"].str.contains("4852990449509144467")])
    df["销售金额"] = df["销售金额"].astype(float)
    df["销售单价"] = df["销售单价"].astype(float)
    df["成本占比"] = df["成本占比"].astype(float)
    df["销售金额1"] = df["销售金额"] * df["成本占比"]
    df["销售单价1"] = df["销售金额1"] / df["购买数量"]
    df["退款金额"].fillna(0, inplace=True)
    df["回款金额"].fillna(0, inplace=True)
    df["是否回款"].fillna("否", inplace=True)
    df["快递运费"].fillna(0, inplace=True)
    df["回款金额"] = df["回款金额"].astype(float)
    df["退款金额"] = df["退款金额"].astype(float)
    df["回款金额1"] = df["回款金额"] * df["成本占比"]
    df["退款金额1"] = df["退款金额"] * df["成本占比"]

    # df_group = df.groupby(["订单编号"]).agg({"销售金额1":"sum","回款金额1":"sum","退款金额1":"sum","销售金额":"mean","回款金额":"mean","退款金额":"mean"})
    # print(df_group.head(1).to_markdown())
    # df_group.to_pickle("data/拆分后再次汇总金额.pkl")
    del df["销售单价"]
    del df["销售金额"]
    del df["退款金额"]
    del df["回款金额"]
    del df["成本总金额"]
    df.rename(columns={"销售金额1": "销售金额", "销售单价1": "销售单价", "回款金额1": "回款金额", "退款金额1": "退款金额"}, inplace=True)
    # df["成本占比"] = df["成本占比"].apply(lambda x:round(x,2))
    # df["销售金额"] = df["销售金额"].apply(lambda x:round(x,2))
    # df["销售单价"] = df["销售单价"].apply(lambda x:round(x,2))
    # df["回款金额"] = df["回款金额"].apply(lambda x:round(x,2))
    # df["退款金额"] = df["退款金额"].apply(lambda x:round(x,2))
    df.dropna(subset=["订单编号"], inplace=True)
    df = df[~df["订单编号"].str.contains("nan")]
    df["iid"] = df.index
    df["出现序号"] = df.groupby(df["订单编号"])["iid"].rank(method='dense')
    df_count = pd.DataFrame(df["订单编号"].value_counts()).reset_index()
    # print(df_count.to_markdown())
    df_count.columns = ["订单编号", "总序号"]
    df = df.merge(df_count, how="left", on="订单编号")
    # df = df.apply(lambda x:x.replace("nan","").replace("NAN",""))
    print(df.tail(1).to_markdown())
    df["订单编号"] = df["订单编号"].astype(str)
    df["主订单编号"] = df["主订单编号"].astype(str)
    df["商家编码"] = df["商家编码"].astype(str)
    df["物流公司"] = df["物流公司"].astype(str)
    df["收货地址"] = df["收货地址"].astype(str)
    df["物流单号"] = df["物流单号"].astype(str)
    df["快递运费"] = df.apply(lambda x: get_express(x["物流公司"], x["收货地址"], x["出现序号"], x["物流单号"]), axis=1)

    df = df[
        ["数据来源", "平台", "订单编号", "主订单编号", "店铺", "出现序号", "总序号", "商品名称", "购买数量", "订单状态", "订单时间", "支付方式", "商家编码", "销售单价",
         "销售金额", "成本占比", "成本单价", "成本金额", "物流单号", "物流公司", "收货人姓名", "收货地址", "联系手机", "退款金额",
         "是否回款", "结算方式", "回款金额", "回款日期", "快递运费"]]
    df = df.sort_values(by=["数据来源", "订单编号", "出现序号"])
    print(df.tail(1).to_markdown())
    # print("定位4")
    # print(df[df["订单编号"].str.contains("4852990449509144467")])
    df = df.sort_values(by=["数据来源", "订单编号", "出现序号"])

    index = 0
    print("第{}个表格,记录数:{}".format(index, df.shape[0]))
    print(df.head(1).to_markdown())
    # df.to_excel(r"work/合并表格_test.xlsx")
    print("账单总行数：")
    print(df.shape[0])
    for i in range(0, int(df.shape[0] / 800000) + 1):
        print("存储分页：{}  from:{} to:{}".format(i, i * 800000, (i + 1) * 800000))
        # df.iloc[i*500000:(i+1)*500000].to_csv(default_dir + "\{}_合并表格_{}.csv".format(index,i))
        df.iloc[i * 800000:(i + 1) * 800000].to_csv("data/{}抖音处理后的报表{}.csv".format(index, i), index=False)


def get_express(x, y, z, e):
    if z == 1:
        if ((x.find("ems") >= 0) or (x.find("邮政") >= 0) or (x.find("邮政") >= 0)):
            if y.find("广东") >= 0:
                return 3.9
            elif ((y.find("湖南") >= 0) or (y.find("海南") >= 0) or (y.find("广西") >= 0) or (y.find("江西") >= 0) or (
                    y.find("湖北") >= 0) or (y.find("福建") >= 0)):
                return 5.2
            elif ((y.find("贵州") >= 0) or (y.find("河南") >= 0) or (y.find("云南") >= 0) or (y.find("浙江") >= 0) or (
                    y.find("上海") >= 0) or (y.find("重庆") >= 0) or (y.find("江苏") >= 0) or (y.find("河北") >= 0) or (
                          y.find("陕西") >= 0) or (y.find("四川") >= 0) or (y.find("山东") >= 0) or (y.find("安徽") >= 0) or (
                          y.find("山西") >= 0)):
                return 6.5
            elif ((y.find("北京") >= 0) or (y.find("天津") >= 0) or (y.find("甘肃") >= 0) or (y.find("宁夏") >= 0) or (
                    y.find("内蒙古") >= 0) or (y.find("辽宁") >= 0) or (y.find("吉林") >= 0) or (y.find("黑龙江") >= 0)):
                return 9.1
            elif ((y.find("西藏") >= 0) or (y.find("青海") >= 0) or (y.find("新疆") >= 0)):
                return 16
            else:
                return 0
        elif x.find("顺丰") >= 0:
            if ((y.find("深圳") >= 0) or (y.find("广东") >= 0)):
                return 8
            elif ((y.find("湖南") >= 0) or (y.find("海南") >= 0) or (y.find("广西") >= 0) or (y.find("江西") >= 0) or (
                    y.find("福建") >= 0)):
                return 10
            elif ((y.find("贵州") >= 0) or (y.find("河南") >= 0) or (y.find("云南") >= 0) or (y.find("浙江") >= 0) or (
                    y.find("上海") >= 0) or (y.find("重庆") >= 0) or (y.find("江苏") >= 0) or (y.find("湖北") >= 0) or (
                          y.find("四川") >= 0) or (y.find("安徽") >= 0)):
                return 13
            elif ((y.find("北京") >= 0) or (y.find("河北") >= 0) or (y.find("辽宁") >= 0) or (y.find("山东") >= 0) or (
                    y.find("山西") >= 0) or (y.find("陕西") >= 0) or (y.find("天津") >= 0)):
                return 14
            elif ((y.find("甘肃") >= 0) or (y.find("呼和浩特") >= 0) or (y.find("包头") >= 0) or (y.find("乌兰察布") >= 0) or (
                    y.find("鄂尔多斯") >= 0) or (y.find("巴彦淖尔") >= 0) or (y.find("乌海") >= 0) or (y.find("阿拉善盟") >= 0) or (
                          y.find("赤峰") >= 0) or (y.find("通辽") >= 0) or (y.find("锡林郭勒盟") >= 0) or (
                          y.find("宁夏") >= 0) or (
                          y.find("青海") >= 0)):
                return 17
            elif ((y.find("黑龙江") >= 0) or (y.find("吉林") >= 0) or (y.find("呼伦贝尔") >= 0) or (y.find("兴安盟") >= 0)):
                return 19
            elif ((y.find("新疆") >= 0) or (y.find("西藏") >= 0)):
                return 21
            else:
                return 0
        elif len(x) < 2:
            if len(e) > 3:
                return 3.5
            else:
                return 0
        elif ((x.find("nan") >= 0) or (x.find("NAN") >= 0)):
            if len(e) > 3:
                return 3.5
            else:
                return 0
        else:
            return 3.5
            # elif x.find("中通") >= 0:
            #     return 3.5
            # elif x.find("申通") >= 0:
            #     return 3.5
            # elif x.find("圆通") >= 0:
            #     return 3.5
    else:
        return 0


def bill_table(file):
    df = pd.read_pickle(file)


def cover(file1, file2):
    if file1.find("csv") >= 0:
        df1 = pd.read_csv(file1, dtype=str)
    else:
        df1 = pd.read_excel(file1, dtype=str)
    if file2.find("csv") >= 0:
        df2 = pd.read_csv(file2, dtype=str)
    else:
        df2 = pd.read_excel(file2, dtype=str)

    df1.to_pickle("data/抖音订单_合并表2022.01.pkl")
    df2.to_pickle("data/抖音账单_合并表2022.01.pkl")


if __name__ == "__main__":
    # 文件路径
    # order_file = r"D:\沙井\财务账单\2022\抖音\2022.01\原文件\订单_合并表格.xlsx"
    order_file = r"D:\南山\抖音报表\2022.01\订单\合并表格.csv"
    # order_file = r"/Users/maclove/Downloads/抖音报表/Montooth/订单\合并表格.xlsx"
    # order_file = "data/抖音订单_合并表2022.01.pkl"
    # order_file2 = "data/抖音订单2_合并表2022.01.pkl"
    # order_file3 = "data/抖音订单3_合并表2022.01.pkl"
    # order_file4 = "data/抖音订单4_合并表2022.01.pkl"
    # order_file5 = "data/抖音订单5_合并表2022.01.pkl"
    # order_file6 = "data/抖音订单6_合并表2022.01.pkl"
    # order_file7 = "data/抖音订单7_合并表2022.01.pkl"
    # order_file = "/itdd/zengyanghui/抖音订单_合并表2022.01.pkl"
    # bill_file = r"D:\沙井\财务账单\2022\抖音\2022.01\原文件\账单_合并表格.xlsx"
    bill_file = r"D:\南山\抖音报表\2022.01\账单\合并表格.csv"
    # bill_file = r"/Users/maclove/Downloads/抖音报表/Montooth/账单\合并表格.xlsx"
    # bill_file = "data/抖音账单_合并表2022.01.pkl"
    # bill_file2 = "data/抖音账单2_合并表2022.01.pkl"
    # bill_file3 = "data/抖音账单3_合并表2022.01.pkl"
    # bill_file4 = "data/抖音账单4_合并表2022.01.pkl"
    # bill_file5 = "data/抖音账单5_合并表2022.01.pkl"
    # bill_file6 = "data/抖音账单6_合并表2022.01.pkl"
    # bill_file7 = "data/抖音账单7_合并表2022.01.pkl"
    # bill_file = "/itdd/zengyanghui/抖音账单_合并表2022.01.pkl"
    costs_file = r"D:\南山\抖音报表\抖音成本(6).xlsx"
    # costs_file = "/itdd/zengyanghui/抖音成本(1)(3).xlsx"
    # result_file = "data/抖音处理后报表.pkl"

    # 调试步骤
    # get_order(order_file)
    # get_bill(bill_file)
    # merge_table(order_file,bill_file)
    # get_costs(costs_file)
    cover(order_file, bill_file)

    # 拆分sku
    # split_table(order_file,bill_file,costs_file,result_file)

    # 抖音订单+账单+成本
    # merge_cost(order_file,bill_file,costs_file) #弃用

    # cost_merge(order_file, bill_file, costs_file)
    # global default_index
    # default_index = 0
    # cost_merge(order_file1, bill_file1, costs_file)

    # default_index = 1
    # cost_merge(order_file2, bill_file2, costs_file)

    # default_index = 2
    # cost_merge(order_file3, bill_file3, costs_file)

    # 这个报错！！！
    # default_index = 3
    # cost_merge(order_file4, bill_file4, costs_file)

    # default_index = 4
    # cost_merge(order_file5, bill_file5, costs_file)

    # default_index = 5
    # cost_merge(order_file6, bill_file6, costs_file)

    # default_index = 6
    # cost_merge(order_file7, bill_file7, costs_file)

    # 抖音账单
    # bill_table(bill_file)
