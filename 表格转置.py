import pandas as pd

def cal_error_percent(fn, oms):
    # print(fn)
    # print(oms)
    if len((str(fn)))==0:
        return "100%"
    elif len((str(oms)))==0:
        return "100%"
    elif int(oms) == 0:
        return "100%"
    elif int(fn) == 0:
        return "100%"
    else:
        return "{:.2f}".format(abs(int(fn) - int(oms)) / max(fn, oms) * 100.00) + "%"


def report_v2h(filename, iyear):
    # 平台	店铺名称	年度	月份	财务-订单数量	财务-订单金额	导出-订单数量	导出-订单金额	数量差异（财务/订单）	金额差异（财务/订单）
    df = pd.read_excel(filename)
    df = df[df["年度"].isin([iyear])]
    df["财务-订单数量"].fillna(0, inplace=True)
    df["财务-订单金额"].fillna(0, inplace=True)

    df["导出-订单数量"].fillna(0, inplace=True)
    df["导出-订单金额"].fillna(0, inplace=True)

    df["财务-订单数量"] = df["财务-订单数量"].astype(int)
    df["导出-订单数量"] = df["导出-订单数量"].astype(int)

    df2 = df.copy()

    for i in range(1, 13):
        df2["订单行(财务/oms)-{}".format(i)] = df2.apply(
            lambda x: "{}/{}".format(x["财务-订单数量"], x["导出-订单数量"]) if int(x["月份"]) ==i else "", axis=1)
        df2["订单金额(财务/oms)-{}".format(i)] = df2.apply(
            lambda x: "{}/{}".format(x["财务-订单金额"], x["导出-订单金额"]) if int(x["月份"]) == i else "", axis=1)

        df2["行差异(财务/订单)-{}".format(i)] = df2.apply(
            lambda x: cal_error_percent(x["财务-订单数量"], x["导出-订单数量"]) if int(x["月份"]) == i else "", axis=1)
        df2["金额差异(财务/订单)-{}".format(i)] = df2.apply(
            lambda x: cal_error_percent(x["财务-订单金额"], x["导出-订单金额"]) if int(x["月份"]) == i else "", axis=1)

    del df2["Unnamed: 0"]
    del df2["月份"]
    del df2["财务-订单数量"]
    del df2["财务-订单金额"]
    del df2["导出-订单数量"]
    del df2["导出-订单金额"]
    del df2["数量差异（财务/订单）"]
    del df2["金额差异（财务/订单）"]

    df3 = df2.groupby(["平台", "店铺名称", "年度"]).agg(
        {"订单行(财务/oms)-1": np.max, "订单金额(财务/oms)-1": np.max, "行差异(财务/订单)-1": np.max, "金额差异(财务/订单)-1": np.max,
         "订单行(财务/oms)-2": np.max, "订单金额(财务/oms)-2": np.max, "行差异(财务/订单)-2": np.max, "金额差异(财务/订单)-2": np.max,
         "订单行(财务/oms)-3": np.max, "订单金额(财务/oms)-3": np.max, "行差异(财务/订单)-3": np.max, "金额差异(财务/订单)-3": np.max,
         "订单行(财务/oms)-4": np.max, "订单金额(财务/oms)-4": np.max, "行差异(财务/订单)-4": np.max, "金额差异(财务/订单)-4": np.max,
         "订单行(财务/oms)-5": np.max, "订单金额(财务/oms)-5": np.max, "行差异(财务/订单)-5": np.max, "金额差异(财务/订单)-5": np.max,
         "订单行(财务/oms)-6": np.max, "订单金额(财务/oms)-6": np.max, "行差异(财务/订单)-6": np.max, "金额差异(财务/订单)-6": np.max,
         "订单行(财务/oms)-7": np.max, "订单金额(财务/oms)-7": np.max, "行差异(财务/订单)-7": np.max, "金额差异(财务/订单)-7": np.max,
         "订单行(财务/oms)-8": np.max, "订单金额(财务/oms)-8": np.max, "行差异(财务/订单)-8": np.max, "金额差异(财务/订单)-8": np.max,
         "订单行(财务/oms)-9": np.max, "订单金额(财务/oms)-9": np.max, "行差异(财务/订单)-9": np.max, "金额差异(财务/订单)-9": np.max,
         "订单行(财务/oms)-10": np.max, "订单金额(财务/oms)-10": np.max, "行差异(财务/订单)-10": np.max, "金额差异(财务/订单)-10": np.max,
         "订单行(财务/oms)-11": np.max, "订单金额(财务/oms)-11": np.max, "行差异(财务/订单)-11": np.max, "金额差异(财务/订单)-11": np.max,
         "订单行(财务/oms)-12": np.max, "订单金额(财务/oms)-12": np.max, "行差异(财务/订单)-12": np.max,
         "金额差异(财务/订单)-12": np.max}).reset_index()
    print(df3.to_markdown())
    df3.to_excel(r"work/财务和导出订单的数量和金额差距_横向.xlsx")


if __name__ == "__main__":
    # print("开始:", time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()))

    report_v2h(r"work/财务和导出订单的数量和金额差距(2).xlsx", 2019)
