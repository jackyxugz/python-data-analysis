import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import time
import sys


def find_user():
    print("加载考勤日历文件")
    df_date = pd.read_excel(r"/Users/maclove/Downloads/考勤脚本/考勤政策\合并表格.xlsx")
    df_date["ddate"] = df_date["ddate"].astype("datetime64[ns]").dt.date
    print(df_date.head().to_markdown())

    print("请输入需要处理的文件路径：")
    file = input()
    # file = r"/Users/maclove/Downloads/考勤脚本/2021年11月_考勤明细表.xlsx"
    print("你输入的文件是：{}".format(file))

    list = ["李春雷", "魏佳琳", "罗娟娟", "刘凯兆", "杨景全", "时磊", "肖文军", "江畅", "李远福", "彭志颖", "何艳", "陈超", "曾杨辉", "谢崇港", "金昆",
            "徐兰珠", "王珂", "叶婷", "李志", "刘欢", "黄明华", "杨渭欢", "张志平", "陈茂祥", "陈树凡", "徐贵中"]

    df = pd.read_excel(file, dtype=str)
    if "员工名称" in df.columns:
        df.rename(columns={"员工名称": "姓名", "打卡日期": "日期", "上班时间": "上班打卡时间", "下班时间": "下班打卡时间"}, inplace=True)

    df["日期"] = df["日期"].astype("datetime64[ns]").dt.date
    df = pd.merge(df, df_date, how="left", left_on="日期", right_on="ddate")
    print(df.head(1).to_markdown())

    df1 = None
    for name in list:
        if df1 is None:
            df1 = df.loc[df["姓名"] == name]
        else:
            df0 = df.loc[df["姓名"] == name]
            df1 = df1.append(df0)

    print(df1.head(1).to_markdown())
    df1.dropna(subset=["上班打卡时间", "下班打卡时间"], inplace=True)

    # 上班打卡时间按30分钟往后取整
    df1["上班打卡(用于加班计算)"] = df1["上班打卡时间"].apply(lambda x: datetime.strptime(x[:14] + "30:00", '%Y-%m-%d %H:%M:%S') if int(x[14:16]) < 30 else datetime.strptime(x[:11] + str(int(x[11:13]) + 1) + ":00:00", '%Y-%m-%d %H:%M:%S'))
    df1["上班打卡(用于加班计算)"] = df1["上班打卡(用于加班计算)"].astype(str)
    df1["上班打卡(用于加班计算)"] = df1["上班打卡(用于加班计算)"].apply(lambda x: datetime.strptime(x[:11] + "08:00:00", '%Y-%m-%d %H:%M:%S') if (datetime.strptime(x, "%Y-%m-%d %H:%M:%S").time() <= datetime.strptime("08:00:00", "%H:%M:%S").time()) else datetime.strptime(x, "%Y-%m-%d %H:%M:%S"))

    # 下班打卡时间按30分钟往前取整
    # df1["下班打卡(用于加班计算)"] = df1["下班打卡时间"].apply(lambda x: datetime.strptime(
    #     x[:14] + "00:00", '%Y-%m-%d %H:%M:%S') if int(x[14:16]) < 29 else datetime.strptime(
    #     x[:14] + "30:00", '%Y-%m-%d %H:%M:%S'))
    df1["下班打卡(用于加班计算)"] = df1["下班打卡时间"].apply(lambda x: get_overtime(x))

    # 计算早到加班
    df1["早上加班(分钟)"] = df1.apply(lambda x: 30 if (
            (datetime.strptime(x["上班打卡时间"], "%Y-%m-%d %H:%M:%S").time() <= datetime.strptime("08:00:00",
                                                                                             "%H:%M:%S").time()) & (
                    datetime.strptime(x["下班打卡时间"], "%Y-%m-%d %H:%M:%S").time() >= datetime.strptime("18:00:00",
                                                                                                    "%H:%M:%S").time()) & (
                        x["daytype"] == 0)) else 0, axis=1)

    # 统一计算晚上加班分钟
    df1["晚上加班(分钟)"] = (pd.to_datetime(df1["下班打卡(用于加班计算)"]) - pd.to_datetime(df1["上班打卡(用于加班计算)"])) / np.timedelta64(1,
                                                                                                                   "m")
    df1["晚上加班(分钟)"] = df1["晚上加班(分钟)"] - df1["早上加班(分钟)"] - 480 - 60 - 30

    # 周日、节假日重新计算加班分钟
    df1["晚上加班(分钟)"] = df1.apply(lambda x: (datetime.strptime(x["下班打卡时间"], "%Y-%m-%d %H:%M:%S") - datetime.strptime(
        x["上班打卡时间"], "%Y-%m-%d %H:%M:%S")).seconds / 60 - 60 if ((x["daytype"] == 1) | (x["daytype"] == 2)) else x[
        "晚上加班(分钟)"], axis=1)

    # 把分钟转换为小时
    df1["早上加班(小时)"] = df1["早上加班(分钟)"] / 60
    # df1["晚上加班(小时)"] = (df1["晚上加班(分钟)"] // 30) / 2
    df1["晚上加班(小时)"] = df1["晚上加班(分钟)"].apply(lambda x: x // 30 / 2 if x > 0 else 0)
    df1["晚上加班(小时)"] = df1.apply(lambda x: x["晚上加班(小时)"] + 0.5 if x["晚上加班(分钟)"] % 30 == 29 else x["晚上加班(小时)"], axis=1)

    print(df1.head().to_markdown())

    outfile = "".join(file.split(".")[:-1]) + "-IT.xlsx"
    df1.to_excel(outfile, index=False)

    print("已筛选IT人员")


def get_overtime(time):
    if int(time[14:16]) < 29:
        time = datetime.strptime(time[:14] + "00:00", '%Y-%m-%d %H:%M:%S')
        return time
    elif int(time[14:16]) < 59:
        time = datetime.strptime(time[:14] + "30:00", '%Y-%m-%d %H:%M:%S')
        return time
    else:
        time = datetime.strptime(time[:17] + "00", '%Y-%m-%d %H:%M:%S')
        return time


if __name__ == "__main__":
    find_user()
