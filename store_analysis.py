# _*_ coding: utf-8 _*_
# @Version: 1.0
# @Description:SKU库龄分析
# @Date: 2022/10/31
# @Author: Jacky.Xu

import os
import time
import tkinter as tk
from tkinter import filedialog
import numpy as np
import pandas as pd
import sys

root = tk.Tk()
root.withdraw()

# 设置文件对话框会显示的文件类型
my_filetypes = [('text excel files', '.xlsx'), ('all excel files', '.xls'), ('all excel files', '.csv')]

# 获取文件
def get_file(title_message):
    # 请求选择文件
    filename = filedialog.askopenfilename(initialdir=os.getcwd(),
                                          title=title_message,
                                          filetypes=my_filetypes)

    if len(filename) == 0:
        sys.exit()

    print("你选择的文件名是：", filename)
    return filename


# 主函数
def store_analysis(my_year):
    # 读取文件
    filename = get_file("请选择你要处理的文件:")
    print("\n------正在读取{}年进销存数据------".format(my_year))
    df = pd.read_excel(filename, sheet_name='{}年香港仓进销存'.format(my_year))

    # 清洗数据
    # ['本期开始时间', '仓库名称', '产品ID', '产品条码', '内部参考', '产品名称', '计量单位', '期初数量', '期初单价',
    # '期初金额', '采购入库数量', '采购入库金额', '销售出库数量', '销售出库金额', '销售退货入库数量',
    # '销售退货入库金额', '采购退货出库数量', '采购退货出库金额', '其他入库数量', '其他入库金额', '其他出库数量','其他出库金额',
    # '本期入库数量', '本期入库金额', '本期出库数量', '本期出库金额', '期末数量', '期末单价','月加权平均单价', '期末金额']
    df["本期开始时间"].astype(str)
    df["产品ID"].astype(str)
    df["产品条码"].astype(str)
    df["内部参考"].astype(str)
    df["产品名称"].astype(str)
    df["计量单位"].astype(str)
    df['期初数量'].astype(int)
    df['期末数量'].astype(int)
    df['期初金额'].astype(float)
    df['期末金额'].astype(float)
    df['期初数量'].fillna(0, inplace=True)
    df['期末数量'].fillna(0, inplace=True)
    df['期初金额'].fillna(0.00, inplace=True)
    df['期末金额'].fillna(0.00, inplace=True)
    df['销售出库金额'].fillna(0, inplace=True)
    df['其他出库金额'].fillna(0, inplace=True)

    df = df[~df['内部参考'].isna()]  # 去掉'内部参考'为空的值

    df_desc=df.groupby(['内部参考', '产品ID', '产品条码', '产品名称', '计量单位'])["期末金额"].sum()
    df_desc=pd.DataFrame(df_desc).reset_index()

    # 取原表的期初数量和期初金额
    df_num_money_qichu = df[['本期开始时间', '内部参考', '期初数量', '期初金额']]

    df_num_money_qichu = df_num_money_qichu[df["本期开始时间"].str.contains("-01-01")]
    df_num_money_qichu = df_num_money_qichu.groupby('内部参考').sum().reset_index(drop=False)
    print('1111')
    print(df_num_money_qichu.head(5).to_markdown())

    # 取原表的期末数量和期末金额
    df_num_money_qimo = df[['本期开始时间', '内部参考', '期末数量', '期末金额']]
    df_num_money_qimo = df_num_money_qimo[df["本期开始时间"].str.contains("-12-01")]
    df_num_money_qimo = df_num_money_qimo.groupby('内部参考').sum().reset_index(drop=False)
    print('2222')
    print(df_num_money_qimo.head(5).to_markdown())

    # 取原表的销售出库金额和其他出库金额的汇总
    df_sale_qt_money = df[['本期开始时间', '内部参考', '销售出库金额', '其他出库金额']]
    df_sale_qt_money = df_sale_qt_money.groupby('内部参考').sum().reset_index(drop=False)
    print('3333')
    print(df_sale_qt_money.head(5).to_markdown())

    # 合并所有数据
    df_result = df_num_money_qichu.merge(df_num_money_qimo[['内部参考','期末数量', '期末金额']], how='left', on='内部参考')
    df_result = df_result.merge(df_sale_qt_money[['内部参考','销售出库金额', '其他出库金额']], how='left', on='内部参考')
    df_result = pd.merge(df_result,df_desc[['内部参考', '产品ID', '产品条码', '产品名称', '计量单位']],how='left',on='内部参考')
    print('4444')
    print(df_result.head(5).to_markdown())

    # 去除金额为0的记录
    df_result = df_result.loc[~((df_result['销售出库金额'] == 0) & (df_result['其他出库金额'] == 0) & (df_result['期初金额'] == 0) & (
            df_result['期末金额'] == 0) & (df_result['期初数量'] == 0) & (df_result['期末数量'] == 0))]
    print('5555')
    print(df_result.shape[0])

    #  计算报表所需要的指标值：'存货平均余额','出库成本','周转率','周转天数'
    df_result['天数'] = 365
    df_result['存货平均余额'] = (df_result['期初金额'] + df_result['期末金额']) / 2
    df_result['出库成本'] = df_result['销售出库金额'] + df_result['其他出库金额']
    df_result['周转率'] = df_result['出库成本'] / df_result['存货平均余额']
    df_result['周转天数'] = df_result['天数'] / df_result['周转率']
    print('6666')
    print(df_result.head(5).to_markdown())
    print(df_result.shape[0])
    df_result.to_excel('{}年存货周转率汇总表.xlsx'.format(my_year),sheet_name='{}年存货周转率汇总表'.format(my_year))


if __name__ == '__main__':
    print('\n开始时间：', time.strftime('%Y-%m-%d %H:%M:%S', time.localtime()))
    start = time.time()

    try:
        # store_analysis(19)
        # store_analysis(20)
        store_analysis(21)
    except Exception as ex:
        print('\n程序错误信息:')
        print(ex)

    print('\n程序运行完毕！结束时间：', time.strftime('%Y-%m-%d %H:%M:%S', time.localtime()))

    end = time.time()
    print('\n运行总用时:', '%.2f' % (end - start), '秒\n')

    print('按任意键退出程序。。。')
    input()
    sys.exit()
