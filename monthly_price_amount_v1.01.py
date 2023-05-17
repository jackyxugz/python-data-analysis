# _*_ coding: utf-8 _*_
# @Version: 1.0
# @File:
# @Description:
# @Date: 2022/7/19
# @Author: r-max

import os
import time
import tkinter as tk
from tkinter import filedialog
import tabulate

import numpy as np
import pandas as pd

root = tk.Tk()
root.withdraw()


## 将数据导出到 Excel Sheet
def export_excel(file_name, sheet_df):
    """

    将数据导出到 Excel Sheet
    :param file_name:
    :param sheet_df:
    """
    writer = pd.ExcelWriter(file_name)
    for sheet in sheet_df:
        sheet_df[sheet].to_excel(writer, sheet_name=sheet, index=False)
    writer.save()


## 主函数
def main():
    """
    main
    """

    # <editor-fold desc="开始文件选择">
    print("\n**********开始文件选择**********")

    print("\n-----1/2请选择需要处理的'财务底稿'文件:")
    manuscript_file_dir = filedialog.askopenfilename()
    # manuscript_file_dir = r'D:\f_r-max\contrast\bi_2022\monthly_price_amount\\麦凯莱\22年麦凯莱发出商品-结转成本0818.xlsx'
    if manuscript_file_dir is None or manuscript_file_dir.__len__() == 0:
        print("缺少'财务底稿'文件!!!")
        raise 'error'
    print("已选择的'财务底稿'文件:", manuscript_file_dir)

    print("\n-----2/2:请选择需要处理的'进销存明细'文件:")
    jxc_file_dir = filedialog.askopenfilename()
    # jxc_file_dir = r'D:\f_r-max\contrast\bi_2022\monthly_price_amount\\麦凯莱\22年进销存明细_麦凯莱_0817_04_hsc.xlsx'
    if jxc_file_dir is None or jxc_file_dir.__len__() == 0:
        print("缺少'进销存明细'文件!!!")
        raise 'error'
    print("已选择的'进销存明细'文件:", jxc_file_dir)

    print("\n**********结束文件选择**********")
    # </editor-fold>

    print("\n读取Excel数据中，请稍等...")

    # <editor-fold desc="处理财务底稿文件">
    print("\n**********开始处理财务底稿文件**********")

    manuscript_file_name = os.path.basename(manuscript_file_dir)
    print(str.format('读取-财务底稿数据...: {}', manuscript_file_name))
    df_manuscript = pd.read_excel(manuscript_file_dir, dtype={'产品参考': str, '产品条码': str})
    print(df_manuscript.head(100).to_markdown())
    if df_manuscript.empty or df_manuscript.shape[0] == 0:
        print('异常文件: ' + manuscript_file_name + ', 文件为空')
        raise 'error'
    print('文件: {}, 总共 {} 行'.format(manuscript_file_name, df_manuscript.shap[0]))
    df_manuscript["产品参考"] = df_manuscript["产品参考"].astype(str)
    df_manuscript["产品条码"] = df_manuscript["产品条码"].astype(str)

    print(df_manuscript[['产品参考', '产品条码']].to_markdown())

    df_manuscript.rename(
        columns={'税率': '销售税率',
                 '应收金额': '总应收款'},
        inplace=True)

    print(df_manuscript.head(5).to_markdown())

    df_manuscript['出货日期'] = df_manuscript['出货日期'].apply(
        lambda x: str(x).strip() if (str(x).__contains__(':')) else str(x).split(' ')[0].strip())
    df_manuscript["出货日期"] = pd.to_datetime(df_manuscript["出货日期"])
    # df_manuscript["订单日期"] = df_manuscript["订单日期"].apply(lambda x: str(x).replace('nan', '').replace('NaN', '').replace('NaT', '')[:10])
    df_manuscript["出货日期"] = df_manuscript["出货日期"].apply(
        lambda x: str(x).replace('nan', '').replace('NaN', '').replace('NaT', '')[:10])
    df_manuscript["日期-月份"] = df_manuscript["出货日期"].apply(lambda x: str(x)[:7])
    df_manuscript["产品参考"] = df_manuscript["产品参考"].apply(
        lambda x: x.replace('nan', '').replace('NaN', '').replace('NaT', '').replace('无', '').replace(
            '无SKU', ''))
    df_manuscript["产品参考"] = df_manuscript["产品参考"].apply(lambda x: "" if str(x) == '0' else str(x))
    df_manuscript["产品参考"] = df_manuscript["产品参考"].str.upper()
    df_manuscript["产品条码"] = df_manuscript["产品条码"].apply(
        lambda x: x.replace('nan', '').replace('NaN', '').replace('NaT', '').replace('无', '').replace(
            '无SKU', ''))
    df_manuscript["产品条码"] = df_manuscript["产品条码"].apply(lambda x: "" if str(x) == '0' else str(x))

    df_manuscript_all = df_manuscript.copy()

    print("\n**********结束处理财务底稿文件**********")
    print("\n**********开始处理ERP进销存明细数据**********")

    jxc_file_name = os.path.basename(jxc_file_dir)
    print(str.format('读取-ERP进销存明细数据...: {}', jxc_file_name))
    df_jxc = pd.read_excel(jxc_file_dir, sheet_name='Result 1', dtype={'产品参考': str, '产品条码': str})
    print(df_jxc.head(100).to_markdown())
    if df_jxc.empty or df_jxc.shape[0] == 0:
        print('异常文件: ' + jxc_file_name + ', 文件为空')
        raise 'error'
    print('文件: {}, 总共 {} 行'.format(jxc_file_name, df_jxc.shape[0]))

    # df_jxc.rename(
    #     columns={'产品参考': '产品参考'}, inplace=True)

    df_jxc = df_jxc[['本期开始时间', '产品参考', '产品条码', '月加权平均单价']]
    df_jxc["产品参考"] = df_jxc["产品参考"].astype(str)
    df_jxc["产品条码"] = df_jxc["产品条码"].astype(str)
    df_jxc['本期开始时间'] = df_jxc['本期开始时间'].apply(
        lambda x: str(x).strip() if (str(x).__contains__(':')) else str(x).split(' ')[0].strip())
    df_jxc['本期开始时间'] = pd.to_datetime(df_jxc['本期开始时间'])
    df_jxc['本期开始时间'] = df_jxc['本期开始时间'].apply(lambda x: str(x)[:10])
    df_jxc['日期-月份'] = df_jxc['本期开始时间'].apply(lambda x: str(x)[:7])
    df_jxc["产品参考"] = df_jxc["产品参考"].apply(
        lambda x: str(x).replace('nan', '').replace('NaN', '').replace('NaT', '').replace('无', '').replace('无SKU', ''))
    df_jxc["产品参考"] = df_jxc["产品参考"].apply(lambda x: "" if str(x) == '0' else str(x))
    # df_jxc["产品参考"] = df_jxc["产品参考"].apply(
    #     lambda x: str(int(x)).upper() if (str.isnumeric(x)) else str(x).upper())
    # df_jxc["产品参考"] = df_jxc["产品参考"].apply(lambda x: str(x)).upper()
    df_jxc["产品参考"] = df_jxc["产品参考"].str.upper()
    df_jxc["产品条码"] = df_jxc["产品条码"].apply(
        lambda x: str(x).replace('nan', '').replace('NaN', '').replace('NaT', '').replace('无', '').replace('无SKU', ''))
    df_jxc["产品条码"] = df_jxc["产品条码"].apply(lambda x: "" if str(x) == '0' else str(x))
    df_jxc = df_jxc[(df_jxc['月加权平均单价'] >= 0)].drop_duplicates().reset_index()

    df_jxc_all = df_jxc.copy()

    df_jxc_01 = df_jxc[['日期-月份', '产品条码', '月加权平均单价']]
    df_jxc_01 = df_jxc_01[(df_jxc_01['产品条码'] != '')]

    df_jxc_02 = df_jxc[['日期-月份', '产品参考', '月加权平均单价']]
    df_jxc_02 = df_jxc_02[(df_jxc_02['产品参考'] != '')]

    print("\n**********结束处理ERP进销存明细数据**********")
    # </editor-fold>

    # <editor-fold desc="处理整合文件数据">
    print("\n**********开始处理整合文件数据**********")

    df_manuscript['产品参考'] = df_manuscript['产品参考'].astype(str)
    df_manuscript['产品参考'] = df_manuscript['产品参考'].apply(
        lambda x: str(x).replace('nan', '').replace('NaN', '').replace('NaT', '').replace('无',
                                                                                          '').replace(
            '无SKU', ''))
    df_manuscript['产品条码'] = df_manuscript['产品条码'].astype(str)
    df_manuscript['产品条码'] = df_manuscript['产品条码'].apply(
        lambda x: str(x).replace('nan', '').replace('NaN', '').replace('NaT', '').replace('无',
                                                                                          '').replace(
            '无SKU', ''))
    df_manuscript_01 = df_manuscript[(df_manuscript['产品条码'] != '')]
    df_manuscript_02 = df_manuscript[((df_manuscript['产品条码'] == '') & (df_manuscript['产品参考'] != ''))]
    df_manuscript_03 = df_manuscript[((df_manuscript['产品条码'] == '') & (df_manuscript['产品参考'] == ''))]

    df_manuscript_compare_frames = []
    df_manuscript_compare_01 = pd.merge(df_manuscript_01, df_jxc_01, how="left", on=["日期-月份", "产品条码"])
    if df_manuscript_compare_01 is not None and df_manuscript_compare_01.shape[0] > 0:
        df_manuscript_compare_frames.append(df_manuscript_compare_01)
    df_manuscript_compare_02 = pd.merge(df_manuscript_02, df_jxc_02, how="left", on=["日期-月份", "产品参考"])
    if df_manuscript_compare_02 is not None and df_manuscript_compare_02.shape[0] > 0:
        df_manuscript_compare_frames.append(df_manuscript_compare_02)
    df_manuscript_compare_03 = df_manuscript_03
    if df_manuscript_compare_03 is not None and df_manuscript_compare_03.shape[0] > 0:
        df_manuscript_compare_frames.append(df_manuscript_compare_03)

    df_manuscript_compare = pd.concat(df_manuscript_compare_frames)
    print(df_manuscript_compare.head(10).to_markdown(), df_manuscript_compare.shape[0])

    df_manuscript_compare['完成数量'].fillna(0, inplace=True)
    # df_manuscript_compare['月加权平均单价'].fillna(0, inplace=True)
    df_manuscript_compare['月加权平均单价-金额'] = df_manuscript_compare.apply(
        lambda x: 0.00 if (str(x['月加权平均单价']) == '') else (x['完成数量'] * x['月加权平均单价']), axis=1)
    df_manuscript_compare['月加权平均单价-金额'].fillna(0, inplace=True)
    df_manuscript_compare['月加权平均单价-金额'] = round(df_manuscript_compare['月加权平均单价-金额'], 8)
    df_manuscript_compare = df_manuscript_compare.reset_index()
    del df_manuscript_compare['index']

    df_manuscript_compare['开票日期'] = df_manuscript_compare['开票日期'].astype(str)
    df_manuscript_compare_issued_01 = df_manuscript_compare[(df_manuscript_compare['开票日期'].str.contains('发出'))]
    df_manuscript_compare_unissued_01 = df_manuscript_compare[(df_manuscript_compare['开票日期'].str.contains('退货'))]
    df_manuscript_compare_inv = df_manuscript_compare[(
            (~df_manuscript_compare['开票日期'].str.contains('发出'))
            & (~df_manuscript_compare['开票日期'].str.contains('退货'))
            & (df_manuscript_compare['开票日期'] != ''))]
    df_manuscript_compare_inv['开票日期'] = pd.to_datetime(df_manuscript_compare_inv['开票日期'])
    df_manuscript_compare_inv['开票日期'] = df_manuscript_compare_inv['开票日期'].apply(lambda x: str(x)[:10])
    df_manuscript_compare_inv['开票日期-月份'] = df_manuscript_compare_inv['开票日期'].apply(lambda x: str(x)[:7])
    df_manuscript_compare_issued_02 = df_manuscript_compare_inv[(
            (df_manuscript_compare_inv['开票日期-月份'] != '')
            & (df_manuscript_compare_inv['开票日期-月份'] != df_manuscript_compare_inv['日期-月份']))]
    df_manuscript_compare_unissued_02 = df_manuscript_compare_inv[(
            (df_manuscript_compare_inv['开票日期-月份'] == '')
            | (df_manuscript_compare_inv['开票日期-月份'] == df_manuscript_compare_inv['日期-月份']))]

    df_manuscript_compare_issued = pd.concat([df_manuscript_compare_issued_01, df_manuscript_compare_issued_02])
    df_manuscript_compare_unissued = pd.concat([df_manuscript_compare_unissued_01, df_manuscript_compare_unissued_02])

    df_manuscript_compare_issued_compare = df_manuscript_compare_issued.groupby(
        ['销售客户', '产品参考', '产品条码', '产品名称', '日期-月份']).agg(
        {'完成数量': np.sum, '总应收款': np.sum, '月加权平均单价-金额': np.sum}).reset_index()
    df_manuscript_compare_unissued_compare = df_manuscript_compare_unissued.groupby(
        ['销售客户', '产品参考', '产品条码', '产品名称', '日期-月份']).agg(
        {'完成数量': np.sum, '总应收款': np.sum, '月加权平均单价-金额': np.sum}).reset_index()

    print('处理后, 数据条数: %s' % df_manuscript_compare.shape[0])

    df_manuscript_compare_issued['月加权平均单价'] = df_manuscript_compare_issued['月加权平均单价'].astype(str)
    df_manuscript_compare_issued['月加权平均单价'] = df_manuscript_compare_issued['月加权平均单价'].apply(
        lambda x: str(x).replace('nan', '').replace('NaN', '').replace('NaT', ''))
    df_manuscript_compare_unissued['月加权平均单价'] = df_manuscript_compare_unissued['月加权平均单价'].astype(str)
    df_manuscript_compare_unissued['月加权平均单价'] = df_manuscript_compare_unissued['月加权平均单价'].apply(
        lambda x: str(x).replace('nan', '').replace('NaN', '').replace('NaT', ''))

    print("\n**********结束处理整合文件数据**********")
    # </editor-fold>

    # <editor-fold desc="处理汇总文件数据">
    print("\n**********开始处理汇总文件数据**********")

    df_manuscript_compare_gather_detail = df_manuscript_compare.copy()
    df_manuscript_compare_gather_detail["出货日期-月份"] = df_manuscript_compare_gather_detail["出货日期"].apply(
        lambda x: str(x)[5:7])
    df_manuscript_compare_gather_detail["开票日期-月份"] = df_manuscript_compare_gather_detail["开票日期"].apply(
        lambda x: str(x)[5:7])
    df_manuscript_compare_gather_detail["出货日期-月份"] = df_manuscript_compare_gather_detail["出货日期-月份"].apply(
        lambda x: int(x) if str(x).isnumeric() else '0')
    df_manuscript_compare_gather_detail["开票日期-月份"] = df_manuscript_compare_gather_detail["开票日期-月份"].apply(
        lambda x: int(x) if str(x).isnumeric() else '0')
    df_manuscript_compare_gather_detail['出货日期-月份'].fillna(0, inplace=True)
    df_manuscript_compare_gather_detail['开票日期-月份'].fillna(0, inplace=True)
    df_manuscript_compare_gather_detail['出货日期-月份'] = df_manuscript_compare_gather_detail['出货日期-月份'].astype(
        int)
    df_manuscript_compare_gather_detail['开票日期-月份'] = df_manuscript_compare_gather_detail['开票日期-月份'].astype(
        int)
    df_manuscript_compare_gather_detail = df_manuscript_compare_gather_detail[
        ((df_manuscript_compare_gather_detail['出货日期-月份'] > 0)
         & (df_manuscript_compare_gather_detail['开票日期-月份'] > 0)
         & (df_manuscript_compare_gather_detail['出货日期-月份'] != df_manuscript_compare_gather_detail[
                    '开票日期-月份']))]

    df_manuscript_compare_gather_detail.rename(columns={"总应收款": "总应收款（人民币）"}, inplace=True)
    df_manuscript_compare_gather = df_manuscript_compare_gather_detail[
        ['日期-月份', '销售客户', '产品条码', '产品名称', '完成数量', '总应收款（人民币）', '月加权平均单价',
         '月加权平均单价-金额']]
    df_manuscript_compare_gather = df_manuscript_compare_gather.groupby(
        ['日期-月份', '销售客户', '产品条码', '产品名称', '月加权平均单价']).agg(
        {'完成数量': np.sum, '总应收款（人民币）': np.sum, '月加权平均单价-金额': np.sum}
    ).reset_index()
    df_manuscript_compare_gather = df_manuscript_compare_gather[
        ['日期-月份', '销售客户', '产品条码', '产品名称', '完成数量', '总应收款（人民币）', '月加权平均单价',
         '月加权平均单价-金额']]

    print("\n**********结束处理汇总文件数据**********")
    # </editor-fold>

    # <editor-fold desc="导出数据">
    print("\n**********开始导出数据**********")

    print('\n导出本次处理文件数据...')
    jxc_file_name_arr = jxc_file_name.split('_', 2)
    jxc_file_name_tmp = jxc_file_name_arr[2].replace('.xlsx', '').replace('.xls', '').replace('.csv', '')
    manuscript_file_name_tmp = manuscript_file_name.replace('.xlsx', '').replace('.xls', '').replace('.csv', '')
    save_file_name = str.format('处理结果-{}-{}-{}.xlsx', jxc_file_name_tmp,
                                manuscript_file_name_tmp, time.strftime('%H%M', time.localtime()))
    export_dir = os.path.join(os.path.dirname(manuscript_file_dir), 'result')
    if not os.path.exists(export_dir):
        os.mkdir(export_dir)
    save_file = os.path.join(export_dir, save_file_name)

    with pd.ExcelWriter(save_file) as xlsx:
        df_manuscript_compare_issued_compare.to_excel(xlsx, sheet_name="发出商品-金额-汇总")
        df_manuscript_compare_unissued_compare.to_excel(xlsx, sheet_name="非发出商品-金额-汇总")
        df_manuscript_compare_issued.to_excel(xlsx, sheet_name="发出商品明细")
        df_manuscript_compare_unissued.to_excel(xlsx, sheet_name="非发出商品明细")
        df_manuscript_compare.to_excel(xlsx, sheet_name="月加权单价-金额")
        # df_manuscript_compare_gather.to_excel(xlsx, sheet_name="发出商品差异汇总")
        # df_manuscript_compare_gather_detail.to_excel(xlsx, sheet_name="发出商品差异明细")
    print(str.format('\n处理结果文件: {}', save_file))

    # if df_manuscript_compare is not None and df_manuscript_compare.shape[0] >= 1000000:
    #     year_month_arr = df_manuscript_compare['出货日期'].apply(lambda x: str(x)[:7])
    #     year_month_arr = year_month_arr.drop_duplicates().sort_values().tolist()
    #     # if year_month_arr is not None and year_month_arr.__contains__('nan'):
    #     #     year_month_arr.remove('nan')
    #     for year_month in year_month_arr:
    #         print(year_month)
    #         sheets = {}
    #         save_file_tmp = save_file.replace('处理结果', ('处理结果(%s)' % year_month))
    #         df_manuscript_compare['出货日期'] = df_manuscript_compare['出货日期'].astype(str)
    #         df_manuscript_compare_tmp = df_manuscript_compare[(df_manuscript_compare['出货日期'].str.contains(year_month))]
    #         if df_manuscript_compare_tmp.shape[0] < 1000000:
    #             sheets[("(%s)月加权单价-金额" % year_month)] = df_manuscript_compare_tmp
    #         else:
    #             for i, temp in enumerate(np.array_split(df_manuscript_compare_tmp, 2)):
    #                 sheets[("(%s)月加权单价-金额-%s" % (year_month, i))] = temp
    #
    #         df_manuscript_compare_gather['日期-月份'] = df_manuscript_compare_gather['日期-月份'].astype(str)
    #         df_manuscript_compare_gather_tmp = df_manuscript_compare_gather[
    #             (df_manuscript_compare_gather['日期-月份'].str.contains(year_month))]
    #         if df_manuscript_compare_gather_tmp.shape[0] < 1000000:
    #             sheets["发出商品差异汇总"] = df_manuscript_compare_gather_tmp
    #         else:
    #             for i, temp in enumerate(np.array_split(df_manuscript_compare_gather_tmp, 2)):
    #                 sheets["发出商品差异汇总-%s" % i] = temp
    #
    #         df_manuscript_compare_gather_detail['出货日期'] = df_manuscript_compare_gather_detail['出货日期'].astype(str)
    #         df_manuscript_compare_gather_detail_tmp = df_manuscript_compare_gather_detail[
    #             (df_manuscript_compare_gather_detail['出货日期'].str.contains(year_month))]
    #         if df_manuscript_compare_gather_detail_tmp.shape[0] < 1000000:
    #             sheets["发出商品差异明细"] = df_manuscript_compare_gather_detail_tmp
    #         else:
    #             for i, temp in enumerate(np.array_split(df_manuscript_compare_gather_detail_tmp, 2)):
    #                 sheets["发出商品差异明细-%s" % i] = temp
    #
    #         export_excel(save_file_tmp, sheets)
    #
    #     print(str.format('\n处理结果文件夹: {}', export_dir))
    # else:
    #     with pd.ExcelWriter(save_file) as xlsx:
    #         df_manuscript_compare_issued_compare.to_excel(xlsx, sheet_name="发出商品-金额-汇总")
    #         df_manuscript_compare_unissued_compare.to_excel(xlsx, sheet_name="非发出商品-金额-汇总")
    #         df_manuscript_compare_issued.to_excel(xlsx, sheet_name="发出商品明细")
    #         df_manuscript_compare_unissued.to_excel(xlsx, sheet_name="非发出商品明细")
    #         df_manuscript_compare.to_excel(xlsx, sheet_name="月加权单价-金额")
    #         df_manuscript_compare_gather.to_excel(xlsx, sheet_name="发出商品差异汇总")
    #         df_manuscript_compare_gather_detail.to_excel(xlsx, sheet_name="发出商品差异明细")
    #
    #     print(str.format('\n处理结果文件: {}', save_file))
    # 找出两表的差异
    print("\n**********结束导出数据**********")
    # </editor-fold>

    print("=====开始处理进销存和财务底稿差异=====")
    find_difference(df_manuscript_all, df_jxc_all)


def find_difference(df_manuscript_all, df_jxc_all):
    df_manuscript_tmp = df_manuscript_all[['日期-月份', '产品参考', '产品条码']]
    # df_manuscript_tmp.columns = ['出货日期-月份', '产品参考', '产品条码']
    df_jxc_tmp = df_jxc_all[['日期-月份', '产品参考', '产品条码']]

    df_manuscript_bar = df_manuscript_tmp[['日期-月份', '产品条码']].drop_duplicates().reset_index()
    df_manuscript_bar['TMP_INDEX'] = df_manuscript_bar.apply(lambda x: '%s%s' % (x['日期-月份'], x['产品条码']), axis=1)
    df_jxc_bar = df_jxc_tmp[['日期-月份', '产品条码']].drop_duplicates().reset_index()
    df_jxc_bar['TMP_INDEX'] = df_jxc_bar.apply(lambda x: '%s%s' % (x['日期-月份'], x['产品条码']), axis=1)

    bar_indexes = df_manuscript_bar['TMP_INDEX'].tolist()
    df_manuscript_bar_diff = df_jxc_bar[(~df_jxc_bar['TMP_INDEX'].isin(bar_indexes))]
    bar_indexes = df_jxc_bar['TMP_INDEX'].tolist()
    df_jxc_bar_diff = df_manuscript_bar[(~df_manuscript_bar['TMP_INDEX'].isin(bar_indexes))]

    input()

    #
    #
    #
    # df_manuscript_caiwu = df_manuscript_all
    # df_jxc_all = df_jxc_all
    #
    # print("\n财务底稿：")
    # print(df_manuscript_caiwu.head(5).to_markdown())
    # print("\n进销存底稿：")
    # print(df_jxc_all.head(5).to_markdown())
    #
    # df_manuscript_caiwu["产品参考"] = df_manuscript_caiwu["产品参考"].astype(str)
    # df_manuscript_caiwu["产品条码"] = df_manuscript_caiwu["产品条码"].astype(str)
    # df_jxc_all["产品参考"] = df_jxc_all["产品参考"].str.strip()
    # df_jxc_all["产品条码"] = df_jxc_all["产品条码"].str.strip()
    #
    # lst_order_no = df_manuscript_caiwu['产品参考'].drop_duplicates().tolist()
    # df_chayi_caiwu = df_manuscript_caiwu[(~df_manuscript_caiwu['产品参考'].isin(lst_order_no))]
    # print("财务底稿差异：")
    # print(df_chayi_caiwu.head(5).to_markdown())
    #
    # df_chayi_caiwu.to_excel(r"C:\Users\sjit36\Desktop\tob核对\财务底稿差异_20220820.xlsx")
    #
    # lst_order_no = df_jxc_all['产品参考'].drop_duplicates().tolist()
    # df_chayi_jxc = df_jxc_all[(~df_jxc_all['产品参考'].isin(lst_order_no))]
    # print("进销存差异：")
    # print(df_chayi_caiwu.head(5).to_markdown())
    #
    # df_chayi_jxc.to_excel(r"C:\Users\sjit36\Desktop\tob核对\进销存底稿差异_20220820.xlsx")


if __name__ == '__main__':

    print('\n执行开始\n\n开始时间：', time.strftime('%Y-%m-%d %H:%M:%S', time.localtime()))
    start = time.time()

    try:
        main()

    except Exception as ex:
        print('\n\n程序错误:')
        print(ex)

    print('\n**********OK**********\n')

    print('\n执行结束\n结束时间:', time.strftime('%Y-%m-%d %H:%M:%S', time.localtime()))
    end = time.time()
    print('执行用时:', '%.2f' % (end - start), '秒\n')

    print('-------如需关闭窗口，请回车-------')
    input("按任意键退出......")
