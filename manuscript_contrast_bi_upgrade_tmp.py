# _*_ coding: utf-8 _*_
# @Version: 1.0
# @File:
# @Description:
# @Date: 2022/6/29
# @Author: Jacky

import os
import time
import tkinter as tk
from tkinter import filedialog

import numpy as np
import pandas as pd
from tabulate import tabulate

root = tk.Tk()
root.withdraw()


# 读取数据表格
def read_excel(file_dirs, skip_rows):
    """
    读取数据表格
    :param file_dirs:
    :param skip_rows:
    :return:
    """
    try:
        files_frames = []
        for dir_file in file_dirs:
            if skip_rows > 0:
                df_file = pd.read_excel(dir_file, skiprows=skip_rows)
            else:
                df_file = pd.read_excel(dir_file)
            if df_file.empty or df_file.shape[0] == 0:
                print('异常文件: ' + dir_file + ', 文件为空')
                # continue
            files_frames.append(df_file)
            print('文件: {}, 总共 {} 行'.format(dir_file, df_file.shape[0]))

    except Exception as e:
        print('load fn excel data failure !')
        print(e)
        raise 'error'

    df_files = pd.concat(files_frames)
    return df_files


# 对比数据
def contrast_file(pd_cg_cur, pd_two_b_cur, pd_dp_cur, pd_relation_cur, pd_erp_cur, save_file_cur):
    """
    对比数据
    :param pd_cg_cur:
    :param pd_two_b_cur:
    :param pd_dp_cur:
    :param pd_relation_cur:
    :param pd_erp_cur:
    :param save_file_cur:
    """
    # 筛选本次测算公司主体的数据
    subject_current = pd_erp_cur['公司'].drop_duplicates().tolist()
    print(str.format('\n本次测算的公司主体: {}', subject_current))
    pd_cg_current = pd_cg_cur[(pd_cg_cur['公司主体'].isin(subject_current))]
    pd_two_b_current = pd_two_b_cur[(pd_two_b_cur['主体'].isin(subject_current))]
    pd_dp_current = pd_dp_cur[(pd_dp_cur['主体'].isin(subject_current))]

    print("\n正在计算中，请稍等...")

    # 处理-采购含税金额-供应商
    print('\n处理-采购含税金额-供应商...')
    cg = pd_cg_current.copy()
    cg["送货日期"] = pd.to_datetime(cg["送货日期"])
    cg["送货日期"] = cg["送货日期"].apply(lambda x: str(x)[:7])
    cg = cg[["公司主体", "供应商名称", "送货日期", "采购含税金额-供应商"]]
    cg["采购含税金额-供应商"].fillna(0, inplace=True)
    cg["供应商名称"] = cg["供应商名称"].astype(str).apply(lambda x: x.strip())
    cg = cg.groupby(["公司主体", "供应商名称", "送货日期"]).agg({"采购含税金额-供应商": np.sum}).reset_index()

    # 处理-交易总额
    print('\n处理-交易总额...')
    cg1 = pd_cg_current.copy()
    cg1["送货日期"] = pd.to_datetime(cg1["送货日期"])
    cg1["送货日期"] = cg1["送货日期"].apply(lambda x: str(x)[:7])
    cg1["交易总额-对应子公司销售给“麦凯莱“"].fillna(0, inplace=True)
    cg1 = cg1[["公司主体", "送货日期", "交易总额-对应子公司销售给“麦凯莱“"]]
    cg1 = cg1.groupby(["公司主体", "送货日期"]).agg({"交易总额-对应子公司销售给“麦凯莱“": np.sum}).reset_index()

    # 处理-2B表数据
    print('\n处理-2B表数据...')
    two_b = pd_two_b_current.copy()
    two_b = two_b[["出货日期", "购买方名称-销售客户", "销售出库-价税合计", "类型", "业务类型"]]
    two_b["出货日期"] = pd.to_datetime(two_b["出货日期"])
    two_b["出货日期"] = two_b["出货日期"].apply(lambda x: str(x)[:7])
    two_b["销售出库-价税合计"].fillna(0, inplace=True)
    two_b["类型"].fillna('2B', inplace=True)
    two_b['类型'] = two_b['类型'].astype(str)
    two_b["2B-销售出库-价税合计"] = 0
    two_b["2B发出商品-销售出库-价税合计"] = 0
    if not two_b.empty:
        two_b["2B-销售出库-价税合计"] = two_b.apply(
            lambda x: x['销售出库-价税合计'] if (x['类型'].__eq__('2B')) else x["2B-销售出库-价税合计"], axis=1)
        two_b["2B发出商品-销售出库-价税合计"] = two_b.apply(
            lambda x: x['销售出库-价税合计']
            if (x['类型'].__contains__('2B') & x['类型'].__contains__('发出'))
            else x["2B发出商品-销售出库-价税合计"],
            axis=1)
    two_b = two_b.groupby(["出货日期", "购买方名称-销售客户"]).agg(
        {"销售出库-价税合计": np.sum, "2B-销售出库-价税合计": np.sum, "2B发出商品-销售出库-价税合计": np.sum}
    ).reset_index()

    # 处理-店铺销售数据
    print('\n处理-店铺销售数据...')
    dp = pd_dp_current.copy()
    dp = dp[["主体", "平台", "财务店铺名称", "回款合计", '减21发出商品后', '加2022年1-6月发出商品']]
    dp["回款合计"].fillna(0, inplace=True)
    dp["减21发出商品后"].fillna(0, inplace=True)
    dp["加2022年1-6月发出商品"].fillna(0, inplace=True)
    dp['主体'] = dp['主体'].astype(str)
    dp["平台"] = dp["平台"].apply(lambda x: '枫叶' if (str(x).__contains__('枫叶')) else x)
    dp.rename(columns={"减21发出商品后": "2C商品回款", "加2022年1-6月发出商品": "2C发出商品回款"}, inplace=True)
    dp = dp.groupby(["主体", "平台", "财务店铺名称"]).agg(
        {"回款合计": np.sum, "2C商品回款": np.sum, "2C发出商品回款": np.sum}).reset_index()
    dp["财务店铺名称"] = dp["财务店铺名称"].str.lower()
    dp["主体"] = dp["主体"].apply(lambda x: x.replace("(", "（").strip())
    dp["主体"] = dp["主体"].apply(lambda x: x.replace(")", "）").strip())

    # 处理-ERP移动库存-采购金额
    print('\n处理-ERP移动库存-采购金额...')
    erp = pd_erp_cur.copy()
    erp = erp[['公司', '采购供应商', '日期', '业务类型', '采购/退货(含税)金额(RMB)', '加工费(含税)金额']]
    erp['业务类型'] = erp['业务类型'].astype(str)
    erp = erp[erp['业务类型'].str.contains('采购')]
    erp['采购供应商'] = erp['采购供应商'].astype(str)
    erp = erp[~erp['采购供应商'].str.contains('深圳市麦凯莱科技有限公司')]
    erp["日期"] = pd.to_datetime(erp["日期"])
    erp["日期"] = erp["日期"].apply(lambda x: str(x)[:7])
    erp["采购/退货(含税)金额(RMB)"].fillna(0, inplace=True)
    erp["加工费(含税)金额"].fillna(0, inplace=True)
    erp["ERP采购含税金额-供应商"] = erp["采购/退货(含税)金额(RMB)"] + erp["加工费(含税)金额"]
    erp = erp[['公司', '采购供应商', '日期', 'ERP采购含税金额-供应商']]
    erp["采购供应商"] = erp["采购供应商"].astype(str).apply(lambda x: x.strip())
    erp = erp.groupby(['公司', '采购供应商', '日期']).agg({"ERP采购含税金额-供应商": np.sum}).reset_index()
    erp.columns = ["公司主体", "供应商名称", "送货日期", "ERP采购含税金额-供应商"]

    # 处理-ERP移动库存-交易总额
    print('\n处理-ERP移动库存-交易总额...')
    erp1 = pd_erp_cur.copy()
    erp1 = erp1[
        ['公司', '采购供应商', '日期', '业务类型', '销售客户', '加工费(含税)金额', '销售已交货含税金额(RMB)']]
    erp1 = erp1[erp1['业务类型'].str.contains('销售')]
    erp1['销售客户'] = erp1['销售客户'].astype(str)
    erp1 = erp1[erp1['销售客户'].str.contains('深圳市麦凯莱科技有限公司')]
    erp1["日期"] = pd.to_datetime(erp1["日期"])
    erp1["日期"] = erp1["日期"].apply(lambda x: str(x)[:7])
    erp1["销售已交货含税金额(RMB)"].fillna(0, inplace=True)
    erp1["加工费(含税)金额"].fillna(0, inplace=True)
    erp1["ERP交易总额-对应子公司销售给“麦凯莱“"] = erp1["销售已交货含税金额(RMB)"] + erp1["加工费(含税)金额"]
    erp1 = erp1.groupby(['公司', '日期']).agg({"ERP交易总额-对应子公司销售给“麦凯莱“": np.sum}).reset_index()
    erp1.columns = ['公司主体', '送货日期', 'ERP交易总额-对应子公司销售给“麦凯莱“']

    # 处理-ERP移动库存-2B
    print('\n处理-ERP移动库存-2B...')
    erp_two_b = pd_erp_cur.copy()
    erp_two_b = erp_two_b[['日期', '类型', "业务类型", '销售客户', '销售已交货含税金额(RMB)']]
    erp_two_b['销售客户'] = erp_two_b['销售客户'].astype(str)
    erp_two_b = erp_two_b[~erp_two_b['销售客户'].str.contains('-')]
    erp_two_b = erp_two_b[~erp_two_b['销售客户'].str.contains("麦凯莱")]
    erp_two_b["日期"] = pd.to_datetime(erp_two_b["日期"])
    erp_two_b["日期"] = erp_two_b["日期"].apply(lambda x: str(x)[:7])
    erp_two_b["销售已交货含税金额(RMB)"].fillna(0, inplace=True)
    # erp_two_b["类型"].fillna('2B', inplace = True)
    erp_two_b['类型'] = erp_two_b['类型'].astype(str)
    erp_two_b["ERP2B-销售出库-价税合计"] = 0
    erp_two_b["ERP2B发出商品-销售出库-价税合计"] = 0
    if not erp_two_b.empty:
        erp_two_b["ERP2B-销售出库-价税合计"] = erp_two_b.apply(
            lambda x: x['销售已交货含税金额(RMB)']
            if (x['类型'].__eq__('2B'))
            else x["ERP2B-销售出库-价税合计"], axis=1)
        erp_two_b["ERP2B发出商品-销售出库-价税合计"] = erp_two_b.apply(
            lambda x: x['销售已交货含税金额(RMB)']
            if (x['类型'].__contains__('2B') & x['类型'].__contains__('发出'))
            else x["ERP2B发出商品-销售出库-价税合计"],
            axis=1)
    erp_two_b = erp_two_b.groupby(['日期', '销售客户']).agg(
        {"销售已交货含税金额(RMB)": np.sum, "ERP2B-销售出库-价税合计": np.sum,
         "ERP2B发出商品-销售出库-价税合计": np.sum}
    ).reset_index()
    erp_two_b.columns = ["出货日期", "购买方名称-销售客户", "ERP销售出库-价税合计", 'ERP2B-销售出库-价税合计',
                         'ERP2B发出商品-销售出库-价税合计']

    # 处理-ERP移动库存-店铺销售
    print('\n处理-ERP移动库存-店铺销售...')
    erp_shop = pd_erp_cur.copy()
    erp_shop = erp_shop[['公司', '类型', '销售客户', '销售已交货含税金额(RMB)', '完成金额(RMB)', '销售税率']]
    # 销售(只含2C, 不含2B和Mega)
    erp_shop['销售客户'] = erp_shop['销售客户'].astype(str)
    erp_shop = erp_shop[erp_shop['销售客户'].str.contains('-')]
    erp_shop["平台"] = erp_shop["销售客户"].apply(lambda x: x.split("-")[0])
    erp_shop["店铺"] = erp_shop["销售客户"].apply(lambda x: x.split("-")[1])
    erp_shop["平台"] = erp_shop["平台"].apply(lambda x: '枫叶' if (str(x).__contains__('枫叶')) else x)
    erp_shop = erp_shop[['公司', '类型', '平台', '店铺', '销售已交货含税金额(RMB)', '完成金额(RMB)', '销售税率']]
    erp_shop["销售已交货含税金额(RMB)"].fillna(0, inplace=True)
    erp_shop["完成金额(RMB)"].fillna(0, inplace=True)
    erp_shop["ERP2C商品回款"] = 0
    erp_shop["ERP2C发出商品回款"] = 0
    if not erp_shop.empty:
        erp_shop["ERP2C商品回款"] = erp_shop.apply(
            lambda x: x['销售已交货含税金额(RMB)'] if (str(x['类型']).__eq__('2C')) else 0, axis=1)
        erp_shop["ERP2C发出商品回款"] = erp_shop.apply(
            lambda x: x['销售已交货含税金额(RMB)']
            if (str(x['类型']).__contains__('2C') & str(x['类型']).__contains__('发出'))
            else 0,
            axis=1)
    erp_shop = erp_shop.groupby(['公司', '平台', '店铺', '销售税率']).agg(
        {"销售已交货含税金额(RMB)": np.sum, "ERP2C商品回款": np.sum, "ERP2C发出商品回款": np.sum,
         "完成金额(RMB)": np.sum}
    ).reset_index()
    erp_shop['ERP完成金额(公式计算)'] = 0
    if not erp_shop.empty:
        erp_shop['ERP完成金额(公式计算)'] = erp_shop.apply(
            lambda x: (float(x['销售已交货含税金额(RMB)']) / 1.03 * 0.7)
            if (str(x['销售税率']).__contains__('税收3%') or str(x['销售税率']).__contains__('税收3％'))
            else x['ERP完成金额(公式计算)'],
            axis=1)
        erp_shop['ERP完成金额(公式计算)'] = erp_shop.apply(
            lambda x: (float(x['销售已交货含税金额(RMB)']) / 1.13 * 0.7)
            if (str(x['销售税率']).__contains__('税收13%') or str(x['销售税率']).__contains__('税收13％'))
            else x['ERP完成金额(公式计算)'],
            axis=1)
    erp_shop['ERP完成金额差异'] = (erp_shop["ERP完成金额(公式计算)"] - erp_shop["完成金额(RMB)"])
    erp_shop['ERP完成金额差异'].fillna(0, inplace=True)
    erp_shop['ERP完成金额差异'] = round(erp_shop['ERP完成金额差异'], 8)
    erp_shop = erp_shop.groupby(['公司', '平台', '店铺']).agg(
        {"销售已交货含税金额(RMB)": np.sum, "ERP2C商品回款": np.sum, "ERP2C发出商品回款": np.sum,
         "ERP完成金额(公式计算)": np.sum, "完成金额(RMB)": np.sum, "ERP完成金额差异": np.sum}
    ).reset_index()
    erp_shop.columns = ["主体", "平台", "财务店铺名称", "ERP回款合计", "ERP2C商品回款", "ERP2C发出商品回款",
                        "ERP完成金额(公式计算)", "ERP完成金额", "ERP完成金额差异"]
    erp_shop["财务店铺名称"] = erp_shop["财务店铺名称"].str.lower()
    erp_shop['主体'] = erp_shop['主体'].astype(str)
    erp_shop["主体"] = erp_shop["主体"].apply(lambda x: x.replace("(", "（").strip())
    erp_shop["主体"] = erp_shop["主体"].apply(lambda x: x.replace(")", "）").strip())

    # 处理-ERP移动库存-店铺销售成本
    print('\n处理-ERP移动库存-店铺销售成本...')
    erp_shop_cost = pd_erp_cur.copy()
    erp_shop_cost = erp_shop_cost[
        ['公司', '采购供应商', '销售客户', '业务类型', '销售税率', '采购/退货(含税)金额(原币别)', '加工费(含税)金额']]
    erp_shop_cost['采购/退货(含税)金额(原币别)'].fillna(0, inplace=True)
    erp_shop_cost['加工费(含税)金额'].fillna(0, inplace=True)
    erp_shop_cost['销售税率'] = erp_shop_cost['销售税率'].astype(str).replace(['nan', 'NaN'], '')
    tax_arr = erp_shop_cost['销售税率']
    erp_shop_cost['业务类型'] = erp_shop_cost['业务类型'].astype(str)
    erp_shop_cost = erp_shop_cost[erp_shop_cost['业务类型'].str.contains('采购')]
    # 是否含外采
    # erp_shop_cost['采购供应商'] = erp_shop_cost['采购供应商'].astype(str)
    # erp_shop_cost = erp_shop_cost[erp_shop_cost['采购供应商'].str.contains('深圳市麦凯莱科技有限公司')]
    erp_shop_cost['采购/退货(含税)金额(原币别)'] = (
            erp_shop_cost['采购/退货(含税)金额(原币别)'] + erp_shop_cost['加工费(含税)金额']
    )
    erp_shop_cost = erp_shop_cost[['公司', '采购/退货(含税)金额(原币别)']]
    erp_shop_cost = erp_shop_cost.groupby(['公司']).agg({'采购/退货(含税)金额(原币别)': np.sum}).reset_index()
    erp_shop_cost['3%采购/退货(含税)金额(原币别)'] = 0
    erp_shop_cost['13%采购/退货(含税)金额(原币别)'] = 0
    tax_current = 0
    if not tax_arr.empty:
        tax_arr = tax_arr.drop_duplicates().tolist()
        if tax_arr.__contains__(''):
            tax_arr.remove('')
        if tax_arr is not None:
            if tax_arr.__len__() == 1:
                tax = str(tax_arr[0])
                if tax.__contains__('税收3%') or tax.__contains__('税收3％'):
                    tax_current = 1.03
                    erp_shop_cost['3%采购/退货(含税)金额(原币别)'] = erp_shop_cost['采购/退货(含税)金额(原币别)']
                if tax.__contains__('税收13%') or tax.__contains__('税收13％'):
                    tax_current = 1.13
                    erp_shop_cost['13%采购/退货(含税)金额(原币别)'] = erp_shop_cost['采购/退货(含税)金额(原币别)']
            else:
                tax_current = 2
        else:
            tax_current = 1
    erp_shop_cost = erp_shop_cost[
        ['公司', '采购/退货(含税)金额(原币别)', '3%采购/退货(含税)金额(原币别)', '13%采购/退货(含税)金额(原币别)']]
    erp_shop_cost.columns = ["主体", 'ERP采购含税金额合计', 'ERP采购含税(3%)金额合计', 'ERP采购含税(13%)金额合计']
    erp_shop_cost['主体'] = erp_shop_cost['主体'].astype(str)
    erp_shop_cost["主体"] = erp_shop_cost["主体"].apply(lambda x: x.replace("(", "（").strip())
    erp_shop_cost["主体"] = erp_shop_cost["主体"].apply(lambda x: x.replace(")", "）").strip())

    # 处理-ERP移动库存-外采数量透视
    print('\n处理-ERP移动库存-外采数量透视...')
    erp_outsourcing_qty = pd_erp_cur.copy()
    erp_outsourcing_qty = erp_outsourcing_qty[
        ['公司', '日期', '业务类型', '产品参考', '产品条码', '产品类别', '采购供应商', '销售客户', '完成数量']]
    erp_outsourcing_qty['日期2'] = erp_outsourcing_qty['日期']
    erp_outsourcing_qty['日期2'] = erp_outsourcing_qty['日期2'].apply(lambda x: str(x)[0:10])
    erp_outsourcing_qty['月份'] = erp_outsourcing_qty['日期'].apply(lambda x: str(x).split('-')[1])
    erp_outsourcing_qty['产品类别'] = erp_outsourcing_qty['产品类别'].astype(str)
    erp_outsourcing_qty = erp_outsourcing_qty[((erp_outsourcing_qty['产品类别']).__eq__('成品'))]
    erp_outsourcing_qty['采购供应商'] = erp_outsourcing_qty['采购供应商'].astype(str).replace(['nan', 'NaN'], '')
    erp_outsourcing_qty['销售客户'] = erp_outsourcing_qty['销售客户'].astype(str).replace(['nan', 'NaN'], '')
    erp_outsourcing_qty['产品参考'] = erp_outsourcing_qty['产品参考'].astype(str).replace(['nan', 'NaN'], '')
    erp_outsourcing_qty['产品条码'] = erp_outsourcing_qty['产品条码'].astype(str).replace(['nan', 'NaN'], '')
    erp_outsourcing_qty['采购供应商'] = erp_outsourcing_qty['采购供应商'].astype(str)
    skus = erp_outsourcing_qty[(
            (~((erp_outsourcing_qty['采购供应商'].str.contains('深圳市麦凯莱科技有限公司'))
               | ((erp_outsourcing_qty['采购供应商']).__eq__(''))))
            & (~((erp_outsourcing_qty['产品条码']).__eq__('')))
    )]['产品条码']
    if not skus.empty:
        skus = skus.drop_duplicates().tolist()
        erp_outsourcing_qty = erp_outsourcing_qty[erp_outsourcing_qty['产品条码'].isin(skus)]
    if not erp_outsourcing_qty.empty:
        erp_outsourcing_qty = erp_outsourcing_qty.groupby(
            ['公司', '月份', '日期2', '业务类型', '产品参考', '产品条码', '采购供应商', '销售客户']
        ).agg({'完成数量': np.sum}
              ).sort_values(['产品条码', '月份', '日期2', '业务类型']).reset_index()
        # erp_outsourcing_qty.sort_values(['产品条码', '日期2', '业务类型'], inplace = True)
        erp_outsourcing_qty['完成数量'] = erp_outsourcing_qty.apply(
            lambda x: 0 - x['完成数量'] if (str(x['业务类型']).__contains__('销售')) else x['完成数量'], axis=1)
        erp_outsourcing_qty['期初'] = 0
        erp_outsourcing_qty['采购'] = erp_outsourcing_qty.apply(
            lambda x: x['完成数量'] if (str(x['业务类型']).__contains__('采购')) else 0, axis=1)
        erp_outsourcing_qty['销售'] = erp_outsourcing_qty.apply(
            lambda x: x['完成数量'] if (str(x['业务类型']).__contains__('销售')) else 0, axis=1)
        sku_arr = erp_outsourcing_qty['产品条码'].drop_duplicates().tolist()
        for sku in sku_arr:
            pd_sku_tmp = erp_outsourcing_qty[erp_outsourcing_qty['产品条码'] == sku]
            pd_sku_tmp = pd_sku_tmp.sort_values(['月份', '日期2'])
            mix_index = min(pd_sku_tmp.index)
            last_stock = 0
            for index, row in pd_sku_tmp.iterrows():
                erp_outsourcing_qty.at[index, '期初'] = last_stock
                if index == mix_index:
                    erp_outsourcing_qty.at[index, '期初'] = 0
                erp_outsourcing_qty.at[index, '结存'] = (
                        erp_outsourcing_qty.at[index, '期初'] + erp_outsourcing_qty.at[index, '采购']
                        + erp_outsourcing_qty.at[index, '销售']
                )
                last_stock = erp_outsourcing_qty.at[index, '结存']
    erp_outsourcing_qty_total = erp_outsourcing_qty.copy()
    erp_outsourcing_qty_total["日期2"] = pd.to_datetime(erp_outsourcing_qty_total["日期2"])
    erp_outsourcing_qty_total['日期2'] = erp_outsourcing_qty_total['日期2'].apply(lambda x: str(x)[:7])
    erp_outsourcing_qty_total['业务类型'] = '月度(采购/销售)合计'
    erp_outsourcing_qty_total['采购供应商'] = '/'
    erp_outsourcing_qty_total['销售客户'] = '合计:'
    erp_outsourcing_qty_total = erp_outsourcing_qty_total.groupby(
        ['公司', '月份', '日期2', '业务类型', '产品参考', '产品条码', '采购供应商', '销售客户']
    ).agg({'完成数量': np.sum, '期初': np.sum, '采购': np.sum, '销售': np.sum, '结存': np.sum}
          ).sort_values(['产品条码', '月份', '日期2']).reset_index()
    # erp_outsourcing_qty_total['期初'] = 0
    # erp_outsourcing_qty_total['结存'] = erp_outsourcing_qty_total.apply(lambda x: x['期初'] + x['采购'] + x['销售'], axis = 1)
    sku_arr = erp_outsourcing_qty_total['产品条码'].drop_duplicates().tolist()
    for sku in sku_arr:
        pd_sku_tmp = erp_outsourcing_qty_total[erp_outsourcing_qty_total['产品条码'] == sku]
        pd_sku_tmp = pd_sku_tmp.sort_values(['月份', '日期2'])
        mix_index = min(pd_sku_tmp.index)
        last_stock = 0
        for index, row in pd_sku_tmp.iterrows():
            erp_outsourcing_qty_total.at[index, '期初'] = last_stock
            if index == mix_index:
                erp_outsourcing_qty_total.at[index, '期初'] = 0
            erp_outsourcing_qty_total.at[index, '结存'] = (
                    erp_outsourcing_qty_total.at[index, '期初'] + erp_outsourcing_qty_total.at[index, '采购']
                    + erp_outsourcing_qty_total.at[index, '销售']
            )
            last_stock = erp_outsourcing_qty_total.at[index, '结存']
    erp_outsourcing_qty = pd.concat([erp_outsourcing_qty, erp_outsourcing_qty_total])
    erp_outsourcing_qty.columns = ['主体', '月份', '日期', '业务类型', '产品参考', '产品条码', '采购供应商', '销售客户',
                                   '完成数量', '期初', '采购', '销售', '结存']
    erp_outsourcing_qty = erp_outsourcing_qty.sort_values(['产品条码', '月份', '日期', '业务类型']).reset_index()
    erp_outsourcing_qty = erp_outsourcing_qty[
        ['月份', '日期', '业务类型', '产品参考', '产品条码', '采购供应商', '销售客户', '完成数量', '期初', '采购',
         '销售', '结存']]

    # 处理-移动库存-采购入库与销售出库对比(时间[yyyy-MM-dd]-内部参考号(产品条码)-数量,建立唯一索引对比)
    erp_product_purchase_sale = pd_erp_cur.copy()
    erp_product_purchase_sale = erp_product_purchase_sale[
        ['库存移动ID', '日期', '类型', '公司', '业务类型', '产品参考', '产品条码', '产品名称', '产品类别', '采购供应商',
         '完成数量', '完成金额(RMB)', '销售已交货含税金额(RMB)', '采购/退货(含税)金额(原币别)']]
    erp_product_purchase_sale['日期'] = erp_product_purchase_sale['日期'].apply(lambda x: str(x)[0:10])
    erp_product_purchase_sale['产品参考'] = erp_product_purchase_sale['产品参考'].astype(str).replace(
        ['nan', 'NaN', '.0'], '')
    erp_product_purchase_sale['产品条码'] = erp_product_purchase_sale['产品条码'].astype(str).replace(
        ['nan', 'NaN', '.0'], '')
    erp_product_purchase_sale['产品类别'] = erp_product_purchase_sale['产品类别'].astype(str)
    erp_product_purchase_sale = erp_product_purchase_sale[((erp_product_purchase_sale['产品类别']).__eq__('成品'))]
    erp_product_purchase_sale = erp_product_purchase_sale[(~((erp_product_purchase_sale['产品条码']).__eq__('')))]
    erp_product_purchase_sale['业务类型'] = erp_product_purchase_sale['业务类型'].astype(str)
    erp_product_purchase = erp_product_purchase_sale[((erp_product_purchase_sale['业务类型']).str.contains('采购'))]
    erp_product_purchase['Index'] = ''
    if not erp_product_purchase.empty:
        erp_product_purchase['Index'] = erp_product_purchase.apply(
            lambda x: str.format("{}#{}#{}",
                                 x['日期'],
                                 (str(int(x['产品参考'])) if str.isnumeric(x['产品参考']) else str(x['产品参考'])),
                                 int(x['完成数量'])), axis=1)
    erp_product_sale = erp_product_purchase_sale[
        ((erp_product_purchase_sale['业务类型']).str.contains('销售'))]
    erp_product_sale['Index'] = ''
    if not erp_product_sale.empty:
        erp_product_sale['Index'] = erp_product_sale.apply(
            lambda x: str.format("{}#{}#{}", x['日期'],
                                 (str(int(x['产品参考'])) if str.isnumeric(x['产品参考']) else str(x['产品参考'])),
                                 int(x['完成数量'])),
            axis=1)
    erp_product_purchase_index = erp_product_purchase['Index'].drop_duplicates().tolist()
    if erp_product_purchase_index.__len__() > 0 and erp_product_purchase_index.__contains__(''):
        erp_product_purchase_index = erp_product_purchase_index.remove('')
    erp_product_sale_index = erp_product_sale['Index'].drop_duplicates().tolist()
    if erp_product_sale_index.__len__() > 0 and erp_product_sale_index.__contains__(''):
        erp_product_sale_index = erp_product_sale_index.remove('')
    erp_product_purchase_not_sold = erp_product_purchase[
        (~((erp_product_purchase['Index']).isin(erp_product_sale_index)))]
    erp_product_sale_not_purchased = erp_product_sale[
        (~((erp_product_sale['Index']).isin(erp_product_purchase_index)))]

    df_erp_diff = pd.concat([erp_product_purchase_not_sold, erp_product_sale_not_purchased])
    df_erp_diff['对比类型'] = 'ERP采购与销售'
    df_erp_diff['建议'] = ''
    if not df_erp_diff.empty:
        df_erp_diff['业务类型'] = df_erp_diff['业务类型'].astype(str)
        df_erp_diff['建议'] = df_erp_diff.apply(
            lambda x: '采少, 销多' if (str(x['业务类型']).__contains__('销售')) else '采多, 销少',
            axis=1)
    df_erp_diff = df_erp_diff.sort_values(['日期', '业务类型'], ascending=[True, True])

    # 处理-ERP移动库存-采购入库与销售出库对比(时间[yyyy-MM-dd]-内部参考号(产品条码)-数量-含税金额,建立唯一索引对比)
    mega_product_sold = pd_relation_cur.copy()
    mega_product_sold['类型'] = mega_product_sold['类型'].astype(str)
    mega_product_sold['业务类型'] = mega_product_sold['业务类型'].astype(str)
    mega_product_sold['销售客户'] = mega_product_sold['销售客户'].astype(str)
    mega_product_sold = mega_product_sold[(
            ((mega_product_sold['类型']).str.contains('公司关联交易'))
            & ((mega_product_sold['业务类型']).str.contains('销售'))
            & ((mega_product_sold['销售客户']).isin(subject_current))
    )]
    mega_product_sold = mega_product_sold[
        ['库存移动ID', '日期', '类型', '公司', '业务类型', '产品参考', '产品条码', '产品名称', '产品类别', '完成数量',
         '完成金额(RMB)', '销售已交货含税金额(RMB)', '采购/退货(含税)金额(原币别)']]
    mega_product_sold['日期'] = mega_product_sold['日期'].apply(lambda x: str(x)[0:10])
    mega_product_sold['产品参考'] = mega_product_sold['产品参考'].astype(str).replace(['nan', 'NaN', '.0'], '')
    mega_product_sold['产品条码'] = mega_product_sold['产品条码'].astype(str).replace(['nan', 'NaN', '.0'], '')
    mega_product_sold['产品类别'] = mega_product_sold['产品类别'].astype(str)
    mega_product_sold = mega_product_sold[((mega_product_sold['产品类别']).__eq__('成品'))]
    mega_product_sold = mega_product_sold[(~((mega_product_sold['产品条码']).__eq__('')))]
    mega_product_sold['Index'] = ''
    if not mega_product_sold.empty:
        mega_product_sold['Index'] = mega_product_sold.apply(
            lambda x: str.format("{}#{}#{}",
                                 x['日期'],
                                 (str(int(x['产品参考'])) if str.isnumeric(x['产品参考']) else str(x['产品参考'])),
                                 int(x['完成数量'])),
            axis=1)
    mega_product_sold_index = mega_product_sold['Index'].drop_duplicates().tolist()
    if mega_product_sold_index.__len__() > 0 and mega_product_sold_index.__contains__(''):
        mega_product_sold_index = mega_product_sold_index.remove('')
    erp_product_purchased = pd_erp_cur.copy()
    erp_product_purchased['业务类型'] = erp_product_purchased['业务类型'].astype(str)
    erp_product_purchased['采购供应商'] = erp_product_purchased['采购供应商'].astype(str)
    erp_product_purchased = erp_product_purchased[(
            ((erp_product_purchased['业务类型']).str.contains('采购'))
            & ((erp_product_purchased['采购供应商']).str.contains('深圳市麦凯莱科技有限公司'))
    )]
    erp_product_purchased = erp_product_purchased[
        ['库存移动ID', '日期', '类型', '公司', '业务类型', '产品参考', '产品条码', '产品名称', '产品类别', '完成数量',
         '完成金额(RMB)', '销售已交货含税金额(RMB)', '采购/退货(含税)金额(原币别)']]
    erp_product_purchased['日期'] = erp_product_purchased['日期'].apply(lambda x: str(x)[0:10])
    erp_product_purchased['产品参考'] = erp_product_purchased['产品参考'].astype(str).replace(['nan', 'NaN', '.0'], '')
    erp_product_purchased['产品条码'] = erp_product_purchased['产品条码'].astype(str).replace(['nan', 'NaN', '.0'], '')
    erp_product_purchased['产品类别'] = erp_product_purchased['产品类别'].astype(str)
    erp_product_purchased = erp_product_purchased[((erp_product_purchased['产品类别']).__eq__('成品'))]
    erp_product_purchased = erp_product_purchased[(~((erp_product_purchased['产品条码']).__eq__('')))]
    erp_product_purchased['Index'] = ''
    if not erp_product_purchased.empty:
        erp_product_purchased['Index'] = erp_product_purchased.apply(
            lambda x: str.format("{}#{}#{}",
                                 x['日期'],
                                 (str(int(x['产品参考'])) if str.isnumeric(x['产品参考']) else str(x['产品参考'])),
                                 int(x['完成数量'])),
            axis=1)
    erp_product_purchased_index = erp_product_purchased['Index'].drop_duplicates().tolist()
    if erp_product_purchased_index.__len__() > 0 and erp_product_purchased_index.__contains__(''):
        erp_product_purchased_index = erp_product_purchased_index.remove('')
    mega_product_sold_erp_not_purchased = mega_product_sold[
        (~((mega_product_sold['Index']).isin(erp_product_purchased_index)))]
    erp_product_purchased_mega_not_sold = erp_product_purchased[
        (~((erp_product_purchased['Index']).isin(mega_product_sold_index)))]

    df_mega_erp_diff = pd.concat([mega_product_sold_erp_not_purchased, erp_product_purchased_mega_not_sold])
    df_mega_erp_diff['对比类型'] = 'MEGA销售与ERP采购'
    df_mega_erp_diff['建议'] = ''
    df_mega_erp_diff = df_mega_erp_diff.sort_values(['日期', '业务类型'], ascending=[True, False])
    df_diff = pd.concat([df_erp_diff, df_mega_erp_diff])
    df_diff["完成数量"].fillna(0, inplace=True)
    df_diff["完成金额(RMB)"].fillna(0, inplace=True)
    df_diff["销售已交货含税金额(RMB)"].fillna(0, inplace=True)
    df_diff["采购/退货(含税)金额(原币别)"].fillna(0, inplace=True)
    if not df_diff.empty:
        df_diff['产品参考'] = df_diff['产品参考'].apply(lambda x: str(x).replace('.0', ''))
        df_diff['产品条码'] = df_diff['产品条码'].apply(lambda x: str(x).replace('.0', ''))
        df_diff['销售已交货含税金额(RMB)'] = round(df_diff['销售已交货含税金额(RMB)'], 2)
        df_diff['采购/退货(含税)金额(原币别)'] = round(df_diff['采购/退货(含税)金额(原币别)'], 2)
        df_diff = df_diff.sort_values(['日期'])

    # 处理-ERP移动库存-根据销售出库透视测算加权
    print('\n处理-ERP移动库存-根据销售出库透视测算加权...')
    erp_sale_weighting = pd_erp_cur.copy()
    erp_sale_weighting = erp_sale_weighting[
        ['公司', '日期', '业务类型', '产品参考', '产品条码', '产品类别', '月加权平均价(RMB)', '完成数量',
         '完成金额(RMB)', '销售不含税金额(原币别)']]
    erp_sale_weighting['日期2'] = erp_sale_weighting['日期']
    erp_sale_weighting['日期2'] = erp_sale_weighting['日期2'].apply(lambda x: str(x)[0:10])
    erp_sale_weighting['月份'] = erp_sale_weighting['日期'].apply(lambda x: str(x).split('-')[1])
    erp_sale_weighting['产品类别'] = erp_sale_weighting['产品类别'].astype(str)
    erp_sale_weighting = erp_sale_weighting[((erp_sale_weighting['产品类别']).__eq__('成品'))]
    erp_sale_weighting['业务类型'] = erp_sale_weighting['业务类型'].astype(str)
    erp_sale_weighting = erp_sale_weighting[((erp_sale_weighting['业务类型']).str.contains('销售'))]
    erp_sale_weighting = erp_sale_weighting.groupby(['月份', '产品参考', '产品条码', '月加权平均价(RMB)']).agg(
        {'完成数量': np.sum, '完成金额(RMB)': np.sum, '销售不含税金额(原币别)': np.sum}).reset_index()
    erp_sale_weighting['财务完成金额'] = erp_sale_weighting['销售不含税金额(原币别)'] * 0.7
    erp_sale_weighting['财务加权单价'] = erp_sale_weighting.apply(
        lambda x: 0.00 if (x['完成数量'] == 0) else (x['财务完成金额'] / x['完成数量']), axis=1)
    erp_sale_weighting['加权单价差异'] = (erp_sale_weighting['月加权平均价(RMB)'] - erp_sale_weighting['财务加权单价'])
    erp_sale_weighting['加权单价差异'].fillna(0, inplace=True)
    erp_sale_weighting['加权单价差异'] = round(erp_sale_weighting['加权单价差异'], 8)
    erp_sale_weighting = erp_sale_weighting[
        ['月份', '产品参考', '产品条码', '月加权平均价(RMB)', '完成数量', '完成金额(RMB)', '销售不含税金额(原币别)',
         '财务完成金额', '财务加权单价', '加权单价差异']]

    # 整合-采购进销明细对比
    print('\n整合-采购进销明细对比...')
    # cg_compare = pd.merge(cg, erp2, how="outer", on=["公司主体", "供应商名称", "送货日期"])
    cg_compare = pd.merge(cg, erp, how="outer", on=["公司主体", "供应商名称", "送货日期"])
    cg_compare["采购含税金额-供应商"].fillna(0, inplace=True)
    cg_compare["ERP采购含税金额-供应商"].fillna(0, inplace=True)
    cg_compare["采购含税金额差异"] = (cg_compare["采购含税金额-供应商"] - cg_compare["ERP采购含税金额-供应商"])
    cg_compare["采购含税金额差异"].fillna(0, inplace=True)
    cg_compare["采购含税金额差异"] = round(cg_compare["采购含税金额差异"], 8)
    cg_compare = cg_compare.sort_values('采购含税金额差异', ascending=False)
    cg_compare['公司主体'] = cg_compare['公司主体'].astype(str)
    cg_compare["公司主体"] = cg_compare["公司主体"].apply(lambda x: x.replace("(", "（").strip())
    cg_compare["公司主体"] = cg_compare["公司主体"].apply(lambda x: x.replace(")", "）").strip())
    cg_compare = cg_compare[
        ['公司主体', '供应商名称', '送货日期', '采购含税金额-供应商', 'ERP采购含税金额-供应商', '采购含税金额差异']]

    # 整合-内部销售对比-交易总额
    print('\n整合-内部销售对比-交易总额...')
    cg1_compare = pd.merge(cg1, erp1, how="outer", on=["公司主体", "送货日期"])
    cg1_compare["交易总额-对应子公司销售给“麦凯莱“"].fillna(0, inplace=True)
    cg1_compare["ERP交易总额-对应子公司销售给“麦凯莱“"].fillna(0, inplace=True)
    cg1_compare["交易总额差异"] = (
            cg1_compare["交易总额-对应子公司销售给“麦凯莱“"] - cg1_compare["ERP交易总额-对应子公司销售给“麦凯莱“"])
    cg1_compare["交易总额差异"].fillna(0, inplace=True)
    cg1_compare["交易总额差异"] = round(cg1_compare["交易总额差异"], 8)
    cg1_compare = cg1_compare.sort_values('交易总额差异', ascending=False)
    cg1_compare['公司主体'] = cg1_compare['公司主体'].astype(str)
    cg1_compare["公司主体"] = cg1_compare["公司主体"].apply(lambda x: x.replace("(", "（").strip())
    cg1_compare["公司主体"] = cg1_compare["公司主体"].apply(lambda x: x.replace(")", "）").strip())
    cg1_compare = cg1_compare[
        ["公司主体", "送货日期", "交易总额-对应子公司销售给“麦凯莱“", "ERP交易总额-对应子公司销售给“麦凯莱“",
         "交易总额差异"]]

    # 整合-2B销售明细对比
    print('\n整合-2B销售明细对比...')
    two_b_compare = pd.merge(two_b, erp_two_b, how="outer", on=["出货日期", "购买方名称-销售客户"])
    two_b_compare["销售出库-价税合计"].fillna(0, inplace=True)
    two_b_compare["ERP销售出库-价税合计"].fillna(0, inplace=True)
    two_b_compare["2B-销售出库-价税合计"].fillna(0, inplace=True)
    two_b_compare["ERP2B-销售出库-价税合计"].fillna(0, inplace=True)
    two_b_compare["2B发出商品-销售出库-价税合计"].fillna(0, inplace=True)
    two_b_compare["ERP2B发出商品-销售出库-价税合计"].fillna(0, inplace=True)
    two_b_compare["销售出库-价税差异"] = (two_b_compare["销售出库-价税合计"] - two_b_compare["ERP销售出库-价税合计"])
    two_b_compare["销售出库-价税差异"].fillna(0, inplace=True)
    two_b_compare["销售出库-价税差异"] = round(two_b_compare["销售出库-价税差异"], 8)
    two_b_compare["2B-销售出库-价税差异"] = (
            two_b_compare["2B-销售出库-价税合计"] - two_b_compare["ERP2B-销售出库-价税合计"])
    two_b_compare["2B-销售出库-价税差异"].fillna(0, inplace=True)
    two_b_compare["2B-销售出库-价税差异"] = round(two_b_compare["2B-销售出库-价税差异"], 8)
    two_b_compare["2B发出商品-销售出库-价税差异"] = (
            two_b_compare["2B发出商品-销售出库-价税合计"] - two_b_compare["ERP2B发出商品-销售出库-价税合计"])
    two_b_compare["2B发出商品-销售出库-价税差异"].fillna(0, inplace=True)
    two_b_compare["2B发出商品-销售出库-价税差异"] = round(two_b_compare["2B发出商品-销售出库-价税差异"], 8)
    two_b_compare = two_b_compare.sort_values('销售出库-价税差异', ascending=False)
    two_b_compare = two_b_compare[~two_b_compare['购买方名称-销售客户'].str.contains("nan")]
    two_b_compare = two_b_compare[
        ["出货日期", "购买方名称-销售客户", "销售出库-价税合计", "ERP销售出库-价税合计", "销售出库-价税差异",
         "2B-销售出库-价税合计", "ERP2B-销售出库-价税合计", "2B-销售出库-价税差异",
         "2B发出商品-销售出库-价税合计", "ERP2B发出商品-销售出库-价税合计", "2B发出商品-销售出库-价税差异"]]

    # 整合-店铺销售对比
    print(str.format('\n整合-店铺销售对比(税率值: {})...', tax_current))
    dp_compare = pd.merge(dp, erp_shop, how="outer", on=["主体", "平台", "财务店铺名称"])
    dp_compare["回款合计"].fillna(0, inplace=True)
    dp_compare["2C商品回款"].fillna(0, inplace=True)
    dp_compare["2C发出商品回款"].fillna(0, inplace=True)
    dp_compare["ERP回款合计"].fillna(0, inplace=True)
    dp_compare["ERP2C商品回款"].fillna(0, inplace=True)
    dp_compare["ERP2C发出商品回款"].fillna(0, inplace=True)
    dp_compare["ERP完成金额(公式计算)"].fillna(0, inplace=True)
    dp_compare["ERP完成金额"].fillna(0, inplace=True)
    dp_compare["ERP完成金额差异"].fillna(0, inplace=True)
    dp_compare["回款合计差异"] = (dp_compare["回款合计"] - dp_compare["ERP回款合计"])
    dp_compare["回款合计差异"].fillna(0, inplace=True)
    dp_compare["回款合计差异"] = round(dp_compare["回款合计差异"], 8)
    dp_compare["2C商品回款差异"] = (dp_compare["2C商品回款"] - dp_compare["ERP2C商品回款"])
    dp_compare["2C商品回款差异"].fillna(0, inplace=True)
    dp_compare["2C商品回款差异"] = round(dp_compare["2C商品回款差异"], 8)
    dp_compare["2C发出商品回款差异"] = (dp_compare["2C发出商品回款"] - dp_compare["ERP2C发出商品回款"])
    dp_compare["2C发出商品回款差异"].fillna(0, inplace=True)
    dp_compare["2C发出商品回款差异"] = round(dp_compare["2C发出商品回款差异"], 8)
    dp_compare["完成金额(公式计算)"] = 0
    dp_compare["税率说明"] = ''
    if tax_current == 0:
        print('\n税率不为3%或13%!!!')
        dp_compare['完成金额(公式计算)'] = dp_compare.apply(lambda x: (float(x['回款合计']) / 1 * 0.7), axis=1)
        dp_compare["税率说明"] = '税率不为3%或13%, 请人工核验'
    elif tax_current == 1:
        dp_compare['完成金额(公式计算)'] = dp_compare.apply(lambda x: (float(x['回款合计']) / 1 * 0.7), axis=1)
        dp_compare["税率说明"] = '税率为1%或空, 请人工核验'
    elif tax_current == 1.03:
        dp_compare['完成金额(公式计算)'] = dp_compare.apply(lambda x: (float(x['回款合计']) / 1.03 * 0.7), axis=1)
        dp_compare["税率说明"] = '税率为3%'
    elif tax_current == 1.13:
        dp_compare['完成金额(公式计算)'] = dp_compare.apply(lambda x: (float(x['回款合计']) / 1.13 * 0.7), axis=1)
        dp_compare["税率说明"] = '税率为13%'
    elif tax_current == 2:
        print('\n税率含3%, 13%等多种!!!')
        dp_compare['完成金额(公式计算)'] = dp_compare.apply(lambda x: (float(x['回款合计']) / 1 * 0.7), axis=1)
        dp_compare["税率说明"] = '税率含3%, 13%等多种, 请人工核验'
    dp_compare['完成金额差异'] = (dp_compare["完成金额(公式计算)"] - dp_compare["ERP完成金额"])
    dp_compare['完成金额差异'].fillna(0, inplace=True)
    dp_compare['完成金额差异'] = round(dp_compare['完成金额差异'], 8)
    dp_compare = dp_compare.sort_values('回款合计差异', ascending=False)

    # 整合-店铺销售对比-成本合计
    print('\n整合-店铺销售对比-成本合计...')
    dp_compare_current = dp_compare.copy()
    dp_compare_current['平台'] = '/'
    dp_compare_current['财务店铺名称'] = '合计:'
    dp_compare_total = dp_compare_current.groupby(['主体', '平台', '财务店铺名称']).agg(
        {'回款合计': np.sum, 'ERP回款合计': np.sum, '回款合计差异': np.sum, '2C商品回款': np.sum,
         'ERP2C商品回款': np.sum, '2C商品回款差异': np.sum,
         '2C发出商品回款': np.sum, 'ERP2C发出商品回款': np.sum, '2C发出商品回款差异': np.sum,
         '完成金额(公式计算)': np.sum, 'ERP完成金额(公式计算)': np.sum,
         'ERP完成金额': np.sum, '完成金额差异': np.sum, 'ERP完成金额差异': np.sum}
    ).reset_index()
    dp_compare_total = pd.merge(dp_compare_total, erp_shop_cost, how="outer", on=["主体"])
    dp_compare["ERP采购含税金额合计"] = 0
    dp_compare["ERP采购含税(3%)金额合计"] = 0
    dp_compare["ERP采购含税(13%)金额合计"] = 0
    dp_compare = pd.concat([dp_compare, dp_compare_total])
    dp_compare['主体'] = dp_compare['主体'].astype(str)
    dp_compare['主体'] = dp_compare['主体'].apply(lambda x: x.replace("(", "（").strip())
    dp_compare['主体'] = dp_compare['主体'].apply(lambda x: x.replace(")", "）").strip())
    dp_compare = dp_compare[
        ["主体", "平台", "财务店铺名称", "回款合计", "ERP回款合计", "回款合计差异", "2C商品回款", "ERP2C商品回款",
         "2C商品回款差异", "2C发出商品回款", "ERP2C发出商品回款",
         "2C发出商品回款差异", "完成金额(公式计算)", "ERP完成金额", "ERP完成金额(公式计算)", "完成金额差异",
         "ERP完成金额差异", "ERP采购含税金额合计", "ERP采购含税(3%)金额合计", "ERP采购含税(13%)金额合计", "税率说明"]]

    # 筛选本次测算公司主体的数据
    cg_compare = cg_compare[(cg_compare['公司主体'].isin(subject_current))]
    cg1_compare = cg1_compare[(cg1_compare['公司主体'].isin(subject_current))]
    dp_compare = dp_compare[(dp_compare['主体'].isin(subject_current))]

    # 程序校验文件
    print(str.format('\n开始程序校验文件: {}\n', save_file_cur))
    check_ok = True
    failed_msg = []
    print('\n校验: 采购进销明细对比...')
    if not cg_compare.empty:
        table_header = ['公司主体', '供应商名称', '送货日期', '采购含税金额-供应商', 'ERP采购含税金额-供应商',
                        '采购含税金额差异']
        cg_compare_check = cg_compare[(abs(cg_compare["采购含税金额差异"]) > 1)]
        if cg_compare_check.empty:
            print("\n文件: 采购进销明细对比, 验收成功~~~OK~~~")
        else:
            check_ok = False
            failed_msg.append('采购进销明细对比: 采购含税金额差异, 验收异常')
            print("\n文件: 采购进销明细对比: 采购含税金额差异, 验收异常!!!FAILED!!!")

            print("\n查看: 差异(按任意键继续)?")
            # input()
            print(tabulate(cg_compare_check, headers=table_header, tablefmt='orgtbl'))
    else:
        print("\n文件: 采购进销明细对比 -- 无数据, 验收异常!!!FAILED!!!")

    print('\n校验: 内部销售对比...')
    if not cg1_compare.empty:
        table_header = ['公司主体', '送货日期', '交易总额-对应子公司销售给“麦凯莱“',
                        'ERP交易总额-对应子公司销售给“麦凯莱“', '交易总额差异']
        cg1_compare_check = cg1_compare[(abs(cg1_compare["交易总额差异"]) > 1)]
        if cg1_compare_check.empty:
            print("\n文件: 内部销售对比, 验收成功~~~OK~~~")
        else:
            check_ok = False
            failed_msg.append('内部销售对比: 交易总额差异, 验收异常')
            print("\n文件: 内部销售对比: 交易总额差异, 验收异常!!!FAILED!!!")

            print("\n查看: 差异(按任意键继续)?")
            # input()
            print(tabulate(cg1_compare_check, headers=table_header, tablefmt='orgtbl'))
    else:
        print("\n文件: 内部销售对比 -- 无数据, 验收异常!!!FAILED!!!")

    print('\n校验: 2B销售明细对比...')
    if not two_b_compare.empty:
        two_b_compare_check_01 = two_b_compare[(abs(two_b_compare["销售出库-价税差异"]) > 1)]
        if two_b_compare_check_01.empty:
            print("\n文件: 2B销售明细对比: 销售出库-价税差异, 验收成功~~~OK~~~")
        else:
            check_ok = False
            failed_msg.append('2B销售明细对比: 销售出库-价税差异, 验收异常')
            print("\n文件: 2B销售明细对比: 销售出库-价税差异, 验收异常!!!FAILED!!!")

            print("\n查看: 差异(按任意键继续)?")
            # input()
            two_b_compare_check_01 = two_b_compare_check_01[
                ['出货日期', '购买方名称-销售客户', '销售出库-价税合计', 'ERP销售出库-价税合计', '销售出库-价税差异']]
            table_header = ['出货日期', '购买方名称-销售客户', '销售出库-价税合计', 'ERP销售出库-价税合计',
                            '销售出库-价税差异']
            print(tabulate(two_b_compare_check_01, headers=table_header, tablefmt='orgtbl'))

        two_b_compare_check_02 = two_b_compare[(abs(two_b_compare["2B-销售出库-价税差异"]) > 1)]
        if two_b_compare_check_02.empty:
            print("\n文件: 2B销售明细对比: 2B-销售出库-价税差异, 验收成功~~~OK~~~")
        else:
            check_ok = False
            failed_msg.append('2B销售明细对比: 2B-销售出库-价税差异, 验收异常')
            print("\n文件: 2B销售明细对比: 2B-销售出库-价税差异, 验收异常!!!FAILED!!!")

            print("\n查看: 差异(按任意键继续)?")
            # input()
            two_b_compare_check_02 = two_b_compare_check_02[
                ['出货日期', '购买方名称-销售客户', '2B-销售出库-价税合计', 'ERP2B-销售出库-价税合计',
                 '2B-销售出库-价税差异']]
            table_header = ['出货日期', '购买方名称-销售客户', '2B-销售出库-价税合计', 'ERP2B-销售出库-价税合计',
                            '2B-销售出库-价税差异']
            print(tabulate(two_b_compare_check_02, headers=table_header, tablefmt='orgtbl'))

            two_b_compare_check_03 = two_b_compare[(abs(two_b_compare["2B发出商品-销售出库-价税差异"]) > 1)]
            if two_b_compare_check_03.empty:
                print("\n文件: 2B销售明细对比: 2B发出商品-销售出库-价税差异, 验收成功~~~OK~~~")
            else:
                check_ok = False
                failed_msg.append('2B销售明细对比: 2B发出商品-销售出库-价税差异, 验收异常')
                print("\n文件: 2B销售明细对比: 2B发出商品-销售出库-价税差异, 验收异常!!!FAILED!!!")

                print("\n查看: 差异(按任意键继续)?")
                # input()
                two_b_compare_check_03 = two_b_compare_check_03[
                    ['出货日期', '购买方名称-销售客户', '2B发出商品-销售出库-价税合计',
                     'ERP2B发出商品-销售出库-价税合计', '2B发出商品-销售出库-价税差异']]
                table_header = ['出货日期', '购买方名称-销售客户', '2B发出商品-销售出库-价税合计',
                                'ERP2B发出商品-销售出库-价税合计', '2B发出商品-销售出库-价税差异']
                print(tabulate(two_b_compare_check_03, headers=table_header, tablefmt='orgtbl'))
    else:
        print("\n文件: 2B销售明细对比 -- 无数据, 验收异常!!!FAILED!!!")

    print('\n校验: 店铺销售对比...')
    if not dp_compare.empty:
        dp_compare['财务店铺名称'] = dp_compare['财务店铺名称'].astype(str)
        print('\n校验: 店铺销售对比-回款合计...')
        dp_compare_check_01_00 = dp_compare[
            (~(dp_compare['财务店铺名称'].str.contains('合计')) & (abs(dp_compare['回款合计差异']) > 1))]
        if dp_compare_check_01_00.empty:
            print("\n文件: 店铺销售对比-回款合计, 验收成功~~~OK~~~")
        else:
            check_ok = False
            failed_msg.append('店铺销售对比: 回款合计差异, 验收异常')
            print("\n文件: 店铺销售对比-回款合计, 验收异常!!!FAILED!!!")

            print("\n查看: 差异(按任意键继续)?")
            # input()
            dp_compare_check_01_00 = dp_compare_check_01_00[
                ['主体', '平台', '财务店铺名称', '回款合计', 'ERP回款合计', '回款合计差异']]
            table_header = ['主体', '平台', '财务店铺名称', '回款合计', 'ERP回款合计', '回款合计差异']
            print(tabulate(dp_compare_check_01_00, headers=table_header, tablefmt='orgtbl'))

        print('\n校验: 店铺销售对比-2C商品回款...')
        dp_compare_check_01_01 = dp_compare[(
                ~(dp_compare['财务店铺名称'].str.contains('合计'))
                & (abs(dp_compare['2C商品回款差异']) > 1)
        )]
        if dp_compare_check_01_01.empty:
            print("\n文件: 店铺销售对比-2C商品回款, 验收成功~~~OK~~~")
        else:
            check_ok = False
            failed_msg.append('店铺销售对比: 2C商品回款差异, 验收异常')
            print("\n文件: 店铺销售对比-2C商品回款, 验收异常!!!FAILED!!!")

            print("\n查看: 差异(按任意键继续)?")
            # input()
            dp_compare_check_01_01 = dp_compare_check_01_01[
                ['主体', '平台', '财务店铺名称', '2C商品回款', 'ERP2C商品回款', '2C商品回款差异']]
            table_header = ['主体', '平台', '财务店铺名称', '2C商品回款', 'ERP2C商品回款', '2C商品回款差异']
            print(tabulate(dp_compare_check_01_01, headers=table_header, tablefmt='orgtbl'))

        print('\n校验: 店铺销售对比-2C发出商品回款...')
        dp_compare_check_01_02 = dp_compare[(
                ~(dp_compare['财务店铺名称'].str.contains('合计'))
                & (abs(dp_compare['2C发出商品回款差异']) > 1)
        )]
        if dp_compare_check_01_02.empty:
            print("\n文件: 店铺销售对比-2C发出商品回款, 验收成功~~~OK~~~")
        else:
            check_ok = False
            failed_msg.append('店铺销售对比: 2C发出商品回款差异, 验收异常')
            print("\n文件: 店铺销售对比-2C发出商品回款, 验收异常!!!FAILED!!!")

            print("\n查看: 差异(按任意键继续)?")
            # input()
            dp_compare_check_01_02 = dp_compare_check_01_02[
                ['主体', '平台', '财务店铺名称', '2C发出商品回款', 'ERP2C发出商品回款', '2C发出商品回款差异']]
            table_header = ['主体', '平台', '财务店铺名称', '2C发出商品回款', 'ERP2C发出商品回款', '2C发出商品回款差异']
            print(tabulate(dp_compare_check_01_02, headers=table_header, tablefmt='orgtbl'))

        print('\n校验: 店铺销售对比-完成金额...')
        dp_compare_check_02 = dp_compare[
            ((dp_compare['财务店铺名称'].str.contains('合计')) & (abs(dp_compare['完成金额差异']) > 1))]
        if dp_compare_check_02.empty:
            print("\n文件: 店铺销售对比-完成金额, 验收成功~~~OK~~~")
        else:
            completion_amount_ok = False
            print("\n查看: 销售出库透视测算加权数据, 进一步判断完成金额差异问题")
            if not erp_sale_weighting.empty:
                table_header = ['月份', '产品参考', '产品条码', '月加权平均价(RMB)', '完成数量', '完成金额(RMB)',
                                '销售不含税金额(原币别)', '财务完成金额', '财务加权单价', '加权单价差异']
                erp_sale_weighting['加权单价差异'] = erp_sale_weighting['加权单价差异'].astype(float)
                erp_sale_weighting_check = erp_sale_weighting[(abs(erp_sale_weighting['加权单价差异']) > 0.00001)]
                if erp_sale_weighting_check.empty:
                    print("\n查看: 销售出库透视测算加权--已清零--~~~OK~~~")
                else:
                    completion_amount_ok = True
                    check_ok = False
                    failed_msg.append('销售出库透视测算加权: --未清零(绝对值>0.00001)--, 验收异常')
                    print("\n查看: 销售出库透视测算加权--未清零--!!!FAILED!!!")

                    print("\n查看: 差异(按任意键继续)?")
                    # input()
                    print(tabulate(erp_sale_weighting_check, headers=table_header, tablefmt='orgtbl'))

            else:
                print("\n文件: 销售出库透视测算加权数据 -- 无数据, 验收异常!!!FAILED!!!")

            if completion_amount_ok:
                check_ok = False
                failed_msg.append('店铺销售对比: 完成金额差异, 验收异常')
                print("\n文件: 店铺销售对比-完成金额, 验收异常!!!FAILED!!!")

                print("\n查看: 差异(按任意键继续)?")
                # input()
                dp_compare_check_02 = dp_compare_check_02[
                    ['主体', '财务店铺名称', '完成金额(公式计算)', 'ERP完成金额', '完成金额差异', 'ERP采购含税金额合计',
                     'ERP采购含税(3%)金额合计', 'ERP采购含税(13%)金额合计']]
                table_header = ['主体', '财务店铺名称', '完成金额(公式计算)', 'ERP完成金额', '完成金额差异',
                                'ERP采购含税金额合计', 'ERP采购含税(3%)金额合计', 'ERP采购含税(13%)金额合计']
                print(tabulate(dp_compare_check_02, headers=table_header, tablefmt='orgtbl'))
            else:
                print("\n文件: 店铺销售对比-完成金额, 验收成功~~~OK~~~")
    else:
        print("\n文件: 店铺销售对比 -- 无数据, 验收异常!!!FAILED!!!")

    print("\n查看: 产品数量结存")
    if not erp_outsourcing_qty.empty:
        table_header = ['月份', '日期', '业务类型', '产品参考', '产品条码', '采购供应商', '销售客户',
                        '完成数量', '期初', '采购', '销售', '结存']
        erp_outsourcing_qty_check = erp_outsourcing_qty.copy()
        erp_outsourcing_qty_check['销售客户'] = erp_outsourcing_qty_check['销售客户'].astype(str)

        print("\n查看: 产品数量结存-结存")
        erp_outsourcing_qty_check_sum = erp_outsourcing_qty_check[
            ((erp_outsourcing_qty_check['销售客户'].str.contains('合计')) & (
                    (erp_outsourcing_qty_check['结存']) != 0))]
        if erp_outsourcing_qty_check_sum.empty:
            print("\n文件: 产品数量结存-结存, 验收成功~~~OK~~~")
        else:
            check_ok = False
            failed_msg.append('产品数量结存: 结存--未清零--, 验收异常')
            print("\n文件: 产品数量结存-结存--未清零--, 验收异常!!!FAILED!!!")

            print("\n查看: 差异(按任意键继续)?")
            # input()
            print(tabulate(erp_outsourcing_qty_check_sum, headers=table_header, tablefmt='orgtbl'))

        print("\n查看: 产品数量结存-年度结存")
        erp_outsourcing_qty_check_year = erp_outsourcing_qty_check[
            (erp_outsourcing_qty_check['销售客户'].str.contains('合计'))].groupby(['产品条码']).agg(
            {'结存': np.sum}).reset_index()
        erp_outsourcing_qty_check_year_err = erp_outsourcing_qty_check_year[
            ((erp_outsourcing_qty_check_year['结存']) != 0)]
        if erp_outsourcing_qty_check_year_err.empty:
            print("\n查看: 产品数量结存-年度结存--已清零--~~~OK~~~")
        else:
            check_ok = False
            failed_msg.append('产品数量结存: 年度结存--未清零--, 验收异常')
            print("\n查看: 产品数量结存-年度结存--未清零--!!!FAILED!!!")

            print("\n查看: 差异(按任意键继续)?")
            # input()
            error_skus = erp_outsourcing_qty_check_year_err['产品条码'].drop_duplicates().tolist()
            print(tabulate(erp_outsourcing_qty_check[(
                    erp_outsourcing_qty_check['产品条码'].isin(error_skus)
                    & (~erp_outsourcing_qty_check['销售客户'].str.contains('合计'))
            )],
                           headers=table_header, tablefmt='orgtbl'))

        print("\n查看: 产品数量结存-月度结存 ~~ 2022年, 需月度清零")
        erp_outsourcing_qty_check_month_err = erp_outsourcing_qty_check[(
                (erp_outsourcing_qty_check['销售客户'].str.contains('合计'))
                & ((erp_outsourcing_qty_check['结存']) != 0)
        )]
        if erp_outsourcing_qty_check_month_err.empty:
            print("\n查看: 产品数量结存-月度结存--已清零--~~~OK~~~")
        else:
            check_ok = False
            failed_msg.append('产品数量结存: 月度结存--未清零--, 验收异常')
            print("\n查看: 产品数量结存-月度结存--未清零--!!!FAILED!!!")

            print("\n查看: 差异(按任意键继续)?")
            # input()
            error_skus = erp_outsourcing_qty_check_month_err['产品条码'].drop_duplicates().tolist()
            print(tabulate(erp_outsourcing_qty_check[(
                    (erp_outsourcing_qty_check['产品条码'].isin(error_skus))
                    & (~erp_outsourcing_qty_check['销售客户'].str.contains('合计'))
                    & (erp_outsourcing_qty_check['结存'] != 0)
            )],
                           headers=table_header, tablefmt='orgtbl'))
    else:
        print("\n文件: 产品数量结存 -- 无数据, 验收异常!!!FAILED!!!")

    save_file_cur_dir = os.path.dirname(save_file_cur)
    save_file_cur_name = os.path.basename(save_file_cur)
    save_file_cur_name = str.format('{}-财务ERP对比-{}.xlsx',
                                    save_file_cur_name.split('.')[0], time.strftime("%H%M%S", time.localtime()))
    if check_ok:
        save_file_cur = os.path.join(save_file_cur_dir, str.format('OK-{}', save_file_cur_name))
        print("\n\n文件: 验收~~~OK~~~\n\n")
        check_result = str.format('{}\n{}', os.path.basename(save_file_cur), '\033[32;5m验收~~~OK~~~\033[0m')
    else:
        save_file_cur = os.path.join(save_file_cur_dir, str.format('FAILED-{}', save_file_cur_name))
        print("\n\n文件: 验收!!!FAILED!!!\n\n")
        check_result = str.format('{}\n{}\n{}', os.path.basename(save_file_cur),
                                  '\033[31;5m验收!!!FAILED!!!\033[0m', ';\n'.join(failed_msg))

    print('\n导出对比结果...')
    with pd.ExcelWriter(save_file_cur) as xlsx:
        cg_compare.to_excel(xlsx, sheet_name="采购进销明细对比")
        cg1_compare.to_excel(xlsx, sheet_name="内部销售对比")
        two_b_compare.to_excel(xlsx, sheet_name="2B销售明细对比")
        dp_compare.to_excel(xlsx, sheet_name="店铺销售对比")
        erp_sale_weighting.to_excel(xlsx, sheet_name="销售出库透视测算加权")
        erp_outsourcing_qty.to_excel(xlsx, sheet_name="产品数量结存")
        # erp_product_purchase_not_sold.to_excel(xlsx, sheet_name = "产品-采购未销售")
        # erp_product_sale_not_purchased.to_excel(xlsx, sheet_name = "产品-销售未采购")
        # mega_product_sold_erp_not_purchased.to_excel(xlsx, sheet_name = "产品-Mega销售ERP未采购")
        # erp_product_purchased_mega_not_sold.to_excel(xlsx, sheet_name = "产品-ERP采购Mega未销售")
        df_diff.to_excel(xlsx, sheet_name="产品-ERP与MEGA的采购销售差异")
        # worksheet = xlsx.sheets['采购进销明细对比']
        # worksheet.set_column("B:Z", 25)

    print(str.format('\n对比结果文件: {}', save_file_cur))

    return check_result


# 主函数
def main():
    """
    main
    """

    print("\n**********开始文件选择**********")
    print("\n-----1/2请选择需要对比的'采购进销明细'、'2B销售明细'、'店铺销售'、'库存移动_麦凯莱_公司关联交易'所在的文件夹")
    print("-----请选择确保文件夹内只有三份文件，且名称包含以下关键字: ")
    print("----------------------------------------------- 采购进销")
    print("----------------------------------------------- 2B")
    print("----------------------------------------------- 店铺")
    print("----------------------------------------------- 关联交易")
    # basic_files_dir = filedialog.askdirectory()
    basic_files_dir = r'D:\f_r-max\contrast\bi_2022\mov_accept\_basic_2022'
    print("已选择的'财务底稿'文件夹:", basic_files_dir)

    basic_files = os.listdir(basic_files_dir)
    print('总共 {} 个 文件'.format(basic_files.__len__()))

    jx_file_dirs = []
    tb_file_dirs = []
    dp_file_dirs = []
    relation_file_dirs = []
    for file in basic_files:
        print('文件: {}'.format(file))
        file_dir = '{}{}{}'.format(basic_files_dir, os.sep, file)
        if file.__contains__('采购进销'):
            jx_file_dirs.append(file_dir)
        if file.__contains__('2B'):
            tb_file_dirs.append(file_dir)
        if file.__contains__('店铺'):
            dp_file_dirs.append(file_dir)
        if file.__contains__('关联交易'):
            relation_file_dirs.append(file_dir)

    if jx_file_dirs.__len__() == 0:
        print("缺少'采购进销明细'文件!!!")
        raise 'error'
    if tb_file_dirs.__len__() == 0:
        print("缺少'2B销售明细'文件!!!")
        raise 'error'
    if dp_file_dirs.__len__() == 0:
        print("缺少'店铺销售'文件!!!")
        raise 'error'
    if relation_file_dirs.__len__() == 0:
        print("缺少'库存移动_麦凯莱_公司关联交易'文件!!!")
        raise 'error'

    print("\n-----2/2:请选择需要对比的'库存移动'所在的文件夹")
    # mov_files_dir = filedialog.askdirectory()
    mov_files_dir = r'D:\f_r-max\contrast\bi_2022\mov_accept\_tmp'
    print("已选择的'移动库存'文件夹:", mov_files_dir)

    mov_files = os.listdir(mov_files_dir)
    print('总共 {} 个 文件'.format(mov_files.__len__()))

    mov_file_dirs = []
    for file in mov_files:
        print('文件: {}'.format(file))
        file_dir = '{}{}{}'.format(mov_files_dir, os.sep, file)
        mov_file_dirs.append(file_dir)

    if mov_file_dirs.__len__() == 0:
        print("缺少'库存移动'文件!!!")
        raise 'error'
    print("\n**********结束文件选择**********")

    print("\n**********开始处理整合财务底稿文件**********")

    print("\n读取Excel数据中，请稍等...")
    print('读取-采购进销明细表数据...')
    pd_cg = read_excel(jx_file_dirs, 0)
    print('读取-2B表数据...')
    pd_two_b = read_excel(tb_file_dirs, 0)
    print('读取-店铺销售数据...')
    pd_dp = read_excel(dp_file_dirs, 1)
    print('读取-库存移动_麦凯莱_公司关联交易...')
    pd_relation = read_excel(relation_file_dirs, 0)

    # 规范公司主体名称
    pd_cg['公司主体'] = pd_cg['公司主体'].astype(str)
    pd_cg['公司主体'] = pd_cg['公司主体'].apply(lambda x: x.replace("(", "（").strip())
    pd_cg['公司主体'] = pd_cg['公司主体'].apply(lambda x: x.replace(")", "）").strip())
    pd_cg['供应商名称'] = pd_cg['供应商名称'].astype(str)
    pd_cg['供应商名称'] = pd_cg['供应商名称'].apply(lambda x: x.replace("(", "（").strip())
    pd_cg['供应商名称'] = pd_cg['供应商名称'].apply(lambda x: x.replace(")", "）").strip())

    pd_two_b['主体'] = pd_two_b['主体'].astype(str)
    pd_two_b['主体'] = pd_two_b['主体'].apply(lambda x: x.replace("(", "（").strip())
    pd_two_b['主体'] = pd_two_b['主体'].apply(lambda x: x.replace(")", "）").strip())
    pd_two_b['购买方名称-销售客户'] = pd_two_b['购买方名称-销售客户'].astype(str)
    pd_two_b['购买方名称-销售客户'] = pd_two_b['购买方名称-销售客户'].apply(lambda x: x.replace("(", "（").strip())
    pd_two_b['购买方名称-销售客户'] = pd_two_b['购买方名称-销售客户'].apply(lambda x: x.replace(")", "）").strip())

    pd_dp['主体'] = pd_dp['主体'].astype(str)
    pd_dp['主体'] = pd_dp['主体'].apply(lambda x: x.replace("(", "（").strip())
    pd_dp['主体'] = pd_dp['主体'].apply(lambda x: x.replace(")", "）").strip())

    pd_relation['公司'] = pd_relation['公司'].astype(str)
    pd_relation['公司'] = pd_relation['公司'].apply(lambda x: x.replace("(", "（").strip())
    pd_relation['公司'] = pd_relation['公司'].apply(lambda x: x.replace(")", "）").strip())
    pd_relation['采购供应商'] = pd_relation['采购供应商'].astype(str)
    pd_relation['采购供应商'] = pd_relation['采购供应商'].apply(lambda x: x.replace("(", "（").strip())
    pd_relation['采购供应商'] = pd_relation['采购供应商'].apply(lambda x: x.replace(")", "）").strip())
    pd_relation['销售客户'] = pd_relation['销售客户'].astype(str)
    pd_relation['销售客户'] = pd_relation['销售客户'].apply(lambda x: x.replace("(", "（").strip())
    pd_relation['销售客户'] = pd_relation['销售客户'].apply(lambda x: x.replace(")", "）").strip())

    if not pd_relation.columns.__contains__('销售已交货含税金额(RMB)'):
        pd_relation['销售已交货含税金额(RMB)'] = pd_relation['销售已交货含税金额(原币别)']

    print("\n**********结束处理整合财务底稿文件**********")

    print("\n**********开始处理ERP移动库存数据**********")

    print('读取-ERP移动库存数据...')
    limit_max_price = 600
    contrast_result = {}
    max_price_dic = {}
    for dir_mov in mov_file_dirs:
        file_name = os.path.basename(dir_mov)
        print(str.format('\n\n读取-ERP移动库存数据...: {}', dir_mov))
        pd_erp = pd.read_excel(dir_mov, sheet_name='Result 1')

        pd_erp['公司'] = pd_erp['公司'].astype(str)
        pd_erp['公司'] = pd_erp['公司'].apply(lambda x: x.replace("(", "（").strip())
        pd_erp['公司'] = pd_erp['公司'].apply(lambda x: x.replace(")", "）").strip())
        pd_erp['采购供应商'] = pd_erp['采购供应商'].astype(str)
        pd_erp['采购供应商'] = pd_erp['采购供应商'].apply(lambda x: x.replace("(", "（").strip())
        pd_erp['采购供应商'] = pd_erp['采购供应商'].apply(lambda x: x.replace(")", "）").strip())
        pd_erp['销售客户'] = pd_erp['销售客户'].astype(str)
        pd_erp['销售客户'] = pd_erp['销售客户'].apply(lambda x: x.replace("(", "（").strip())
        pd_erp['销售客户'] = pd_erp['销售客户'].apply(lambda x: x.replace(")", "）").strip())

        if not pd_erp.columns.__contains__('销售已交货含税金额(RMB)'):
            pd_erp['销售已交货含税金额(RMB)'] = pd_erp['销售已交货含税金额(原币别)']

        ## 排序
        pd_erp_limit_prices = pd_erp[(pd_erp['销售含税单价(原币别)'] >= limit_max_price)]
        if not pd_erp_limit_prices.empty and pd_erp_limit_prices.shape[0] > 0:
            max_price_dic[
                file_name] = str.format(
                '\033[33;5m提示: 共计 {} 行数据 --> 销售含税单价(原币别)大于600元~~, 请确认数据是否异常\033[0m',
                pd_erp_limit_prices.shape[0])

        export_dir = str.format('{}{}', os.path.dirname(dir_mov), '_contrast_result')
        if not os.path.exists(export_dir):
            os.mkdir(export_dir)
        save_file = os.path.join(export_dir, file_name)

        result = contrast_file(pd_cg, pd_two_b, pd_dp, pd_relation, pd_erp, save_file)
        contrast_result[file_name] = result

    print("\n**********结束处理ERP移动库存数据**********")

    print("\n\n**********验收结果**********\n")

    for ky in contrast_result.keys():
        print(str.format('\n文件: {}\n验收结果: {}', ky, contrast_result[ky]))

        over_limit_price = max_price_dic.get(ky)
        if over_limit_price is not None and over_limit_price.__ne__(''):
            print('----------------------')
            print(over_limit_price)

    print("\n**********验收结果**********\n")


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
    input()
