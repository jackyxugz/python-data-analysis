import sys, os
import itertools

import pandas as pd
from sqlalchemy import create_engine
import pymysql
import time
import openpyxl

pymysql.install_as_MySQLdb()


OMS_PROD_DB_NAME = "megaorderbill"
OMS_PROD_DB_USER = "chenxiaoselect"
OMS_PROD_DB_PWD = "NTC1abr2tqa6bev-hmr"
OMS_PROD_DB_HOST = "megaoms.rwlb.rds.aliyuncs.com"
OMS_PROD_DB_PORT = "3306"


LOCAL_FILE_PATH = r'C:\odoo报表'

bill_year = 2021
bill_month = range(1, 13)


def download_bill_record(p_engine, p_bill_year, p_bill_month):
    select_sql = r"""select concat({bill_year_show},'年') 年,concat({bill_month_show},'月') 月份,comp.name as 主体,
                    shop.SHOP_NAME 店铺名称,bill.platform1 平台,'=TEXTJOIN("-",FALSE,B2,C2,D2,E2)' 原拼接,'' 与名称更新对应的拼接,bill.INCOME_AMOUNT1 回款金额,
                    bill.EXPEND_AMOUNT1 退款金额,0 海外回款金额,0 海外退款金额,
                    bill.INCOME_AMOUNT1+bill.EXPEND_AMOUNT1 收入金额,
                    (bill.INCOME_AMOUNT1+bill.EXPEND_AMOUNT1)/1.13 收入金额（不含税）,bill.SHOPCODE
                    from (
                        select (case platform when 'TAOBAO' then '淘宝'
                                              when 'TMALL' then '天猫'
                                              when 'DY' then '抖音'
                                              when 'JD' then '京东'
                                              when 'JN' then '金牛'
                                              when 'KAOLA' then '网易考拉'
                                              when 'KS' then '快手' 
                                              when 'PDD' then '拼多多'
                                              when 'WPH' then '唯品会'
                                              when 'XHS' then '小红书'
                                              when 'YZ' then '有赞'
                                              when 'ALIBABA' then '阿里巴巴'
                                              when 'WM' then '微盟'
                                              when 'FY' then '枫叶小店'
                                              when 'MD' then '马到'
                                              when 'BD' then '百度小店'
                                              when 'ZMB' then '做梦吧'
                            else platform end) platform1,
                            platform,
                            SHOPCODE,
                            TRADING_CHANNELS,
                            sum(case when IS_REFUNDAMOUNT = 1 then EXPEND_AMOUNT else 0 end) EXPEND_AMOUNT1,
                            sum(case when IS_AMOUNT = 1 then INCOME_AMOUNT else 0 end) INCOME_AMOUNT1
                        from `megaorder-{v_bill_year}`.order_info_bill_{v_bill_month}   
                        where ((IS_REFUNDAMOUNT=1)
                            or (is_amount = 1))
                        group by platform, SHOPCODE 
                    ) bill
                    left join megaorder.new_shop_info shop on shop.SHOP_CODE=bill.SHOPCODE and shop.PLATFORM=bill.PLATFORM and shop.STATUS =0
                    left join megaorder.new_shop_company_info comp on shop.shop_company_id=comp.id
                    left join megaorder.shop_company_history ht on  ht.shop_company_id=comp.id 
                    
                    ;""".format(
        bill_year_show=p_bill_year, bill_month_show=p_bill_month,v_bill_year=str(p_bill_year).zfill(4), v_bill_month=str(p_bill_month).zfill(2))
    try:
        read_begin_time = time.time()
        df = pd.read_sql_query(select_sql, p_engine)
        read_end_time = time.time()
        print(p_bill_year,'-', p_bill_month, '读取数据用时:', read_end_time - read_begin_time,
              '秒')

    except Exception as e:
        print('select_sql:\n', select_sql)
        print(e)
        raise 'error'

    return df


if __name__ == '__main__':
    OMS_engine = create_engine(
        'mysql://{}:{}@{}:{}/{}'.format(OMS_PROD_DB_USER, OMS_PROD_DB_PWD, OMS_PROD_DB_HOST, OMS_PROD_DB_PORT,
                                        OMS_PROD_DB_NAME),
        echo=True,
        isolation_level='AUTOCOMMIT')

    # OMS_engine = create_engine(
    #     'mysql://{}:{}@{}:{}/{}'.format(OMS_DB_USER, OMS_DB_PWD, OMS_DB_HOST, OMS_DB_PORT,
    #                                     OMS_DB_NAME),
    #     echo=True,
    #     isolation_level='AUTOCOMMIT')


    print('开始计算...')
    all_begin_time = time.time()
    dftotal = pd.DataFrame()
    filename = LOCAL_FILE_PATH + os.sep + 'bill-'+str(bill_year)+'_0303' + '.xlsx'
    # filename = LOCAL_FILE_PATH + 'bill_orderdetail-'+str(bill_year)+'_1217' + '.xlsx'
    if os.path.isfile(filename):
        os.remove(filename)

    for val in itertools.product(bill_month):
        df = pd.DataFrame()
        df = download_bill_record(OMS_engine, bill_year, val[0])
        # df = download_bill_orderdetail_record(OMS_engine, bill_year, val[0])
        dftotal = dftotal.append(df)

    dftotal.to_excel(filename, index=False,sheet_name='oms')


    all_end_time = time.time()
    print('计算完成，从', bill_year, str(bill_month[0]).zfill(2), ' - ', bill_year,
          str(bill_month[len(bill_month) - 1]).zfill(2), '总消耗时间:', all_end_time - all_begin_time, '秒')


