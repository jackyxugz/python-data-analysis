from sqlalchemy import create_engine
import time
from functools import wraps
import threading

# from db_info import JXC_DB_NAME, JXC_DB_USER, JXC_DB_PWD, JXC_DB_HOST, JXC_DB_PORT

JXC_DB_NAME = 'run_report_21_qjl_01'
JXC_DB_USER = 'openpg'
JXC_DB_PWD = 'odoo12'
JXC_DB_HOST = '10.10.121.109'
JXC_DB_PORT = '5432'


def output_time_spend(func):
    @wraps(func)  # 保持原函数名不变
    def wrapper(*args, **kwargs):
        wrap_time_start = time.time()
        func(*args, **kwargs)
        wrap_time_end = time.time()
        if int(wrap_time_end - wrap_time_start) > 2:
            print('函数 {}{} 消耗时间：'.format(func.__name__, str(args)), (wrap_time_end - wrap_time_start), '秒')

    return wrapper


@output_time_spend
def run_procedural_no_llc(conn, v_year, v_month, v_warehouse_id, v_procedural_name):
    proc_sql = """call {}({},{},{});""".format(v_procedural_name, v_year, v_month, v_warehouse_id)
    conn.execute(proc_sql)


@output_time_spend
def run_procedural_with_llc(conn, v_year, v_month, v_warehouse_id, v_llc, v_procedural_name):
    proc_sql = """call {}({},{},{},{});""".format(v_procedural_name, v_year, v_month, v_warehouse_id, v_llc)
    conn.execute(proc_sql)


@output_time_spend
def run_procedural_no_llc_up(conn, v_year, v_month, v_warehouse_id, v_procedural_name, iyear_up, imonth_up):
    proc_sql = """call {}({},{},{},{},{});""".format(v_procedural_name, v_year, v_month, v_warehouse_id, iyear_up,
                                                     imonth_up)
    conn.execute(proc_sql)


@output_time_spend
def run_procedural_with_llc_up(conn, v_year, v_month, v_warehouse_id, v_llc, v_procedural_name, iyear_up, imonth_up):
    proc_sql = """call {}({},{},{},{},{},{});""".format(v_procedural_name, v_year, v_month, v_warehouse_id, v_llc,
                                                        iyear_up, imonth_up)
    conn.execute(proc_sql)


def get_diff(conn, v_year, v_month, v_warehouse_id, v_llc):
    select_sql = """select count(*) from
        (
            select inbound.production_id, abs(inbound.inbound_amount - outbound.outbound_amount) diff
            from
            (
                select spi.production_id,spi.ccp_qty_done as ccp_qty_done,
                    sum(sm_inbound.m_qty_done * sm_inbound.m_jxc_get_price) as inbound_amount
                from (
                    select distinct production_id, ccp_qty_done
                    from sm_production_in
                    where iyear = {v_year}
                        and imonth = {v_month}
                        and ccp_warehouse = {v_warehouse_id}
                        and llc ={v_llc}
                ) spi
                left join stock_move sm_inbound
                    on sm_inbound.production_id = spi.production_id
                        and sm_inbound.m_business_type = 'production_in'
                        and sm_inbound.bvalid = 1
                        and sm_inbound.iyear = {v_year}
                        and sm_inbound.imonth = {v_month}
                        and sm_inbound.warehouse_id = {v_warehouse_id}
                group by spi.production_id, spi.ccp_qty_done
                union
                select sopi.production_id,sopi.ccp_qty_done as ccp_qty_done,
                    sum(sm_o_inbound.m_qty_done * sm_o_inbound.m_jxc_get_price) as inbound_amount
                from (
                    select distinct production_id, ccp_qty_done
                    from sm_outsource_production_in
                    where iyear = {v_year}
                        and imonth = {v_month}
                        and ccp_warehouse = {v_warehouse_id}
                        and llc ={v_llc}
                ) sopi
                left join stock_move sm_o_inbound
                    on sm_o_inbound.production_id = sopi.production_id
                        and sm_o_inbound.m_business_type = 'outsource_production_in'
                        and sm_o_inbound.bvalid = 1
                        and sm_o_inbound.iyear = {v_year}
                        and sm_o_inbound.imonth = {v_month}
                        and sm_o_inbound.warehouse_id = {v_warehouse_id}
                group by sopi.production_id, sopi.ccp_qty_done
            ) inbound
            left join
            (
                select sm_outbound.raw_material_production_id,
                    sum(abs(sm_outbound.m_qty_done * sm_outbound.m_jxc_get_price)) as outbound_amount
                from stock_move as sm_outbound
                where sm_outbound.bvalid = 1
                    and sm_outbound.m_business_type in ('raw_material_out','outsource_material_out')
                    and sm_outbound.raw_material_production_id is not null
                group by sm_outbound.raw_material_production_id
            ) outbound on inbound.production_id = outbound.raw_material_production_id
        ) aaa where diff>1;""".format(v_year=v_year, v_month=v_month, v_warehouse_id=v_warehouse_id, v_llc=v_llc)
    cur = conn.execute(select_sql)
    return cur.fetchone()


def gnrt_single_month_warehouse_llc(conn, v_year, v_month, v_warehouse_id, v_llc, iyear_up, imonth_up):
    print("年月仓库llc依次为：", v_year, v_month, v_warehouse_id, v_llc)

    def run_sql(v_procedural_name):
        run_procedural_with_llc(conn, v_year, v_month, v_warehouse_id, v_llc, v_procedural_name)

    def run_sql_no_llc(v_procedural_name):
        run_procedural_no_llc(conn, v_year, v_month, v_warehouse_id, v_procedural_name)

    def run_sql_up(v_procedural_name):
        run_procedural_with_llc_up(conn, v_year, v_month, v_warehouse_id, v_llc, v_procedural_name, iyear_up, imonth_up)

    def run_sql_no_llc_up(v_procedural_name):
        run_procedural_no_llc_up(conn, v_year, v_month, v_warehouse_id, v_procedural_name, iyear_up, imonth_up)

    if v_llc == 99:
        run_sql_no_llc_up("proc_cal_begin")
        run_sql_no_llc_up("proc_cal_pure_zero")
        run_sql_no_llc("proc_cal_purchase_in")
        run_sql_no_llc_up("proc_cal_purchase_out")
        run_sql_no_llc_up("proc_cal_inventory_in")
        run_sql_no_llc_up("proc_cal_interval_in")
        run_sql_no_llc_up("proc_cal_district_in")
        run_sql_no_llc_up("proc_cal_purchase_other_in")
        run_sql_no_llc("proc_insert_other_out_num")
        run_sql_no_llc("proc_insert_pr_in_num")
        run_sql_no_llc("proc_insert_uni_in_num")
        run_sql_no_llc_up("proc_cal_avg_price")
        run_sql_no_llc("proc_cal_sale_in_up")

    run_sql_no_llc("proc_update_raw_material_out_price")
    cycle_num = 20
    for i in range(1, cycle_num):
        run_sql_up("proc_update_pr_in_price")
        run_sql_up("proc_update_outsouce_pr_in_price")
        run_sql_no_llc("proc_update_yearmonth_price")
        run_sql("proc_update_consum_unbuild_price")
        run_sql_up("proc_update_uni_in_price")
        run_sql_no_llc("proc_update_yearmonth_price")
        run_sql_no_llc("proc_update_raw_material_out_price")
        res = get_diff(conn, v_year, v_month, v_warehouse_id, v_llc)
        if res and res[0] > 0:
            print('正在循环修正成本，第 ', i, '次...')
        else:
            if i > 1:
                print('在第', i, '次成本修正成功！')
            break

    if v_llc == 0:
        run_sql_no_llc("proc_update_other_out_price")
        run_sql_no_llc("proc_cal_sale_in")
        run_sql_no_llc("proc_cal_inventory_losses")
        run_sql_no_llc("proc_insert_begin_num_price")
        run_sql_no_llc("proc_update_in_out_num_price")
        run_sql_no_llc("proc_update_end_num_price")


def gnrt_single_month_all_warehouse_llc_one(conn, v_year, v_month, iyear_up, imonth_up, v_warehouse_id):
    for llc in [99, 6, 5, 4, 3, 2, 1, 0]:
        gnrt_single_month_warehouse_llc(conn, v_year, v_month, v_warehouse_id, llc, iyear_up, imonth_up)


def gnrt_single_month_all_warehouse_llc(conn, v_year, v_month, iyear_up, imonth_up):
    WAREHOUSE_IDS = [
        # 5,    # 麦凯莱
        # 4,    # 麦凯莱
        # 12,   # 麦凯莱
        1,  # 麦凯莱
        # 134   # 麦凯莱
        # 2,    # 卖家优选
        # 3,    # 卖家联合
        # 10,   # 艾法
        # 13,   # 鲁文
        # 17,   # 盈养泉
        # 22,   # 樱岚
        # 23,   # 多瑞
        # 27,   # 鑫桂
        # 24,  # 尚西
        # 16,   # 勃狄
        # 33,   # 精酿
        # 34,   # 睿旗
        # 35,   # 宏炽
        # 37,   # 白皮书
        # 38,   # 造白
        # 39 ,  # 可瘾
        # 41,   # 伯艾地
        # 42,   # 肯妮诗
        # 43,   # 配颜师
    ]
    for warehouse_id in WAREHOUSE_IDS:
        for llc in [99, 5, 4, 3, 2, 1, 0]:
            # for llc in [99, 3, 2, 1, 0]:
            gnrt_single_month_warehouse_llc(conn, v_year, v_month, warehouse_id, llc, iyear_up, imonth_up)


if __name__ == '__main__':
    time_all_start = time.time()
    engine1 = create_engine(
        'postgresql://{}:{}@{}:{}/{}'.format(JXC_DB_USER, JXC_DB_PWD, JXC_DB_HOST, JXC_DB_PORT, JXC_DB_NAME),
        echo=False, isolation_level='AUTOCOMMIT')
    conn1 = engine1.connect()

    engine2 = create_engine(
        'postgresql://{}:{}@{}:{}/{}'.format(JXC_DB_USER, JXC_DB_PWD, JXC_DB_HOST, JXC_DB_PORT, JXC_DB_NAME),
        echo=False, isolation_level='AUTOCOMMIT')
    conn2 = engine2.connect()
    engine3 = create_engine(
        'postgresql://{}:{}@{}:{}/{}'.format(JXC_DB_USER, JXC_DB_PWD, JXC_DB_HOST, JXC_DB_PORT, JXC_DB_NAME),
        echo=False, isolation_level='AUTOCOMMIT')
    conn3 = engine3.connect()
    engine4 = create_engine(
        'postgresql://{}:{}@{}:{}/{}'.format(JXC_DB_USER, JXC_DB_PWD, JXC_DB_HOST, JXC_DB_PORT, JXC_DB_NAME),
        echo=False, isolation_level='AUTOCOMMIT')
    conn4 = engine4.connect()
    engine5 = create_engine(
        'postgresql://{}:{}@{}:{}/{}'.format(JXC_DB_USER, JXC_DB_PWD, JXC_DB_HOST, JXC_DB_PORT, JXC_DB_NAME),
        echo=False, isolation_level='AUTOCOMMIT')
    conn5 = engine5.connect()

    month_start = 1
    month_end = 12


    def run_warehouse_one1(v_warehouse_id):
        for months in range(month_start, month_end + 1):
            time_start = time.time()
            if months == 1:
                iyear_up = 2020
                imonth_up = 12
            else:
                iyear_up = 2021
                imonth_up = months - 1

            gnrt_single_month_all_warehouse_llc_one(conn1, 2021, months, iyear_up, imonth_up, v_warehouse_id)

            time_end = time.time()
            print('=' * 50)
            print('{}: 本月耗时（{}）：'.format(v_warehouse_id, months), (time_end - time_start) / 60, '分钟')
        time_all_end = time.time()
        print(v_warehouse_id, ': ', month_start, '月 -', month_end, '月耗时：', (time_all_end - time_all_start) / 60, '分钟')


    def run_warehouse_one2(v_warehouse_id):
        for months in range(month_start, month_end + 1):
            time_start = time.time()
            if months == 1:
                iyear_up = 2020
                imonth_up = 12
            else:
                iyear_up = 2021
                imonth_up = months - 1

            gnrt_single_month_all_warehouse_llc_one(conn2, 2021, months, iyear_up, imonth_up, v_warehouse_id)

            time_end = time.time()
            print('=' * 50)
            print('{}: 本月耗时（{}）：'.format(v_warehouse_id, months), (time_end - time_start) / 60, '分钟')
        time_all_end = time.time()
        print(v_warehouse_id, ': ', month_start, '月 -', month_end, '月耗时：', (time_all_end - time_all_start) / 60, '分钟')


    def run_warehouse_one3(v_warehouse_id):
        for months in range(month_start, month_end + 1):
            time_start = time.time()
            if months == 1:
                iyear_up = 2020
                imonth_up = 12
            else:
                iyear_up = 2021
                imonth_up = months - 1

            gnrt_single_month_all_warehouse_llc_one(conn3, 2021, months, iyear_up, imonth_up, v_warehouse_id)

            time_end = time.time()
            print('=' * 50)
            print('{}: 本月耗时（{}）：'.format(v_warehouse_id, months), (time_end - time_start) / 60, '分钟')
        time_all_end = time.time()
        print(v_warehouse_id, ': ', month_start, '月 -', month_end, '月耗时：', (time_all_end - time_all_start) / 60, '分钟')


    def run_warehouse_one4(v_warehouse_id):
        for months in range(month_start, month_end + 1):
            time_start = time.time()
            if months == 1:
                iyear_up = 2020
                imonth_up = 12
            else:
                iyear_up = 2021
                imonth_up = months - 1

            gnrt_single_month_all_warehouse_llc_one(conn4, 2021, months, iyear_up, imonth_up, v_warehouse_id)

            time_end = time.time()
            print('=' * 50)
            print('{}: 本月耗时（{}）：'.format(v_warehouse_id, months), (time_end - time_start) / 60, '分钟')
        time_all_end = time.time()
        print(v_warehouse_id, ': ', month_start, '月 -', month_end, '月耗时：', (time_all_end - time_all_start) / 60, '分钟')


    def run_warehouse_one5(v_warehouse_id):
        for months in range(month_start, month_end + 1):
            time_start = time.time()
            if months == 1:
                iyear_up = 2020
                imonth_up = 12
            else:
                iyear_up = 2021
                imonth_up = months - 1

            gnrt_single_month_all_warehouse_llc_one(conn5, 2021, months, iyear_up, imonth_up, v_warehouse_id)

            time_end = time.time()
            print('=' * 50)
            print('{}: 本月耗时（{}）：'.format(v_warehouse_id, months), (time_end - time_start) / 60, '分钟')
        time_all_end = time.time()
        print(v_warehouse_id, ': ', month_start, '月 -', month_end, '月耗时：', (time_all_end - time_all_start) / 60, '分钟')


    # threading.Thread(target=run_warehouse_one1, args=(5,)).start()
    # threading.Thread(target=run_warehouse_one2, args=(4,)).start()
    # threading.Thread(target=run_warehouse_one3, args=(12,)).start()
    # threading.Thread(target=run_warehouse_one4, args=(1,)).start()
    threading.Thread(target=run_warehouse_one5, args=(68,)).start()