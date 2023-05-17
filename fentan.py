import pandas as pd
import sys
import numpy as np
#
# pd.set_option('display.max_rows',xxx) # 最大行数
# pd.set_option('display.min_rows',xxx) # 最小显示行数
# pd.set_option('display.max_columns',xxx) # 最大显示列数
# pd.set_option ('display.max_colwidth',xxx) #最大列字符数
# pd.set_option( 'display.precision',2) # 浮点型精度
# pd.set_option('display.float_format','{:,}'.format) #逗号分隔数字
# pd.set_option('display.float_format',  '{:,.2f}'.format) #设置浮点精度
# pd.set_option('display.float_format', '{:.2f}%'.format) #百分号格式化
# pd.set_option('plotting.backend', 'altair') # 更改后端绘图方式
# pd.set_option('display.max_info_columns', 200) # info输出最大列数
# pd.set_option('display.max_info_rows', 5) # info计数null时的阈值
# pd.describe_option() #展示所有设置和描述
# pd.reset_option('all') #重置所有设置选项

pd.set_option('display.width', 5000000)
pd.set_option('display.max_colwidth', 5000)
pd.set_option('display.max_columns', 5000)
pd.set_option('display.max_rows', 5000000)

pd.set_option('mode.chained_assignment', None)


def jisuan_shuilv(company_zhuti):
    # 2021年未税收入和成本汇总（不包含21年发出商品）
    df = pd.read_excel('2021年订单回款-明细表（20220327-1稿）.xlsx', sheet_name='2021年未税收入和成本汇总（不包含21年发出商品）', skiprows=1)
    return (df[df['主体'] == company_zhuti]['税率'].to_list()[0])


def get_zhuti(a, b):
    # 2021年未税收入和成本汇总（不包含21年发出商品）
    df = pd.read_excel('2021年订单回款-明细表（20220327-1稿）.xlsx')
    print(df.head(5))
    print(df[(df['店铺'] == a) & (df['支付日期'] == b)]['主体'].to_list()[0])


def tongjibaobiao():
    df3000 = pd.read_excel('2021年订单回款-明细表（20220327-1稿）.xlsx', sheet_name='2021年未税收入和成本汇总（不包含21年发出商品）', skiprows=1)
    df = pd.read_excel('2021年订单回款-明细表（20220327-1稿）.xlsx')
    df1 = (df[df['商家编码'].str.startswith('111')])
    df1111 = df1.groupby(['店铺', '支付日期'])[['实际收款']].sum().reset_index()
    # print(df1111)
    df1111.to_excel('商家编码111打头的.xlsx')
    df122 = df.drop(df[df['商家编码'].str.startswith('111')].index)
    df122.to_excel('需要分摊的数据.xlsx')


def a():
    df5000 = pd.read_excel(r'D:\work\帐单对比\fentan\21勃狄（深圳）化妆品有限公司2C销售订单.xlsx')
    df1111 = pd.read_excel('商家编码11111111打头的.xlsx')
    df3000 = pd.read_excel('需要分摊的数据.xlsx')
    df1000 = pd.DataFrame()
    df2000 = pd.DataFrame()
    df6000 = pd.DataFrame()
    for index, row in df1111.iterrows():
        dianpu = row[1]
        pay_date = row[2]
        fentan_total = row[3]
        df100 = (df3000[(df3000['店铺'] == dianpu) & (df3000['支付日期'] == pay_date)])
        if len(df100) != 0:
            startswith111_total_huikuai = fentan_total
            num1 = len(df100)
            df200 = (df5000[(df5000['店铺'] == dianpu) & (df5000['支付日期'] == pay_date)])
            df200 = df200[df200['商家编码'].str.startswith('111')]
            df6000 = pd.concat([df6000, df200])
            fentan = startswith111_total_huikuai / num1
            df100['总回款'] = df100['总回款'] + fentan
            df100['实际收款'] = df100['实际收款'] + fentan
            df100['平均价格'] = df100['总回款'] / df100['实际卖出']
            df100['计算实际收款'] = df100['平均价格'] * df100['实际卖出']
            df100['税率'] = float(jisuan_shuilv('茉小桃（深圳）化妆品有限公司'))
            df100['不含税实际收款'] = df100['计算实际收款'] / (1 + float(jisuan_shuilv('茉小桃（深圳）化妆品有限公司')))
            df100['70%成本'] = df100['不含税实际收款'] * 0.7
            df100['成本含税'] = df100['70%成本'] * (1 + float(jisuan_shuilv('茉小桃（深圳）化妆品有限公司')))
            df100['成本含税单价'] = df100['成本含税'] / df100['实际卖出']
            df1000 = pd.concat([df1000, df100])
        elif len(df100) == 0:
            df200 = (df5000[(df5000['店铺'] == dianpu) & (df5000['支付日期'] == pay_date)])
            # print(df200)
            df2000 = pd.concat([df2000, df200])
    df1000.to_excel('分摊后的数据.xlsx')
    df2000.to_excel('无法分摊的数据.xlsx')
    df6000.to_excel('可以分摊的数据.xlsx')


def b():
    df5000 = pd.read_excel('2021年订单回款-明细表（20220327-1稿）.xlsx')
    df1111 = pd.read_excel('商家编码111打头的.xlsx')
    df1000 = pd.DataFrame()
    for index, row in df1111.iterrows():
        dianpu = row[1]
        pay_date = row[2]
        # df = df.drop(df[df.score < 50].index)
        df5000 = df5000.drop(df5000[(df5000['店铺'] == dianpu) & (df5000['支付日期'] == pay_date)].index)
        # print (df100)
    df1000 = pd.concat([df1000, df5000])
    df1000.to_excel('与111打头无关的数据.xlsx')


def c1():
    df1 = pd.read_excel('与111打头无关的数据.xlsx')
    df2 = pd.read_excel('分摊后的数据.xlsx')
    df3 = pd.read_excel('无法分摊的数据.xlsx')
    df = [df1, df2, df3]
    df100 = pd.concat(df)
    df100.to_excel('final.xlsx')


# def cc():
#     # df122['总回款'] = df122['总回款'].astype(float)
#     for index, row in df1111.iterrows():
#         dianpu = row[0]
#         pay_date = row[1]
#         print(dianpu,pay_date)
#         # df100 = (df122[(df122['店铺'] == dianpu) & (df122['支付日期'] == pay_date)])
#         df200 = (df1111[(df1111['店铺'] == dianpu) & (df1111['支付日期'] == pay_date)])
#         startswith111_total_huikuai = float(row[2])
#         # print(startswith111_total_huikuai)
#         # #求出商家编码不是以111打头的单据的数量
#         num1 = len(df122)
#         #计算分摊金额
#         fentan_cash = startswith111_total_huikuai / num1
#         # print (fentan_cash)
#
#         df100 = df122[(df122['店铺'] == dianpu) & (df122['支付日期'] == pay_date)]
#
#         print(df122.dtypes)
#
#         if len(df100) > 1:
#             startswith111_total_huikuai = float(row[2])
#             fentan_cash = startswith111_total_huikuai / len(df100)
#             print(startswith111_total_huikuai,fentan_cash)
#             print(df122[(df122['店铺'] == dianpu) & (df122['支付日期'] == pay_date)]['总回款'])
#             df122[(df122['店铺'] == dianpu) & (df122['支付日期'] == pay_date)]['总回款'] = df122[(df122['店铺'] == dianpu) & (df122['支付日期'] == pay_date)]['总回款'] + float(1)
#             print(df100)
#         #     # print (df100['总回款'])
#         #     print(fentan_cash)
#         #     df122[(df122['店铺'] == dianpu) & (df122['支付日期'] == pay_date)]['总回款'] = df122[(df122['店铺'] == dianpu) & (df122['支付日期'] == pay_date)]['总回款'] + fentan_cash
#         #     print (df122[(df122['店铺'] == dianpu) & (df122['支付日期'] == pay_date)]['总回款'])
#             sys.exit()
#         # print(row[0],row[1],fentan_cash)


#     df122[(df122['店铺'] == dianpu) & (df122['支付日期'] == pay_date)]['总回款'] = df122[(df122['店铺'] == dianpu) & (df122['支付日期'] == pay_date)]['总回款'] + fentan_cash
#     df122[(df122['店铺'] == dianpu) & (df122['支付日期'] == pay_date)]['实际收款'] = df122[(df122['店铺'] == dianpu) & (df122['支付日期'] == pay_date)]['实际收款'] + fentan_cash
#     df122[(df122['店铺'] == dianpu) & (df122['支付日期'] == pay_date)]['平均价格'] = df122[(df122['店铺'] == dianpu) & (df122['支付日期'] == pay_date)]['总回款'] / df122['实际卖出']
#     df122[(df122['店铺'] == dianpu) & (df122['支付日期'] == pay_date)]['计算实际收款'] = df122[(df122['店铺'] == dianpu) & (df122['支付日期'] == pay_date)]['平均价格'] * df122['实际卖出']
#     # print (jisuan_shuilv('茉小桃（深圳）化妆品有限公司'))
#     # df122[(df122['店铺'] == dianpu) & (df122['支付日期'] == pay_date)]['税率'] = float(jisuan_shuilv(get_zhuti(dianpu,pay_date)))
#     df122[(df122['店铺'] == dianpu) & (df122['支付日期'] == pay_date)]['税率'] = 0.13
#     df122[(df122['店铺'] == dianpu) & (df122['支付日期'] == pay_date)]['不含税实际收款'] = df122[(df122['店铺'] == dianpu) & (df122['支付日期'] == pay_date)]['计算实际收款'] / (1 + 0.13)
#     df122[(df122['店铺'] == dianpu) & (df122['支付日期'] == pay_date)]['70%成本'] = df122[(df122['店铺'] == dianpu) & (df122['支付日期'] == pay_date)]['不含税实际收款'] * 0.7
#     df122[(df122['店铺'] == dianpu) & (df122['支付日期'] == pay_date)]['成本含税'] = df122[(df122['店铺'] == dianpu) & (df122['支付日期'] == pay_date)]['70%成本'] * (1 + float(jisuan_shuilv('茉小桃（深圳）化妆品有限公司')))
#     df122[(df122['店铺'] == dianpu) & (df122['支付日期'] == pay_date)]['成本含税单价'] = df122[(df122['店铺'] == dianpu) & (df122['支付日期'] == pay_date)]['成本含税'] / df122[(df122['店铺'] == dianpu) & (df122['支付日期'] == pay_date)]['实际卖出']
# df122.to_excel('final.xlsx')


#     print(row)
#     dianpu = row[1]
#     pay_date = row[1]
#     df122 = df.drop(df[df['商家编码'].str.startswith('111')].index)
#     df122.to_excel('需要分摊的数据.xlsx')
# print (df[df['商家编码'].str.startswith('111')])
# Columns: [Unnamed: 0, 主体, 店铺, 平台, 品牌, 类别, 支付日期, 回款日期, 商家编码, 商品名称, 总回款, 总退款, 实际收款, 回款产品, 退款产品, 实际卖出, 平均价格, 最新更新时间, 订单时间, filename]
# 找出商家编码是111打头的单据
# df1 = (df[df['商家编码'].str.startswith('111')])
# # df1.to_excel('111打头的.xlsx')
# # 循环所有商家编码是111打头的单据，找到店铺和支付日期，根据店铺和支付日期找到原表中相同的单据
# df1000 = pd.DataFrame()
# for index, row in df1.iterrows():
#     dianpu = row[2]
#     pay_date = row[6]
#
#     df100 =  (df[(df['店铺']==dianpu) & (df['支付日期']==pay_date)])
#     if len(df100) == 1:
#         continue
#     if df100['商家编码'].str.startswith('111').all():
#         continue
#     # df100.to_excel('d.xlsx')
#     df100['总回款'] = df100['总回款'].astype(float)
#     startswith111_total_huikuai = df100[df100['商家编码'].str.startswith('111')]['总回款'].sum()
#     print(startswith111_total_huikuai)
#     df200 = df100[~df100['商家编码'].str.startswith('111')]
#     # #求出商家编码不是以111打头的单据的数量
#     num1 = len(df200)
#     #计算分摊金额
#     fentan_cash = startswith111_total_huikuai / num1
#     print(fentan_cash)
#     df200['总回款'] = df200['总回款'] + fentan_cash
#     df200['实际收款'] = df200['实际收款'] + fentan_cash
#     df200['平均价格'] = df200['总回款'] / df200['实际卖出']
#     df200['计算实际收款'] = df200['平均价格'] * df200['实际卖出']
#     # print (jisuan_shuilv('茉小桃（深圳）化妆品有限公司'))
#     df200['税率'] = float(jisuan_shuilv('茉小桃（深圳）化妆品有限公司'))
#     df200['不含税实际收款'] = df200['计算实际收款'] / (1 + float(jisuan_shuilv('茉小桃（深圳）化妆品有限公司')))
#     df200['70%成本'] = df200['不含税实际收款'] * 0.7
#     df200['成本含税'] = df200['70%成本'] * (1 + float(jisuan_shuilv('茉小桃（深圳）化妆品有限公司')))
#     df200['成本含税单价'] = df200['成本含税'] / df200['实际卖出']
#     print(df200)
#     df1000 = pd.concat([df1000,df200])
#
# df1000['序号'] =  range(1,len(df1000)+1,1)
# df1000.set_index('序号',inplace=True)
# df1000.to_excel('lizi2.xlsx')
# sys.exit()
#     df200 = df100[~df100['商家编码'].str.startswith('111')]
#     # print(df200)
#     # #求出商家编码不是以111打头的单据的数量
#     num1 = len(df200)
#     # print(startswith111_total_huikuai,num1)
#     #计算分摊金额
#     fentan_cash = startswith111_total_huikuai / num1
#     # print(fentan_cash)
#     df200['总回款'] = df200['总回款'] + fentan_cash
#     df200['实际收款'] = df200['实际收款'] + fentan_cash
#     df200['平均价格'] = df200['总回款'] / df200['实际卖出']
#     df200['计算实际收款'] = df200['平均价格'] * df200['实际卖出']
#     # print (jisuan_shuilv('茉小桃（深圳）化妆品有限公司'))
#     df200['税率'] = float(jisuan_shuilv('茉小桃（深圳）化妆品有限公司'))
#     df200['不含税实际收款'] = df200['计算实际收款'] / (1 + float(jisuan_shuilv('茉小桃（深圳）化妆品有限公司')))
#     df200['70%成本'] = df200['不含税实际收款'] * 0.7
#     df200['成本含税'] = df200['70%成本'] * (1 + float(jisuan_shuilv('茉小桃（深圳）化妆品有限公司')))
#     df200['成本含税单价'] = df200['成本含税'] / df200['实际卖出']
#     print(df200)
#     # df200.to_excel('lizi2.xlsx')


if __name__ == '__main__':
    # jisuan_shuilv()
    c1()
    # cc('宅星人旗舰店','2021-12-21')
