from datetime import time
from random import sample
from time import sleep
import pandas as pd


def SSQ(try_nums):
    """双色球随机选号"""
    dic_li = []
    for i in range(try_nums):
        red_num_li = sample(range(1, 34), 6)
        blue_num_li = sample(range(1, 17), 1)
        red_num_li.sort()  # 从小到大排序
        format_li(red_num_li, blue_num_li)


def DLT(try_nums):
    """大乐透随机选号"""
    for i in range(try_nums):
        red_num_li = sample(range(1, 36), 5)
        blue_num_li = sample(range(1, 13), 2)
        red_num_li.sort()
        blue_num_li.sort()
        format_li(red_num_li, blue_num_li, 'dlt')


def format_li(red_num_li, blue_num_li, type='ssq'):
    """格式化列表，美观显示列表"""
    for i in red_num_li:
        idx = red_num_li.index(i)
        if i < 10:
            red_num_li.remove(i)  # 替换一位数的元素为两位数（前面加0）
            red_num_li.insert(idx, '0{}'.format(i))

    for i in blue_num_li:
        idx = blue_num_li.index(i)
        if i < 10:
            blue_num_li.remove(i)  # 替换一位数的元素为两位数（前面加0）
            blue_num_li.insert(idx, '0{}'.format(i))

    if type == 'ssq':
        print('{},{},{},{},{},{} + {}'.format(red_num_li[0], red_num_li[1], red_num_li[2], red_num_li[3], red_num_li[4],
                                              red_num_li[5], blue_num_li[0]))
    else:
        print('{},{},{},{},{} + {},{}'.format(red_num_li[0], red_num_li[1], red_num_li[2], red_num_li[3], red_num_li[4],
                                              blue_num_li[0], blue_num_li[1]))


def run():
    while True:
        print('\n[彩票随机选号器]选项：1 双色球  2 大乐透  3 退出')
        type = input('请选择彩票类型（输入数字）:')
        if type not in ['1', '2', '3']:
            print('选项有误，请重新输入！')
            continue
        if type == '3':
            print('再见！')
            sleep(1)
            break

        try:
            try_nums = int(input('请输入注数:'))
            if try_nums <= 0:
                try_nums = 1
        except:
            try_nums = 1

        if type == '1':
            SSQ(try_nums)
        elif type == '2':
            DLT(try_nums)


def save_data_to_excel(type, dict_li):
    """把数据保存到Excel中"""
    try:
        print('正在保存数据...')
        writer = pd.ExcelWriter('{}{}.xlsx'.format(type, time.strftime("%Y-%m-%d %H_%M_%S", time.localtime())))
        df1 = pd.DataFrame(data=dict_li)  # 在同一列不同行写入数据
        df1.to_excel(writer, sheet_name='表1', )
        writer.close()
        print('保存数据已完成')
    except Exception as e:
        print('保存数据失败，错误信息：{}'.format(e))


if __name__ == '__main__':
    run()
