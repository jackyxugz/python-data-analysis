# __coding=utf8__
# /** 作者：zengyanghui **/
import re
import sys
import os
import future.backports.socketserver
import pandas as pd
import numpy as np
# from datetime import datetime
import datetime
import time
import os.path
import xlrd
import xlwt
import pprint
import math
import tabulate
from selenium import webdriver


def list_all_files(rootdir, filekey_list):
    if len(filekey_list) > 0:
        filekey_list = filekey_list.replace(",", " ")
        filekey = filekey_list.split(" ")
    else:
        filekey = ''

    _files = []
    list = os.listdir(rootdir)  # 列出文件夹下所有的目录与文件
    for i in range(0, len(list)):
        path = os.path.join(rootdir, list[i])
        if os.path.isdir(path):
            _files.extend(list_all_files(path, filekey_list))
        if os.path.isfile(path):
            if ((path.find("~") < 0) and (path.find(".DS_Store") < 0)):  # 带~符号表示临时文件，不读取
                if len(filekey) > 0:
                    for key in filekey:
                        # print(path)
                        filename = "".join(path.split("\\")[-1:])
                        # print("文件名:",filename)

                        key = key.replace("！", "!")

                        if key.find("!") >= 0:
                            # print("反向选择:",key)
                            if filename.find(key.replace("!", "")) >= 0:  # 此文件不要读取
                                # print("{} 不应该包含 {}，所以剔除:".format(filename,key ))
                                pass
                        elif filename.find(key) > 0:  # 只做文件名的过滤
                            _files.append(path)

                else:
                    _files.append(path)

    # print(_files)
    return _files


def read_bill(filename):
    print(filename)
    if filename.find("xls")>=0:
        df = pd.read_excel(filename)
    else:
        df = pd.read_csv(filename)

    # 判断文件所属年份
    year = ("".join(filename.split(os.sep)[-1:]))[:2]
    print(year)

    # 截取文件公司名
    company = "".join(filename.split("_")[1:2])
    print(company)
    print(df.head(1).to_markdown())
    if "采购/退货(含税)金额(原币别)" not in df.columns:
        df["采购/退货(含税)金额(RMB)"].fillna(0,inplace=True)
        df["销售已交货含税金额(原币别)"].fillna(0,inplace=True)
        df["采购/退货(含税)金额(原币别)"] = 0

        group_df = df.groupby(by=["公司","业务类型"]).agg({"采购/退货(含税)金额(RMB)":"sum","采购/退货(含税)金额(原币别)": "sum", "销售已交货含税金额(原币别)":"sum"})
        group_df = pd.DataFrame(group_df).reset_index()
    else:
        df["采购/退货(含税)金额(RMB)"].fillna(0, inplace=True)
        df["采购/退货(含税)金额(原币别)"].fillna(0, inplace=True)
        df["销售已交货含税金额(原币别)"].fillna(0, inplace=True)

        group_df = df.groupby(by=["公司", "业务类型"]).agg({"采购/退货(含税)金额(RMB)":"sum", "采购/退货(含税)金额(原币别)": "sum", "销售已交货含税金额(原币别)": "sum"})
        group_df = pd.DataFrame(group_df).reset_index()
    print(group_df.head().to_markdown())
    group_df = group_df.loc[(group_df["业务类型"]=="采购入库")|(group_df["业务类型"]=="销售出库")]

    print(group_df.head().to_markdown())

    # 登录 odoo
    driver = webdriver.Chrome()
    driver.get("http://10.10.254.124:8069/web")
    driver.maximize_window()
    driver.implicitly_wait(5)
    if int(year) == 20:
        driver.find_element_by_link_text("jxc20_v2_07").click()
    elif int(year) == 19:
        driver.find_element_by_link_text("jxc19_v8_02_19").click()
        # driver.find_element_by_xpath("//a[normalize-space()='jxc19_v8_02_19']").click()
    driver.implicitly_wait(5)
    driver.find_element_by_id("login").send_keys("admin2")
    driver.implicitly_wait(1)
    driver.find_element_by_id("password").send_keys("1")
    driver.implicitly_wait(1)
    driver.find_element_by_xpath("/html/body/div/div/div/form/div[4]/button").click()
    driver.implicitly_wait(5)

    # 选择公司主体
    driver.find_element_by_xpath("/html/body/header/nav/ul[2]/li[3]/a").click()
    driver.implicitly_wait(5)

    if company.find("麦凯莱") >= 0:
        comp_xpath = "//a[contains(text(),'深圳市麦凯莱科技有限公司')]"
    elif company.find("白皮书") >= 0:
        comp_xpath = "//a[contains(text(),'深圳市白皮书文化传媒有限公司')]"
    elif company.find("可瘾") >= 0:
        comp_xpath = "//a[contains(text(),'可瘾（深圳）化妆品有限公司')]"
    elif company.find("卖家优选") >= 0:
        comp_xpath = "//a[contains(text(),'深圳市卖家优选实业有限公司')]"
    elif company.find("播地艾") >= 0:
        comp_xpath = "//a[contains(text(),'播地艾（广州）化妆品有限公司')]"
    elif company.find("大前海") >= 0:
        comp_xpath = "//a[contains(text(),'深圳大前海物流有限公司')]"
    elif company.find("二十四小时") >= 0:
        comp_xpath = "//a[contains(text(),'深圳市二十四小时七天商贸有限公司')]"
    elif company.find("精酿") >= 0:
        comp_xpath = "//a[contains(text(),'深圳市精酿商贸有限公司')]"
    elif company.find("鲁文") >= 0:
        comp_xpath = "//a[contains(text(),'平湖鲁文国际贸易有限公司')]"
    elif company.find("卖家联合") >= 0:
        comp_xpath = "//a[contains(text(),'深圳市卖家联合商贸有限公司')]"
    elif company.find("樱岚") >= 0:
        comp_xpath = "//a[contains(text(),'深圳樱岚护肤品有限公司')]"
    elif company.find("末隐师") >= 0:
        comp_xpath = "//a[contains(text(),'末隐师（广州）化妆品有限公司')]"
    elif company.find("造白") >= 0:
        comp_xpath = "//a[contains(text(),'造白（广州）化妆品有限公司')]"

    driver.find_element_by_xpath(comp_xpath).click()
    driver.implicitly_wait(5)
    driver.refresh()
    driver.implicitly_wait(5)

    # 查询销售额
    driver.find_element_by_xpath("//a[3]//div[1]").click()
    driver.implicitly_wait(5)
    driver.find_element_by_xpath("//div[@title='移除']").click()
    driver.implicitly_wait(5)
    xs_amount = driver.find_element_by_xpath("//td[@title='总计']").text
    print(xs_amount)
    xs_amount = xs_amount.replace(",", "")
    print(float(xs_amount))
    xs_real_amount = driver.find_element_by_xpath("//td[@title='已交货含税总金额']").text
    print(xs_real_amount)
    xs_real_amount = xs_real_amount.replace(",", "")
    print(float(xs_real_amount))
    driver.implicitly_wait(5)

    # 返回主菜单
    driver.find_element_by_xpath("//a[@title='应用']").click()
    driver.implicitly_wait(5)

    # 查询采购额
    driver.find_element_by_xpath("//a[4]//div[1]").click()
    driver.implicitly_wait(5)
    cg_amount = driver.find_element_by_xpath("//td[@title='不含金额合计']").text
    print(cg_amount)
    cg_amount = cg_amount.replace(",", "")
    print(float(cg_amount))
    cg_real_amount = driver.find_element_by_xpath("//td[@title='合计金额']").text
    print(cg_real_amount)
    cg_real_amount = cg_real_amount.replace(",", "")
    print(float(cg_real_amount))
    driver.implicitly_wait(10)

    driver.quit()

    # 核对报表记录页面金额
    group_df["odoo采购不含税总计"] = group_df["业务类型"].apply(lambda x:cg_amount if x=="采购入库" else 0)
    group_df["odoo采购总计"] = group_df["业务类型"].apply(lambda x:cg_real_amount if x=="采购入库" else 0)
    group_df["odoo销售订单总计"] = group_df["业务类型"].apply(lambda x:xs_amount if x=="销售出库" else 0)
    group_df["odoo销售已交货总计"] = group_df["业务类型"].apply(lambda x:xs_real_amount if x=="销售出库" else 0)

    group_df["odoo采购不含税总计"] = group_df["odoo采购不含税总计"].astype(float)
    group_df["odoo采购总计"] = group_df["odoo采购总计"].astype(float)
    group_df["odoo销售订单总计"] = group_df["odoo销售订单总计"].astype(float)
    group_df["odoo销售已交货总计"] = group_df["odoo销售已交货总计"].astype(float)

    # 计算金额差异
    group_df["采购差异"] = group_df.apply(lambda x:x["采购/退货(含税)金额(RMB)"] - x["odoo采购总计"] if ((x["采购入库"]=="采购入库") & (x["采购/退货(含税)金额(原币别)"] == 0)) else x["采购/退货(含税)金额(原币别)"] - x["odoo采购总计"],axis=1)
    group_df["销售差异"] = group_df["销售已交货含税金额(原币别)"] - group_df["odoo销售已交货总计"]

    group_df["采购差异"] = group_df["采购差异"].map(lambda x: "{:.2f}".format(x))
    group_df["销售差异"] = group_df["销售差异"].map(lambda x: "{:.2f}".format(x))

    return group_df


def get_all_files(rootdir, filekey):
    filelist = list_all_files(rootdir, filekey)
    # print(filelist)
    mySeries = pd.Series(filelist)
    df = pd.DataFrame(mySeries)
    df.columns = ["filename"]
    # df["filename"] = df["filename"].apply(lambda x: x.strip().replace(" ", ""))

    df = df[~df["filename"].str.contains("汇总")]

    return df


def read_all_excel(rootdir, filekey):
    df_files = get_all_files(rootdir, filekey)
    for index, file in df_files.iterrows():
        if 'df' in locals().keys():  # 如果变量已经存在
            dd = read_bill(file["filename"])
            dd["filename"] = file["filename"]
            print("进度表：{}/{}  文件{}，行数{}".format(index + 1, df_files.shape[0], file["filename"], dd.shape[0]))
            df = df.append(dd)


        else:
            df = read_bill(file["filename"])
            df["filename"] = file["filename"]
            # print(file["filename"],  df.shape[0])
            print("进度表：{}/{}  文件{}，行数{}".format(index + 1, df_files.shape[0], file["filename"], df.shape[0]))

    return df


def combine_bill():
    print('订单数据校对逻辑:')
    print('1.财务订单数据需要放在财务数据文件夹下，例如/校对数据/财务数据/...')
    print('2.导出订单数据需要放在导出数据文件夹下，例如/校对数据/导出数据/...')
    print("请输入财务订单和导出订单所在的文件夹：")
    # filedir=""
    # filedir = input()
    # myTuple = shell.SHBrowseForFolder(0, None, "", 64)
    try:
        # path = shell.SHGetPathFromIDList(myTuple[0])
        filedir = input()
    except:
        print("你没有输入任何目录 :(")
        sys.exit()
        return

    # filedir=path.decode('ansi')
    print("你选择的路径是：", filedir)

    global default_dir
    default_dir = filedir

    if len(filedir) == 0:
        print("你没有输入任何目录 :(")
        sys.exit()
        return

    print('请输入要合并的文件名称或者后缀匹配符:（比如都必须包含 序时账 三个字，那么就请输入  "序时账" ， 不输入就表示所有excel都要合并！）')
    filekey = input()

    if len(filedir) == 0:
        print("你没有输入任何关键词 :(")
        filekey = ''
        # sys.exit()
        # return

    print("你希望在'{}'目录下找到所有的包含“{}”文件，然后合并。".format(filedir, filekey))

    table = read_all_excel(filedir, filekey)

    if len(table) > 500000:
        table.to_csv(default_dir + r"\报表与前端核对结果.csv",index=False)
    else:
        table.to_excel(default_dir + r"\报表与前端核对结果.xlsx",index=False)

    # table.dropna(inplace=True)


def odoo_query():
    # option = webdriver.ChromeOptions()
    # option.add_argument(r'user-data-dir=C:\Users\sjit25\AppData\Local\Google\Chrome\User Data')
    # driver = webdriver.Chrome(options=option)
    # driver.get("http://10.10.254.124:8069/web")
    company="麦凯莱"

    # 登录 odoo
    driver = webdriver.Chrome()
    driver.get("http://10.10.254.124:8069/web")
    driver.maximize_window()
    driver.implicitly_wait(1)
    driver.find_element_by_link_text("jxc20_v2_07").click()
    driver.implicitly_wait(1)
    driver.find_element_by_id("login").send_keys("admin2")
    driver.implicitly_wait(1)
    driver.find_element_by_id("password").send_keys("1")
    driver.implicitly_wait(1)
    driver.find_element_by_xpath("/html/body/div/div/div/form/div[4]/button").click()
    driver.implicitly_wait(1)

    # 选择公司主体
    driver.find_element_by_xpath("/html/body/header/nav/ul[2]/li[3]/a").click()

    if company.find("麦凯莱")>=0:
        company == "深圳市麦凯莱科技有限公司"
    driver.find_element_by_xpath("//a[contains(text(),'深圳市麦凯莱科技有限公司')]").click()
    driver.refresh()
    driver.implicitly_wait(1)

    # 查询销售额
    driver.find_element_by_xpath("//a[3]//div[1]").click()
    driver.implicitly_wait(1)
    driver.find_element_by_xpath("//div[@title='移除']").click()
    driver.implicitly_wait(5)
    xs_amount = driver.find_element_by_xpath("//td[@title='总计']").text
    print(xs_amount)
    xs_amount = xs_amount.replace(",","")
    print(float(xs_amount))
    xs_real_amount = driver.find_element_by_xpath("//td[@title='已交货含税总金额']").text
    print(xs_real_amount)
    xs_real_amount = xs_real_amount.replace(",", "")
    print(float(xs_real_amount))
    driver.implicitly_wait(1)

    # 返回主菜单
    driver.find_element_by_xpath("//a[@title='应用']").click()

    # 查询采购额
    driver.find_element_by_xpath("//a[4]//div[1]").click()
    driver.implicitly_wait(1)
    cg_amount = driver.find_element_by_xpath("//td[@title='不含金额合计']").text
    print(cg_amount)
    oc_amount = cg_amount.replace(",", "")
    print(float(cg_amount))
    cg_real_amount = driver.find_element_by_xpath("//td[@title='合计金额']").text
    print(cg_real_amount)
    cg_real_amount = cg_real_amount.replace(",", "")
    print(float(cg_real_amount))
    driver.implicitly_wait(10)

    driver.quit()


if __name__ == "__main__":
    print("开始:", time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()))

    combine_bill()

    print("ok")