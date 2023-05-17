# __coding=utf8__
# /** 作者：Jacky.Xu **/

import pandas as pd
import numpy as np
from collections import Counter
import re
import os
import warnings
warnings.filterwarnings("ignore")
# import Tkinter
import win32api
import win32ui
import win32con
import tabulate
import math
import datetime as dt
import uuid

# 订单拆分合并
# 文件1: order.xlsx
# 文件2: bom.xlsx


# 检查目录是否存在
def mkdir(default_path):
    path1 = default_path+"/订单拆分合并/拆分/拆分-OEM加工厂"
    path2 = default_path+"/订单拆分合并/拆分/拆分-原料"
    path3 = default_path+"/订单拆分合并/拆分/拆分-购成品"
    path4 = default_path+"/订单拆分合并/拆分/拆分-包材"
    path5 = default_path+"/订单拆分合并/合并/合并订单"
    path6 = default_path+"/订单拆分合并/合并/合并仓库单据"

    isExists1 = os.path.exists(path1)
    if not isExists1:
        os.makedirs(path1)
        print(path1 + ' 创建成功')
    else:
        print(path1 + ' 目录已存在')

    isExists2 = os.path.exists(path2)
    if not isExists2:
        os.makedirs(path2)
        print(path2 + ' 创建成功')
    else:
        print(path2 + ' 目录已存在')
    
    isExists3 = os.path.exists(path3)
    if not isExists3:
        os.makedirs(path3)
        print(path3 + ' 创建成功')
    else:
        print(path3 + ' 目录已存在')
    
    isExists4 = os.path.exists(path4)
    if not isExists4:
        os.makedirs(path4)
        print(path4 + ' 创建成功')
    else:
        print(path4 + ' 目录已存在')
        
    isExists5 = os.path.exists(path5)
    if not isExists5:
        os.makedirs(path5)
        print(path5 + ' 创建成功')
    else:
        print(path5 + ' 目录已存在')
        
    isExists6 = os.path.exists(path6)
    if not isExists6:
        os.makedirs(path6)
        print(path6 + ' 创建成功')
    else:
        print(path6 + ' 目录已存在')

def get_order(default_path):
    # 1.打开文件
    # filename = open_file()
    filename = default_path+"/order.xlsx"
    df = pd.read_excel(filename, dtype=str)
    print("读取order文件成功，打印前5行数据：")
    print(df.head(5).to_markdown())
    # 2.选择文件类型
    # type = win32api.MessageBox(0, "请确认你的文件类型,xlsx(xls)选是，csv选否！", "提醒", win32con.MB_YESNO)
    # if type == 6:
    #     df = pd.read_excel(filename, dtype=str)
    #     print("读取文件成功，打印前5行数据：")
    #     print(df.head(5).to_markdown())
    # else:
    #     df = pd.read_csv(filename, dtype=str, encoding="gbk")
    #     print("读取文件成功，打印前5行数据：")
    #     print(df.head(5).to_markdown())

    # 一、格式清理
    # 1.删除表头所有空格、换行符
    for column_name in df.columns:
        df.rename(columns={column_name:column_name.replace(" ","").replace("\n","").strip()},inplace=True)

    # 2.所有字母转为大写
    df = df.apply(lambda x: x.astype(str).str.upper())

    # 3.所有英文圆括号转为中文圆括号
    df = df.apply(lambda x: x.astype(str).str.replace("(", "）").str.replace(")","）"))  # 删除表头空格、换行符

    # 4、日期整理("****-2-29"替换为"****-2-28"，"."替换为"-"，格式转为日期)
    # "订单日期（原料，oem加工，购成品）"值为空，"送货日期（所有补发货单号的都要补送货日期）"不为空，则"订单日期（原料，oem加工，购成品）"值等于"送货日期（所有补发货单号的都要补送货日期）"的值
    # "送货日期（所有补发货单号的都要补送货日期）"值为空，"订单日期（原料，oem加工，购成品）"不为空，则"送货日期（所有补发货单号的都要补送货日期）"值等于"订单日期（原料，oem加工，购成品）"的值
    df["订单日期（原料，oem加工，购成品）"] = df["订单日期（原料，oem加工，购成品）"].apply(lambda x: x.replace(".","-").replace("-2-29","-2-28"))
    df["送货日期（所有补发货单号的都要补送货日期）"] = df["送货日期（所有补发货单号的都要补送货日期）"].apply(lambda x: x.replace(".", "-").replace("-2-29","-2-28").replace("款清发货","2020-01-01"))

    # pandas使用lambda判断元素是否为空或者None
    # f2a_tp2 = df2[df2['combineIdentifyCode'].map(lambda x: len(str(x).strip()) > 0)].copy()  # 识别出合单的订单
    df['订单日期（原料，oem加工，购成品）'] = pd.to_datetime(df['订单日期（原料，oem加工，购成品）'], format='%Y-%m-%d %H:%M:%S')
    df['送货日期（所有补发货单号的都要补送货日期）'] = pd.to_datetime(df['送货日期（所有补发货单号的都要补送货日期）'], format='%Y-%m-%d %H:%M:%S')
    df['订单日期（原料，oem加工，购成品）'] = pd.to_datetime(df["订单日期（原料，oem加工，购成品）"].dt.strftime('%Y-%m-%d'))
    df['送货日期（所有补发货单号的都要补送货日期）'] = pd.to_datetime(df["送货日期（所有补发货单号的都要补送货日期）"].dt.strftime('%Y-%m-%d'))
    # df["订单日期（原料，oem加工，购成品）"] = df["订单日期（原料，oem加工，购成品）"].astype("datetime")
    # df["送货日期（所有补发货单号的都要补送货日期）"] = df["送货日期（所有补发货单号的都要补送货日期）"].astype("datetime")
    df["订单日期（原料，oem加工，购成品）"] = df.apply(lambda x: x["送货日期（所有补发货单号的都要补送货日期）"] if pd.isnull(x["订单日期（原料，oem加工，购成品）"]) else x["订单日期（原料，oem加工，购成品）"], axis=1)
    df["送货日期（所有补发货单号的都要补送货日期）"] = df.apply(lambda x: x["订单日期（原料，oem加工，购成品）"] if pd.isnull(x["送货日期（所有补发货单号的都要补送货日期）"]) else x["送货日期（所有补发货单号的都要补送货日期）"], axis=1)

    # 5.删除表格所有空格、换行符
    # df = df.apply(lambda x: x.astype(str).str.replace(" ","").str.replace(",","").str.replace("\n","").str.strip())
    # df = df.apply(lambda x: x.astype(str).str.replace("NAN","").str.replace("nan","").str.replace("Nan","").str.strip())

    # 5、品牌整理(附文件)
    df["品牌"] = df["品牌"].apply(lambda x: x.replace("ALL NATURAL ADVICE", "肌先知").replace("ALLNATURALADVICE", "肌先知").replace("ANA","肌先知"))
    df["品牌"] = df["品牌"].apply(lambda x: x.replace("BODYAID/博滴", "博滴").replace("BOYAID/博滴", "博滴").replace("BODYAID", "博滴").replace("BOYAID", "博滴"))
    df["品牌"] = df["品牌"].apply(lambda x: x.replace("SAKURAUMATE", "樱语").replace("SAKURAUNATE", "樱语").replace("樱加美", "樱语"))
    df["品牌"] = df["品牌"].apply(lambda x: x.replace("MONTOOTH萌洁齿", "萌洁齿").replace("MONTOOTH", "萌洁齿").replace("萌齿洁", "萌洁齿").replace("VIABLOOM", "植之璨").replace("VITABLOOM", "植之璨"))
    df["品牌"] = df["品牌"].apply(lambda x: x.replace("CAIN", "可瘾").replace("EC", "睐思雅").replace("MAGIC SYMBOL", "魔法符号").replace("MOREI", "多睿").replace("NOCHERN", "若蘅").replace("NUTRIDEA", "盈养泉"))

    # 6、产品整理(附文件)
    df["原零件编号"] = ""
    df["原零件编号"] = df["零件编号（商品条码）"]
    df.loc[df.物品名称=="60ML博滴免冼抑菌冼手液（空白管）","零件编号（商品条码）"]="6923537320059X"
    df.loc[df.物品名称=="100ML博滴免冼抑菌冼手液（印刷管）","零件编号（商品条码）"]="6923537320172XX"
    df.loc[df.物品名称=="100ML博滴免冼抑菌冼手液（空白管）","零件编号（商品条码）"]="6923537320172X"
    df.loc[df.物品名称=="120ML博滴酒精消毒喷雾（有标）","零件编号（商品条码）"]="6925083610394X"
    df.loc[df.物品名称=="120ML博滴酒精消毒喷雾（无标）","零件编号（商品条码）"]="6925083610394XX"
    df.loc[df.物品名称=="500ML免洗抑菌凝胶（标贴款）","零件编号（商品条码）"]="6940843920209X"
    df.loc[df.物品名称=="植之璨30G致嫩雪肌美白清透防晒霜","零件编号（商品条码）"]="6941277011570"
    df.loc[df.物品名称=="450ML博滴75%酒精消毒喷雾（铁罐）（繁体）","零件编号（商品条码）"]="6954299320063X"
    df.loc[df.物品名称=="9ML黄金口腔喷雾（无盒）","零件编号（商品条码）"]="6973007000035X"
    df.loc[df.物品名称=="12ML口腔喷雾（香草冰淇淋）","零件编号（商品条码）"]="6973007000097"
    df.loc[df.物品名称=="50MLMONTOOTH萌齿洁口腔慕斯（朗姆樱桃）","零件编号（商品条码）"]="6973117510028"
    df.loc[df["零件编号（商品条码）"]=="6923537320172", "物品名称"] = "100ML博滴免洗抑菌洗手液"
    df.loc[df["零件编号（商品条码）"]=="6957712933499", "物品名称"] = "博滴琴叶防脱发强根保养洗发水300ML"
    df.loc[df["零件编号（商品条码）"]=="6970464040116", "物品名称"] = "樱加美去污净鞋巾"
    df.loc[df["零件编号（商品条码）"]=="6973003460055", "物品名称"] = "MAGICSYMBOL魔法符号香水免洗干发喷雾（黑暗公爵香型）150ML"
    df.loc[df["零件编号（商品条码）"]=="6973003460062", "物品名称"] = "MAGICSYMBOL魔法符号香水免洗干发喷雾（玫瑰丝带香型）150ML"
    df.loc[df["零件编号（商品条码）"]=="6973007000035", "物品名称"] = "9ML黄金口腔喷雾"
    df.loc[df["零件编号（商品条码）"]=="AL0072-07", "物品名称"] = "肌先知幂爱水润倍护防晒霜50ML-盒子"
    df.loc[df["零件编号（商品条码）"]=="Al0471-06", "物品名称"] = "27ML*5片肌先知艾莉绮美白面膜-内卡"
    df.loc[df["零件编号（商品条码）"]=="Al0471-07", "物品名称"] = "27ML*5片肌先知艾莉绮美白面膜-彩盒"
    df.loc[df["零件编号（商品条码）"]=="AL1259-07", "物品名称"] = "肌先知幂爱水润倍护防晒霜50ML-盒子"

    # 7、"欧诺洁个人护理用品有限公司"替换为"广州欧诺洁个人护理用品有限公司"、"深圳市樱岚护肤品有限公司"替换为"深圳樱岚护肤品有限公司"
    df["委外加工厂"] = df["委外加工厂"].apply(lambda x: x.replace("广州欧诺洁个人护理用品有限公司", "欧诺洁个人护理用品有限公司").replace("欧诺洁个人护理用品有限公司", "广州欧诺洁个人护理用品有限公司"))

    # 8、" 数量 "值为空的填入0
    df.数量.fillna("0",inplace=True)
    df["数量"] = df["数量"].apply(lambda x: x.replace(" ", "0").replace("nan", "0").replace("NAN", "0").replace("Nan", "0"))
    df.loc[~df['数量'].str.isnumeric(), "数量"] = "0"
    # df["数量"] = df["数量"].astype(int)

    # 9、" 含税单价 "值为空的填入0，为"-"的替换为0
    df.含税单价.fillna("0",inplace=True)
    df.含税金额.fillna("0", inplace=True)
    df["含税单价"] = df["含税单价"].apply(lambda x: str(x).replace(" ", "0.00").replace("nan", "0.00").replace("NAN", "0.00").replace("Nan", "0.00").replace("-", "0.00"))
    df["含税金额"] = df["含税金额"].apply(lambda x: str(x).replace(" ", "0.00").replace("nan", "0.00").replace("NAN", "0.00").replace("Nan", "0.00"))
    df["税率"] = df["税率"].apply(lambda x: str(x).replace(" ", "0").replace("nan", "0").replace("NAN", "0").replace("Nan", "0").replace("-", "0"))
    # df.loc[~df['含税单价'].str.isnumeric(), "含税单价"] = "0.00"
    # df.loc[~df['含税金额'].str.isnumeric(), "含税金额"] = "0.00"
    # df.loc[~df['税率'].str.isnumeric(), "税率"] = "0"

    # 10、"单 位"值为"公斤"替换为"kg"
    df["单位"] = df["单位"].apply(lambda x: x.replace("公斤", "kg"))

    # 生成uuid，并填充到空白
    df["采购订单号（原料，包材）"].replace({"NAN":uuid.uuid4(),"nan":uuid.uuid4(),"Nan":uuid.uuid4()},inplace=True)
    df["委外订单号（OEM加工，购成品）"].replace({"NAN":uuid.uuid4(),"nan":uuid.uuid4(),"Nan":uuid.uuid4()},inplace=True)
    df["发货单号（oem加工补，购成品，原料）"].replace({"NAN":uuid.uuid4(),"nan":uuid.uuid4(),"Nan":uuid.uuid4()},inplace=True)
    # df["采购订单号（原料，包材）"] = df["采购订单号（原料，包材）"].apply(lambda x: x.astype(str).str.replace("NAN",uuid.uuid4()).str.replace("nan",uuid.uuid4()).str.replace("Nan",uuid.uuid4()).str.strip())
    # purchase_order = purchase_order.apply(lambda x: x.astype(str).str.replace("NAN","").str.replace("nan","").str.replace("Nan","").str.strip())

    # 修改nan为空
    df = df.apply(lambda x: x.astype(str).str.replace(" ","").str.replace(",","").str.replace("\n","").str.strip())
    df = df.apply(lambda x: x.astype(str).str.replace("NAN","").str.replace("nan","").str.replace("Nan","").str.strip())

    # 11、表格增加序号
    print("正在输出文件，请等待：")
    df = df.reset_index()
    df.to_excel(default_path+"/订单拆分合并/order_clear.xlsx",index=False)
    print("ok！")
    return df


def get_bom(default_path):
    # 1.打开文件
    # filename = open_file()
    filename = default_path + "/bom.xlsx"
    df = pd.read_excel(filename, dtype=str)
    print("读取bom文件成功，打印前5行数据：")
    print(df.head(5).to_markdown())
    # # 2.选择文件类型
    # type = win32api.MessageBox(0, "请确认你的文件类型,xlsx(xls)选是，csv选否！", "提醒", win32con.MB_YESNO)
    # if type == 6:
    #     df = pd.read_excel(filename, dtype=str)
    #     print("读取文件成功，打印前5行数据：")
    #     print(df.head(5).to_markdown())
    # else:
    #     df = pd.read_csv(filename, dtype=str, encoding="gbk")
    #     print("读取文件成功，打印前5行数据：")
    #     print(df.head(5).to_markdown())

    # 一、格式清理
    # 1.删除表头、表格所有空格、换行符
    for column_name in df.columns:
        df.rename(columns={column_name:column_name.replace(" ","").replace("\n","").strip()},inplace=True)
    df = df.apply(lambda x: x.astype(str).str.replace(" ","").str.replace(",","").str.replace("\n","").str.strip())
    # df = df.apply(lambda x: x.astype(str).str.replace("NAN","").str.replace("nan","").str.replace("Nan","").str.strip())

    # 2.所有字母转为大写
    df = df.apply(lambda x: x.astype(str).str.upper())

    # 3.所有英文圆括号转为中文圆括号
    df = df.apply(lambda x: x.astype(str).str.replace("(", "）").str.replace(")","）"))  # 删除表头空格、换行符

    # 4.用量为空的填1
    df.用量.fillna("1")
    df["用量"] = df["用量"].apply(lambda x: x.replace(" ", "1").replace("nan", "1"))

    # 5.品牌整理
    df["品牌名称"] = df["品牌名称"].apply(lambda x:  "博滴" if x.find("BODYAID")>=0 else x)
    df["品牌名称"] = df["品牌名称"].apply(lambda x: "肌先知" if x.find("NATURAL") >= 0 else x)
    df["品牌名称"] = df["品牌名称"].apply(lambda x: "肌先知" if x.find("ANA") >= 0 else x)
    df["品牌名称"] = df["品牌名称"].apply(lambda x: "樱语" if x.find("SAKURA") >= 0 else x)
    df["品牌名称"] = df["品牌名称"].apply(lambda x: "樱语" if x.find("樱加美") >= 0 else x)
    df["品牌名称"] = df["品牌名称"].apply(lambda x: "萌洁齿" if x.find("MONTOOTH") >= 0 else x)
    df["品牌名称"] = df["品牌名称"].apply(lambda x: "可瘾" if x.find("CAIN") >= 0 else x)
    df["品牌名称"] = df["品牌名称"].apply(lambda x: "睐思雅" if x.find("EC") >= 0 else x)
    df["品牌名称"] = df["品牌名称"].apply(lambda x: "魔法符号" if x.find("MAGIC") >= 0 else x)
    df["品牌名称"] = df["品牌名称"].apply(lambda x: "多睿" if x.find("MOREI") >= 0 else x)
    df["品牌名称"] = df["品牌名称"].apply(lambda x: "植之璨" if x.find("VITA") >= 0 else x)
    df["品牌名称"] = df["品牌名称"].apply(lambda x: "若蘅" if x.find("NOCHERN") >= 0 else x)
    df["品牌名称"] = df["品牌名称"].apply(lambda x: "盈养泉" if x.find("NUTRIDEA") >= 0 else x)
    df["品牌名称"] = df["品牌名称"].apply(lambda x: "多睿" if x.find("MOREI") >= 0 else x)
    df["品牌名称"] = df["品牌名称"].apply(lambda x: "美可森" if x.find("MILCOSEN") >= 0 else x)

    # 6.增加品牌简称
    df["品牌简称"] = ""
    df.loc[df["品牌名称"] == "博滴","品牌简称"] = "BO"
    df.loc[df["品牌名称"] == "肌先知","品牌简称"] = "AL"
    df.loc[df["品牌名称"] == "樱语","品牌简称"] = "SA"
    df.loc[df["品牌名称"] == "萌洁齿","品牌简称"] = "MO"
    df.loc[df["品牌名称"] == "睐思雅","品牌简称"] = "EC"
    df.loc[df["品牌名称"] == "魔法符号","品牌简称"] = "MA"
    df.loc[df["品牌名称"] == "多睿","品牌简称"] = "MOR"
    df.loc[df["品牌名称"] == "植之璨","品牌简称"] = "VI"
    df.loc[df["品牌名称"] == "若蘅","品牌简称"] = "NO"
    df.loc[df["品牌名称"] == "盈养泉","品牌简称"] = "NU"
    df.loc[df["品牌名称"] == "美可森","品牌简称"] = "MI"
    df.loc[df["品牌名称"] == "来一泡","品牌简称"] = "DR"
    df.loc[df["品牌名称"] == "可瘾","品牌简称"] = "CA"

    # 7.增加BOM编号
    df1 = df[["条形码编号","品牌简称"]].copy()
    df1["BOM编号"] = ""
    df1 = df1.drop_duplicates(["条形码编号"],keep="first")
    df1 = df1.reset_index()
    for index,row in df1.iterrows():
        df1.loc[index,"BOM编号"]="BOM/" + row["品牌简称"] +"{:0>5d}".format(index+1)
    df = pd.merge(df,df1[["条形码编号","BOM编号"]],how="left",on="条形码编号")

    # 8.表格增加序号
    df = df.reset_index()
    df.to_excel(default_path+"/订单拆分合并/bom_clear.xlsx",index=False)
    print("ok！")
    return df

def data_mark(default_path):
    # 二、数据标记（仅order表）
    df = get_order(default_path)
    df["数据标记"] = ""
    # 1、"类别（OEM加工厂/包材）"值为"购成品"，标记为"委外采购"
    df.loc[df["类别（OEM加工厂/包材）"] == "购成品", "数据标记"] = "委外采购"
    # 2、"类别（OEM加工厂/包材）"值为"原料"，标记为"原料采购"
    df.loc[df["类别（OEM加工厂/包材）"] == "原料", "数据标记"] = "原料采购"
    # 3、"类别（OEM加工厂/包材）"值为"包材"，"公司主体（付款主体）"值不为空且不是"深圳市艾法商贸有限公司"，标记为"包材采购"
    df.loc[df["类别（OEM加工厂/包材）"].str.contains("包材") & df["公司主体（付款主体）"].notnull() & ~df["公司主体（付款主体）"].str.contains("深圳市艾法商贸有限公司"), "数据标记"] = "包材采购"
    # 4、"类别（OEM加工厂/包材）"值为"包材"，"公司主体（付款主体）"值为空，标记为"艾法自制"
    df.loc[df["类别（OEM加工厂/包材）"].str.contains("包材") & df["公司主体（付款主体）"].isnull(), "数据标记"] = "艾法自制"
    # 5、"类别（OEM加工厂/包材）"值为"包材"，"公司主体（付款主体）"值是"深圳市艾法商贸有限公司"，标记为"艾法委外采购"
    df.loc[df["类别（OEM加工厂/包材）"].str.contains("包材") & df["公司主体（付款主体）"].str.contains("深圳市艾法商贸有限公司"), "数据标记"] = "艾法委外采购"
    # 6、"类别（OEM加工厂/包材）"值为"OEM加工厂"，"零件编号（商品条码）"值不是数字条码,且不是X结尾的，标记为"零件采购"
    # df.loc[df["类别（OEM加工厂/包材）"].str.contains("OEM加工厂") & df['原零件编号'].apply(lambda x: x if re.search("^\d+$", str(x)) else np.nan), "数据标记"] = "零件采购"
    df.loc[df["类别（OEM加工厂/包材）"].str.contains("OEM加工厂") & ~df['原零件编号'].str.isnumeric(), "数据标记"] = "零件采购"
    # 7、"类别（OEM加工厂/包材）"值为"OEM加工厂"，"零件编号（商品条码）"值是数字条码或是X结尾的，标记为"委外加工"
    df.loc[df["类别（OEM加工厂/包材）"].str.contains("OEM加工厂") & df['原零件编号'].str.isnumeric(), "数据标记"] = "委外加工"
    df.to_excel(default_path + "/订单拆分合并/order_data_mark.xlsx")
    del df['原零件编号']
    print("ok！")
    return df


# 三、提取数据
def extract_data(default_path):
    # order表
    df = data_mark(default_path)
    # bom表
    df1 = get_bom(default_path)
    # 1、提取业务伙伴信息
    partner_order = df["供应商名称"].copy()
    partner_order = partner_order.reset_index()
    partner_order.to_excel(default_path + "/订单拆分合并/业务伙伴信息order.xlsx", index=False)

    # 2、提取产品信息
    product_order = df[["零件编号（商品条码）", "品牌", "物品名称","单位"]].copy()
    product_order.drop_duplicates("零件编号（商品条码）",inplace=True)
    product_order = pd.DataFrame(product_order)
    product_order["零件编号（商品条码）"].replace("",np.nan,inplace=True)
    product_order["单位"].replace("",np.nan,inplace=True)
    product_order.dropna(axis=0,subset=["零件编号（商品条码）","单位"],inplace=True)
    product_order = product_order.reset_index()
    product_order.to_excel(default_path + "/订单拆分合并/产品信息order.xlsx", index=False)

    # 3、数量<=0的行提取到"退货"表
    return_goods = df[df["数量"].astype(float) < 0].copy()

    return_goods.to_excel(default_path + "/订单拆分合并/退货.xlsx",index=False)

    # 4、数据标记为：1、2、3、4、5、6，提取到"采购订单"表![](../../AppData/Local/Temp/640.webp)
    purchase_order = df[~df["数据标记"].str.contains("委外加工")].copy()
    purchase_order = purchase_order[purchase_order["数量"].astype(float) >= 0].copy()
    purchase_order["备注"] = ""
    purchase_order = purchase_order[
        ["index","采购订单号（原料，包材）", "供应商名称", "公司主体（付款主体）", "订单日期（原料，oem加工，购成品）", "送货日期（所有补发货单号的都要补送货日期）", "采购", "零件编号（商品条码）", "单位",
         "数量", "含税单价", "税率", "备注","委外加工厂","规格(ml/g)","装箱","件数","含税金额","类别（OEM加工厂/包材）","品牌","发货单号（oem加工补，购成品，原料）","物品名称","委外订单号（OEM加工，购成品）"]]
    # purchase_order.to_excel(default_path + "/订单拆分合并/采购订单.xlsx",index=False)
    purchase_order = pd.DataFrame(purchase_order)
    purchase_order["数量"] = purchase_order["数量"].astype(int)

    # purchase_order["采购订单号（原料，包材）"].dropna(inplace=True)
    # for index,row in purchase_order.iterrows():
    #     # print(row["index"],row["采购订单号（原料，包材）"],row["含税单价"])
    #     purchase_order.loc[index,"含税单价"] =  float(row["含税单价"])
    #     purchase_order.loc[index,"含税金额"] =  float(row["含税金额"])

    purchase_order["含税单价"] = purchase_order["含税单价"].astype(float)
    purchase_order["含税金额"] = purchase_order["含税金额"].astype(float)
    purchase_order["税率"] = purchase_order["税率"].astype(float)
    # purchase_order["采购订单号（原料，包材）"] = purchase_order["采购订单号（原料，包材）"].apply(
    #     lambda x: x.astype(str).str.replace("NAN", uuid).str.replace("nan", uuid).str.replace("Nan", uuid).str.strip())

    writer = pd.ExcelWriter(default_path + "/订单拆分合并/采购订单.xlsx", engine='xlsxwriter')
    purchase_order.to_excel(writer, sheet_name='Sheet1',index=False)
    bookObj = writer.book
    writerObj = writer.book.sheetnames['Sheet1']
    format1 = bookObj.add_format({'num_format': '0.00'})
    format2 = bookObj.add_format({'num_format': '0%'})
    writerObj.set_column('K:K', cell_format=format1)
    writerObj.set_column('L:L', cell_format=format2)
    writer.save()


    # 5、数据标记为：7，提取到"委外订单"表
    outsourc = df[df["数据标记"].str.contains("委外加工")].copy()
    outsourc = outsourc[outsourc["数量"].astype(float) >= 0].copy()
    outsourc["备注"] = ""
    outsourc = outsourc[
        ["index", "公司主体（付款主体）", "订单日期（原料，oem加工，购成品）", "采购", "零件编号（商品条码）","单位", "数量", "委外订单号（OEM加工，购成品）",
         "采购订单号（原料，包材）", "备注", "委外加工厂","含税单价","税率", "送货日期（所有补发货单号的都要补送货日期）", "供应商名称",
         "类别（OEM加工厂/包材）", "品牌",  "发货单号（oem加工补，购成品，原料）", "物品名称","规格(ml/g)", "装箱", "件数", "含税金额"]].copy()
    df1.rename(columns={"条形码编号": "零件编号（商品条码）"}, inplace=True)
    df2 = df1.drop_duplicates(["零件编号（商品条码）"], keep="first")
    outsourc_order = pd.merge(outsourc, df2[["BOM编号", "零件编号（商品条码）"]], how="left", on="零件编号（商品条码）")
    # outsourc_order = outsourc_order[
    #     ["index", "公司主体（付款主体）", "订单日期（原料，oem加工，购成品）", "采购", "零件编号（商品条码）","单位", "数量", "委外订单号（OEM加工，购成品）",
    #      "采购订单号（原料，包材）", "备注", "委外加工厂","含税单价","税率", "送货日期（所有补发货单号的都要补送货日期）", "供应商名称",
    #      "类别（OEM加工厂/包材）", "品牌",  "发货单号（oem加工补，购成品，原料）", "物品名称","规格(ml/g)", "装箱", "件数", "含税金额"]].copy()
    outsourc_order = pd.DataFrame(outsourc_order)
    outsourc_order["数量"] = outsourc_order["数量"].astype(int)
    outsourc_order["含税单价"] = outsourc_order["含税单价"].astype(float)
    outsourc_order["含税金额"] = outsourc_order["含税金额"].astype(float)
    outsourc_order["税率"] = outsourc_order["税率"].astype(float)

    # outsourc_order.to_excel(default_path + "/订单拆分合并/委外订单.xlsx", index=False)
    writer = pd.ExcelWriter(default_path + "/订单拆分合并/委外订单.xlsx", engine='xlsxwriter')
    outsourc_order.to_excel(writer, sheet_name='Sheet1', index=False)
    bookObj = writer.book
    writerObj = writer.book.sheetnames['Sheet1']
    format1 = bookObj.add_format({'num_format': '0.00'})
    format2 = bookObj.add_format({'num_format': '0%'})
    writerObj.set_column('L:L', cell_format=format1)
    writerObj.set_column('M:M', cell_format=format2)
    writer.save()

    # bom表
    # 1、提取产品信息
    df1.rename(columns={"零件编号（商品条码）": "条形码编号"}, inplace=True)
    product_bom = df1[["条形码编号", "產品名称", "品牌名称"]].copy()
    product_bom = product_bom.reset_index()
    product_bom.to_excel(default_path + "/订单拆分合并/产品信息bom.xlsx", index=False)

    # 2、提取业务伙伴信息
    partner_bom = df1["供应商"].copy()
    partner_bom = partner_bom.reset_index()
    partner_bom.to_excel(default_path + "/订单拆分合并/业务伙伴信息bom.xlsx", index=False)

    # 3、提取BOM信息
    bom = df1.copy()
    bom["公司"] = ""
    bom["数量"] = ""
    bom = bom[["BOM编号", "公司", "条形码编号", "数量", "單位", "零件編號（更新）", "用量", "單位"]].copy()
    bom.to_excel(default_path + "/订单拆分合并/bom表.xlsx")

    # order、bom提取的产品信息数据合并去重
    product_bom.rename(columns={"条形码编号": "零件编号（商品条码）", "產品名称": "物品名称", "品牌名称": "品牌"}, inplace=True)
    del product_order["index"]
    del product_bom["index"]
    dfs1 = [product_order, product_bom]
    product = pd.concat(dfs1)
    product = product.drop_duplicates(["零件编号（商品条码）"], keep="first")
    product = product.reset_index()
    del product["index"]
    product.to_excel(default_path + "/订单拆分合并/产品信息.xlsx")

    # order、bom提取的业务伙伴信息数据合并去重
    partner_bom.rename(columns={"供应商": "供应商名称"}, inplace=True)
    del partner_order["index"]
    del partner_bom["index"]
    dfs2 = [partner_order, partner_bom]
    partner = pd.concat(dfs2)
    partner = partner.drop_duplicates(["供应商名称"], keep="first")
    partner = partner.reset_index()
    del partner["index"]
    partner.to_excel(default_path + "/订单拆分合并/业务伙伴.xlsx")

    return purchase_order,outsourc_order,return_goods,df


# 四、数据校验
def data_check(default_path):
    message = ""
    # 1、order表数据总行数 = 采购订单表数据总行数 + 委外订单表数据总行数 + 退货表数据总行数
    # 2、order表序号在采购订单表/委外订单表/退货表中存在，且对应序号各字段的值一致相等
    # 3、输出数据校验日志，数据校验日志.xlsx 表格式如下：
    purchase_order, outsourc_order, return_goods, order = extract_data(default_path)
    if (order.shape[0]) == (purchase_order.shape[0]) + (outsourc_order.shape[0]) + (return_goods.shape[0]):
        # if len(order) == len(purchase_order) + len(outsourc_order) + len(return_goods):
        message = message + "\r\n" + ("True:order表数据总行数 = 采购订单表数据总行数 + 委外订单表数据总行数 + 退货表数据总行数")
        message = message + "\r\n" + (
            "{} + {} + {}={}".format(purchase_order.shape[0], outsourc_order.shape[0], return_goods.shape[0],
                                     purchase_order.shape[0] + outsourc_order.shape[0] + return_goods.shape[0]))
    else:
        message = message + "\r\n" + ("False:order表数据总行数 = 采购订单表数据总行数 + 委外订单表数据总行数 + 退货表数据总行数")
        message = message + "\r\n" + (
            "{} + {} + {}={}".format(purchase_order.shape[0], outsourc_order.shape[0], return_goods.shape[0],
                                     purchase_order.shape[0] + outsourc_order.shape[0] + return_goods.shape[0]))

        order_sum = purchase_order["index"].append(outsourc_order["index"]).append(return_goods["index"])
        order_sum = pd.DataFrame(order_sum)

        message = message + "\r\n" + (order[~order["index"].isin(order_sum["index"])].to_markdown())

    if purchase_order[purchase_order.index.isin(order.index)].shape[0] == purchase_order.shape[0]:
        message = message + "\r\n" + ("采购订单表序号均在订单表内")
    else:
        message = message + "\r\n" + ("采购订单表序号不在订单表内")

    if outsourc_order[outsourc_order.index.isin(order.index)].shape[0] == outsourc_order.shape[0]:
        message = message + "\r\n" + ("委外订单表序号均在订单表内")
    else:
        message = message + "\r\n" + ("委外订单表序号不在订单表内")

    if return_goods[return_goods.index.isin(order.index)].shape[0] == return_goods.shape[0]:
        message = message + "\r\n" + ("退货订单序号均在订单表内")
    else:
        message = message + "\r\n" + ("退货订单序号不在订单表内")

    check_purchase = get_defferent_beyondcompare(order, purchase_order, ["index"],
                                                 ["采购订单号（原料，包材）", "供应商名称", "公司主体（付款主体）", "订单日期（原料，oem加工，购成品）",
                                                  "送货日期（所有补发货单号的都要补送货日期）", "采购", "零件编号（商品条码）", "单位", "数量", "含税单价",
                                                  "税率"],
                                                 ["数量", "含税单价", "税率"])
    print("校验结果")
    print(check_purchase.head(10).to_markdown())

    check_purchase.to_excel(default_path + "/订单拆分合并/采购订单数据校验结果.xlsx")

    check_outsourc = get_defferent_beyondcompare(order, outsourc_order, ["index"],
                                                 ["委外订单号（OEM加工，购成品）", "公司主体（付款主体）", "采购", "委外加工厂", "订单日期（原料，oem加工，购成品）",
                                                  "送货日期（所有补发货单号的都要补送货日期）", "零件编号（商品条码）",
                                                  "单位", "数量", "含税单价", "税率"],
                                                 ["数量", "含税单价", "税率"])
    print("校验结果")
    print(check_outsourc.head(10).to_markdown())
    check_outsourc.to_excel(default_path + "/订单拆分合并/委外订单数据校验结果.xlsx")

    check_return = get_defferent_beyondcompare(order, return_goods, ["index"],
                                               ["采购", "供应商名称", "类别（OEM加工厂/包材）", "公司主体（付款主体）", "品牌",
                                                "订单日期（原料，oem加工，购成品）", "送货日期（所有补发货单号的都要补送货日期）", "委外订单号（OEM加工，购成品）",
                                                "采购订单号（原料，包材）", "发货单号（oem加工补，购成品，原料）", "零件编号（商品条码）", "物品名称", "委外加工厂",
                                                "规格(ml/g)", "装箱", "件数", "税率", "单位", "数量", "含税单价", "含税金额"
                                                ],
                                               ["数量", "含税单价", "税率","含税金额"])
    print("校验结果")
    print(check_return.head(10).to_markdown())
    check_return.to_excel(default_path + "/订单拆分合并/退货订单数据校验结果.xlsx")

    return message


# 生成两张表格的差异分析报告
def get_defferent_beyondcompare(df_1,df_2,key_columns,value_columns,number_columns):
    # 假设不重复，假设字段顺序完全一致
    value_columns_len=len(value_columns)

    df1 = df_1.copy()
    df2 = df_2.copy()

    # print(df1.head(3).to_markdown())

    df1["value_combine"]=""
    df2["value_combine"] = ""
    for col in value_columns:
        df1["value_combine"]=df1["value_combine"]+"|"+df1[col].astype(str)
        df2["value_combine"]=df2["value_combine"]+"|"+df2[col].astype(str)

    df1["value_combine"]=df1["value_combine"].apply(lambda x:x[1:])
    df2["value_combine"] = df2["value_combine"].apply(lambda x:x[1:])

    df_left = df1.copy()
    df_right = df2.copy()

    # print(df_left.head(3).to_markdown())

    for col in value_columns:
        del df_left[col]
        del df_right[col]

    # print("左表抽样：")
    # print(df_left.head(5).to_markdown())
    #
    # print("右表抽样：")
    # print(df_right.head(5).to_markdown())


    # print("左表:", df_left.shape[0],"右表:", df_right.shape[0])

    # value_column= "".join(df_left.columns[-1:])
    # print("value_column=",value_column)

    # 附加一个索引列
    df_left["uniqueindex"]=df_left.iloc[:,0]
    df_right["uniqueindex"] = df_right.iloc[:,0]



    # 拼接生成索引行
    # print(df_left.columns)
    # 第一列先设置好，最后一列是计算列，不加入索引
    for c in key_columns[1:]:
        print('字段:',c)
        df_left[c]=df_left[c].astype(str)
        df_right[c] = df_right[c].astype(str)
        df_left["uniqueindex"] =  df_left.apply(lambda x:x["uniqueindex"]+ "|" + x[c] ,  axis=1)
        df_right["uniqueindex"] = df_right.apply(lambda x:x["uniqueindex"]+ "|" + x[c] ,axis=1 )

    # print("左表抽样2：")
    # print(df_left.head(5).to_markdown())
    #
    # print("右表抽样2：")
    # print(df_right.head(5).to_markdown())

    # print("抽样")
    # print(df_left[df_left.uniqueindex.str.contains("太平洋财产保险股份公司")].head(30).to_markdown())
    # print(df_left[df_left.uniqueindex.str.contains("太平洋财产保险股份公司")].head(30).to_markdown())


    # df_result=pd.DataFrame(columns=["uniqueindex1","vouchno1","productcode1","quantity1","match","uniqueindex2","vouchno2","productcode2","quantity2"])
    df_result = pd.DataFrame(columns=["uniqueindex1","value_combine_1", "match", "uniqueindex2","value_combine_2"])

    # 左边多
    # print("列名：",["uniqueindex"]+value_columns)
    df_leftmore = df_left[~df_left["uniqueindex"].isin(df_right["uniqueindex"])][["uniqueindex","value_combine"]].copy()
    df_leftmore=pd.DataFrame(df_leftmore)
    df_leftmore["match"]="+-"
    df_leftmore["other"]=""
    df_leftmore["value_combine"+"_2"] = ""
    # print("左边多:",df_leftmore.shape[0])
    temp_df = df_leftmore[["uniqueindex","value_combine","match","other","value_combine"+"_2"]].copy()
    temp_df.columns = ["uniqueindex1","value_combine"+"_1", "match", "uniqueindex2","value_combine"+"_2"]
    # print(temp_df.head(10).to_markdown())
    df_result = df_result.append(temp_df)


    # 右边多
    df_rightmore = df_right[~df_right["uniqueindex"].isin(df_left["uniqueindex"])][["uniqueindex","value_combine"]].copy()
    df_rightmore = pd.DataFrame(df_rightmore)
    df_rightmore["match"]="-+"
    df_rightmore["other"]=""
    df_rightmore["value_combine" + "_1"] = ""
    # print("右边多:",df_rightmore.shape[0])
    temp_df=df_rightmore[["other","value_combine" + "_1","match","uniqueindex","value_combine" ]].copy()
    temp_df.columns = ["uniqueindex1","value_combine"+"_1", "match", "uniqueindex2","value_combine"+"_2"]
    # print(temp_df.head(10).to_markdown())
    df_result = df_result.append(temp_df)

    #两张表联合查询
    df_inner = df_left.merge(df_right,how="inner",on="uniqueindex")
    df_inner["other"] = df_inner["uniqueindex"]
    # print("联合查询")
    # print(df_inner.head(3).to_markdown())

    # 两边相等
    # df_inner_equal=df_inner.loc[  abs(df_inner[value_column+"_x"].astype(float)-df_inner[value_column+"_y"].astype(float))<0.01].copy()
    df_inner_equal=df_inner.loc[ df_inner["value_combine"+"_x"]==df_inner["value_combine"+"_y"]].copy()
    df_inner_equal = pd.DataFrame(df_inner_equal)
    df_inner_equal["match"] = "="
    temp_df = df_inner_equal[["uniqueindex","value_combine"+"_x", "match", "other","value_combine"+"_y"]].copy()
    temp_df.columns = ["uniqueindex1","value_combine"+"_1", "match", "uniqueindex2","value_combine"+"_2"]

    # print("两边相等:",temp_df.shape[0])
    # print(temp_df.head(10).to_markdown())
    df_result = df_result.append(temp_df)

    # 两边不等于
    # df_inner_not_equal = df_inner.loc[
    #     abs(df_inner[value_column+"_x"].astype(float) - df_inner[value_column+"_y"].astype(float)) >= 0.01].copy()

    df_inner_not_equal = df_inner.loc[df_inner["value_combine" + "_x"] != df_inner["value_combine" + "_y"]].copy()

    df_inner_not_equal["match"] = "<>"
    temp_df = df_inner_not_equal[["uniqueindex","value_combine"+"_x", "match", "other","value_combine"+"_y"]].copy()
    temp_df.columns = ["uniqueindex1","value_combine"+"_1", "match", "uniqueindex2","value_combine"+"_2"]

    # print("两边不相等:",temp_df.shape[0])
    # print(df_inner_not_equal[["uniqueindex", "match", "uniqueindex"]].head(100).to_markdown())
    # print(temp_df.head(3).to_markdown())
    df_result = df_result.append(temp_df)

    df_result.fillna("",inplace=True)
    # print("debug1")
    # print(df_result.head(3).to_markdown())

    # 找回原来的字段列表，形成uniqueindex
    df1["uniqueindex"] = df1.iloc[:, 0]
    df2["uniqueindex"] = df2.iloc[:, 0]
    for c in key_columns[0:-2]:
        # print('字段:', c)
        # print(df1.head(2).to_markdown())
        if c!="product":
            df1[c] = df1[c].astype(str)
            df2[c] = df2[c].astype(str)
            df1["uniqueindex"] = df1.apply(lambda x: x["uniqueindex"] + "|" + x[c], axis=1)
            df2["uniqueindex"] = df2.apply(lambda x: x["uniqueindex"] + "|" + x[c], axis=1)

    # print("检查原始表")
    # print(df1.head(5).to_markdown())
    #
    # print("检查结果表")
    # print(df_result.head(5).to_markdown())

    # 把列名解压缩
    # print("列名解压缩")
    # print(df_left.columns)

    df_result["value_combine_1"]=df_result["value_combine_1"].astype(str)
    df_result["value_combine_2"] = df_result["value_combine_2"].astype(str)

    # print("拆解前")
    # print(df_result.head(5).to_markdown())

    # 拆解出关键索引字段
    # index = 0
    # for c in df_left.columns[0:-2]:
    #     # print(c,index)
    #     df_result[c+"_1"] = df_result.apply(
    #         lambda x:x["uniqueindex1"].split("|")[index] if str(x["uniqueindex1"]).split("|").__len__() > index else '',
    #         axis=1)
    #     df_result[c+"_2"] = df_result.apply(
    #             lambda x:x["uniqueindex2"].split("|")[index] if str(x["uniqueindex2"]).split("|").__len__() > index else '',
    #             axis=1)
    #     index = index+1

    # print("拆解出数据字段")
    # print(df_result.head(5).to_markdown())
    index = 0
    for c in value_columns:
        # print(index,c)
        df_result[c+"_1"] = df_result.apply(
                lambda x:x["value_combine_1"].split("|")[index] if  str(x["value_combine_1"]).split(
                    "|").__len__() > index else '',
                axis=1)
        df_result[c+"_2"] = df_result.apply(
                lambda x:x["value_combine_2"].split("|")[index] if  str(x["value_combine_2"]).split(
                    "|").__len__() > index else '',
                axis=1)
        index = index+1

    # df_result=df_result.merge(df1[["uniqueindex","product","qty","amt"]],how="left",left_on="uniqueindex1",right_on="uniqueindex")
    # df_result = df_result.merge(df2[["uniqueindex","product", "qty", "amt"]], how="left", left_on="uniqueindex2",
    #                             right_on="uniqueindex")

    # print("debug2")
    # df_result=df_result.astype(str)
    # print(df_result.head(5).to_markdown())

    # key_columns,value_columns

    str_column1 = ""
    str_column2 = ""
    for c in key_columns:
        str_column1 = str_column1+","+c+"_1"
        str_column2 = str_column2+","+c+"_2"

    # str_column1 = "uniqueindex1,value_combine_1"+str_column1
    # str_column2 = "uniqueindex2,value_combine_2"+str_column2

    for c in value_columns:
        str_column1 = str_column1+","+c+"_1"
        str_column2 = str_column2+","+c+"_2"

    str_column = str_column1+",match"+str_column2
    str_column=str_column[1:].strip() # 去掉开始的逗号


    # print("字段列表：",str_column)
    str_column=str_column.replace("index_1,","uniqueindex1,").replace("index_2,","uniqueindex2,")
    # print(str_column.split(","))
    series_column = pd.Series(str_column.split(","))
    # print(df_result[series_column].head(10).to_markdown())

    # print(set(number_columns))

    for col in df_result.columns:
        if col!="match":
            df_result[col] = df_result[col].apply(lambda x:"'{}".format(x) if len(str(x)) > 0 else '')  # 强制转字符串,避免转数字
            for col2 in number_columns:
                # print("col=col2 ",col,col2)
                if col.replace("_1","").replace("_2","")==col2:
                    # print("col2:",col2)
                    df_result[col] = df_result[col].apply(lambda x:x.replace("'",""))
                    # df_result[col]=df_result[col].astype(float)

    #number_columns
    # df_result[series_column].to_excel(r"/Users/vicetone/lclproject/python/ITDD/data/财务数据/卖家联合/比对结果(左 {},右 {}).xlsx".format(leftname,rightname))
    # return df_result[series_column]


    # print("不相等的有{}行".format(df_result[df_result.match.str.contains("<>")].shape[0]))

    # print(df_result.head(10).to_markdown())

    df_result["match"]=df_result["match"].apply(lambda x:x.replace("=","等于"))
    # df_result.match.str.contains("=")

    return df_result.loc[df_result.match.str.contains("<>") | df_result.match.str.contains("等于"),series_column]


# 选择文件模块
def open_file():
    dlg = win32ui.CreateFileDialog(1)  # 1表示打开文件对话框
    dlg.SetOFNInitialDir('E:/Python')  # 设置打开文件对话框中的初始显示目录
    dlg.DoModal()
    filename = dlg.GetPathName()  # 获取选择的文件名称
    print("filename=",filename)
    print("read ok")
    return filename


if __name__ == "__main__":

    print("请选择所要拆分的文件所在路径：")
    # default_path = input()
    default_path = r"C:\Users\mega\PycharmProjects\Shangwu\拆分"
    # if len(default_path) == 0:
    #     print("你没有输入任何东西，退出！")
    # else:
    #     # merge_data(default_path)
    #     mkdir(default_path)
    #     order = win32api.MessageBox(0, "是否要处理的order文件，否则跳过处理！", "提醒", win32con.MB_YESNO)
    #     if order == 6:
    #         get_order(default_path)
    #     else:
    #         print("跳过order文件处理")
        # bom = win32api.MessageBox(0, "是否要处理的bom文件，否则跳过处理！", "提醒", win32con.MB_YESNO)
        # if bom == 6:
        #     get_bom()
        # else:
        #     print("跳过bom文件处理")
    message=data_check(default_path)
    print(message)


    #
    # order=pd.read_pickle(r"C:\Users\mega\PycharmProjects\Shangwu\data\order.pkl")
    # purchase_order=pd.read_pickle(r"C:\Users\mega\PycharmProjects\Shangwu\data\purchase_order.pkl")
    #
    # print(order.head(50).to_markdown())


