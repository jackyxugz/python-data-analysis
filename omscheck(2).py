# coding=utf-8

import pandas as pd
import numpy as np
from tkinter import filedialog
import os.path
from pathlib import Path
import time
import json
import warnings
import re


# 定义类
class Fasttable:
    # 成员变量，类似全局变量
    rootdir=""

    # 不限定表格的字段名

    # 不限定表格的字段名
    policy = {"京东海外": {"filekey": "海外", "sheets": [{"sheetname": "", "ignoretop": "0", "title_key": "金额", "bottom_key": "合计",
                            "columns": ""}]},
                   "京东": {"filekey": "!海外", "sheets": [{"sheetname": "", "ignoretop": "0", "title_key": "", "bottom_key": "",
                          "columns": ""}]}}

    file_columns_list = []

    def __init__(self):
        # self.rootdir=rootdir
        pass
        # print("开始解压缩...")

    def set_policy(self,policy_file):
        self.policy = policy_file



    def get_policy(self,filename):
        # 根据文件名，自动识别读取政策
        # 将json数据转为字符串，然后转为字典dict
        # str 也可以从 文件直接读取
        str = json.dumps(self.policy)

        # with open(filename, 'r', encoding="gb2312") as input_file:
        #     load_dict = json.load(input_file)
        data = json.loads(str)
        # print(data)

        # 表格政策列表 不同sheet有不同的格式，相当于不同的政策
        policy_list = []
        for jo in data:
            plat = jo
            filekey_list = data[plat]["filekey"]
            # print("开始匹配策略:",plat)
            policy_match_all=True
            for filekey in filekey_list.split(" "): # 支持多个条件筛选
                # 根据不同的文件匹配不同的policy
                policy_match = False
                if filekey.find("!") >= 0:  # 不能包含 关键词  ，反选！
                    if filename.find(filekey.replace("!", "")) < 0:  # 文件名 没有发现不能包含额关键词
                        # print("平台=", plat)
                        # print("filekey=", data[plat]["filekey"])

                        # print("没有发现不含关键词:",filekey)

                        for sheet in data[plat]["sheets"]:
                            # print("sheet=",sheet)
                            # sheetname = data[plat]["sheets"][item]["sheetname"]
                            sheetname = sheet["sheetname"]
                            title = sheet["title"]
                            ignoretop = sheet["ignoretop"]
                            title_key = sheet["title_key"]
                            bottom_key = sheet["bottom_key"]
                            skiptop = sheet["skiptop"]
                            skipbottom = sheet["skipbottom"]

                            _valid_columns = sheet["columns"]
                            _new_columns = sheet["newcolumns"]
                            _rename = sheet["rename"]
                            # 默认就有的字段，给初值
                            _default_columns = sheet["default_columns"]
                            # if len(_default_columns) > 0:
                            #     default_columns = _default_columns.split(",")

                            # print("json 拆解:")
                            # print(sheetname,ignoretop,title_key,bottom_key,_valid_columns,_new_columns)
                            policy_match=True
                            # policy_list.append([sheetname,ignoretop,title_key,bottom_key,_valid_columns,_new_columns,_rename,_default_columns])
                    else:
                        pass
                elif filekey.find(">") >= 0:  # 天猫>淘宝  天猫出现的次数比淘宝多，则 选中
                    k1= filekey.split(">")[0]
                    k2= filekey.split(">")[1]
                    if filename.count(k1) > filename.count(k2):
                        # print("发现:",data[plat]["title_key"],cishu)
                        # print("发现 {} 比 {} 多".format(k1,k2))

                        for sheet in data[plat]["sheets"]:
                            sheetname = sheet["sheetname"]
                            title = sheet["title"]
                            ignoretop = sheet["ignoretop"]
                            title_key = sheet["title_key"]
                            bottom_key = sheet["bottom_key"]
                            skiptop = sheet["skiptop"]
                            skipbottom = sheet["skipbottom"]

                            _valid_columns = sheet["columns"]
                            _new_columns = sheet["newcolumns"]
                            _rename = sheet["rename"]
                            _default_columns = sheet["default_columns"]
                            # if len(_valid_columns) > 0:
                            #     valid_columns = _valid_columns.split(",")
                            policy_match = True
                            # policy_list.append([sheetname, ignoretop, title_key, bottom_key, _valid_columns,_new_columns,_rename,_default_columns])

                    else:
                        pass
                elif filekey.find("<") >= 0:  # 天猫<淘宝  天猫出现的次数比淘宝少，则 选中
                    k1 = filekey.split(">")[0]
                    k2 = filekey.split(">")[1]
                    if filename.count(k1) < filename.count(k2):
                        # print("发现:",data[plat]["title_key"],cishu)
                        # print("发现 {} 比 {} 少".format(k1, k2))

                        for sheet in data[plat]["sheets"]:
                            sheetname = sheet["sheetname"]
                            title = sheet["title"]
                            ignoretop = sheet["ignoretop"]
                            title_key = sheet["title_key"]
                            bottom_key = sheet["bottom_key"]
                            skiptop = sheet["skiptop"]
                            skipbottom = sheet["skipbottom"]

                            _valid_columns = sheet["columns"]
                            _new_columns = sheet["newcolumns"]
                            _rename = sheet["rename"]
                            _default_columns = sheet["default_columns"]
                            # if len(_valid_columns) > 0:
                            #     valid_columns = _valid_columns.split(",")
                            policy_match = True
                            # policy_list.append([sheetname, ignoretop, title_key, bottom_key, _valid_columns,_new_columns,_rename,_default_columns])

                    else:
                        pass
                elif filekey.find("|") >= 0:  #  包含多个项目
                    for son_key in filekey.split("|"):
                        # print("多个项目:",filekey)
                        # print("子项目:",son_key)
                        if filename.find(son_key) > 0:  # 文件名 没有发现不能包含额关键词
                            # print("发现子项:{}".format(son_key))
                            for sheet in data[plat]["sheets"]:
                                # print("sheet=",sheet)
                                # sheetname = data[plat]["sheets"][item]["sheetname"]
                                sheetname = sheet["sheetname"]
                                title = sheet["title"]
                                ignoretop = sheet["ignoretop"]
                                title_key = sheet["title_key"]
                                bottom_key = sheet["bottom_key"]
                                skiptop = sheet["skiptop"]
                                skipbottom = sheet["skipbottom"]

                                _valid_columns = sheet["columns"]
                                _new_columns = sheet["newcolumns"]
                                _rename = sheet["rename"]
                                _default_columns = sheet["default_columns"]
                                # if len(_valid_columns) > 0:
                                #     valid_columns = _valid_columns.split(",")

                                # print("json 拆解:")
                                # print(sheetname,ignoretop,title_key,bottom_key,_valid_columns,_new_columns)
                                policy_match = True
                                # policy_list.append(
                                #     [sheetname, ignoretop, title_key, bottom_key, _valid_columns, _new_columns, _rename,_default_columns])
                        else:
                            pass
                elif filename.find(filekey) >= 0:
                    # print("平台=", plat)
                    # print("filekey=", data[plat]["filekey"])
                    # print("发现匹配关键词:{}  ".format(filekey))

                    for sheet in data[plat]["sheets"]:
                        print(sheet)
                        sheetname = sheet["sheetname"]
                        title = sheet["title"]
                        ignoretop = sheet["ignoretop"]
                        title_key = sheet["title_key"]
                        bottom_key = sheet["bottom_key"]
                        skiptop = sheet["skiptop"]
                        skipbottom = sheet["skipbottom"]

                        _valid_columns = sheet["columns"]
                        _new_columns = sheet["newcolumns"]
                        _rename = sheet["rename"]
                        _default_columns = sheet["default_columns"]
                        # if len(_valid_columns) > 0:
                        #     valid_columns = _valid_columns.split(",")
                        policy_match = True
                        # policy_list.append([sheetname, ignoretop, title_key, bottom_key, _valid_columns,_new_columns,_rename,_default_columns])

                policy_match_all=policy_match and policy_match_all
                # print("本轮匹配结果:",policy_match_all)
                # print(filekey)


            # 如果所有条件都匹配，通过筛选，则可以使用被策略文件
            if   policy_match_all :
                print("匹配策略成功:",filename)
                # print(plat, sheetname, ignoretop,skiptop,skipbottom, title_key, bottom_key, _valid_columns, _new_columns, _rename, _default_columns)
                policy_list.append(
                    [plat,sheetname, title,ignoretop, skiptop,skipbottom,title_key, bottom_key, _valid_columns, _new_columns, _rename, _default_columns])

            else:
                # print("匹配策略失败:", filekey_list)
                pass

        # return sheetname, ignoretop, title_key, bottom_key,valid_columns
        if  len(policy_list)==0:
            print(filename," 匹配策略失败!")
        else:
            # print(policy_list)
            pass

        return policy_list

    def get_top_row_line(self, temp_df, i):
        # 整行内容合并成一个字符串
        row=""
        if i<temp_df.shape[0]:
            for j in range(0, len(temp_df.columns)):
                # print(df.iloc[i, j] )
                row = row + "|" + str(temp_df.iloc[ i, j])

        return row

    def get_bottom_row_line(self,temp_df,i):
        # 整行内容合并成一个字符串
        row = ""
        for j in range(0, len(temp_df.columns)):
            # print(df.iloc[i, j] )
            # print(temp_df.iloc[temp_df.shape[0] - i , j])
            # print(temp_df.shape[0] , i)
            if temp_df.shape[0] > i:
                if i>0:
                    row = row + "|" + str(temp_df.iloc[temp_df.shape[0] - i , j])

        return row

    def get_row_line(self, temp_df, row_number):
        # 整行内容合并成一个字符串
        row = ""
        for j in range(0, len(temp_df.columns)):
            if temp_df.shape[0] > row_number:
                if row_number >= 0:
                    row = row + "|" + str(temp_df.iloc[row_number, j])

        return row

    def get_key_row(self,temp_df,row1, row2,keys):
        max_rownumber=0
        min_rownumber=0
        for i in range(row1, row2):
            row = self.get_row_line(temp_df, i)
            # print("row3:", row)
            # print(keys)
            if len(keys) > 0:
                keys = keys.replace(" ", ",").replace("，", ",").replace(",,", ",").replace(",,", ",")
                for key in keys.split(","):
                    if len(key) > 0:
                        if row.find(key) > 0:
                            # 末尾通常是合计行
                            print("第{}行发现关键词：".format(i), key, " in:  ", row)
                            # print(i)
                            if i > max_rownumber:
                                max_rownumber= i

                            if i < min_rownumber:
                                min_rownumber=i

        return   min_rownumber,max_rownumber

    def find_top(self,temp_df,title_key):
        print("定位字段名关键词:",title_key);
        skiptop = 0
        row = ""
        for i in range(0,  10):
            row = self.get_top_row_line(temp_df, i)
            # for j in range(0, len(temp_df.columns) - 1):
            #     # print(df.iloc[i, j] )
            #     row = row + "|" + str(temp_df.iloc[i, j])
            # if row.find("金额") >= 0:

            # if row.find("店铺名称")>=0:
            #     shopname= row.replace("店铺名称：","").replace("|","")

            # 用关键词定位标题行
            # if ((row.find("项目") >= 0) or (row.find("金额") >= 0)):
            # if  row.find(title_key) >= 0  :
            #     skiptop = i + 1

            break_flag = False
            # print("title_key:",title_key)
            if len(title_key) > 0:
                title_key = title_key.replace(" ", ",").replace("，", ",").replace(",,", ",").replace(",,", ",")
                for key in title_key.split(","):
                    if not break_flag:
                        # print("row",i)
                        # print(row)
                        if row.find(key) > 0:
                            # print("发现关键字段名:", key, "：", row)
                            skiptop = i + 1
                            break_flag = True

            # 如果没有发现关键词，再做一遍头部字段名的判断
            if not break_flag:
                if "level_0" in temp_df.columns:
                # if row.find("level_0") > 0:
                    print("发现 level_0")
                    # print(row)
                    skiptop = skiptop+1



            # 返回需要忽略的头部行数
            # print("头部忽略行数:{} ".format(skiptop))
            # if skiptop>1:
            #     print(temp_df.head(skiptop+2).to_markdown())

            # return skiptop

        print("最后一次用# 来匹配")
        if skiptop==0:
            for i in range(1, 8):
                row = self.get_top_row_line(temp_df, i)
                print("row===", row)
                if row.find("#") >= 0:  # 如果打头是 # ，表示本行为注释
                    print("发现了#")
                    skiptop = i-2  # 扣掉本行，还有标题行，共2行
                else:
                    print("本行没有#，头部忽略行数2:{} ".format(skiptop))
                    return skiptop

        print("头部忽略{}行".format(skiptop))
        return skiptop


    def find_bottom(self, temp_df, bottom_key):
        print("寻找末尾行 ",temp_df.shape[0])
        skipbottom = 0
        skipbottom1 = 0
        skipbottom2 = 0
        # print("忽略尾部")
        large_row = ""
        small_row = "xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx"
        print("find bottom:",min(temp_df.shape[0] - 1, 15))
        # 最多倒查15行
        for i in range(0, min(temp_df.shape[0] - 1, 15)):
            # print("倒数第{}行".format(i+1))
            row = self.get_bottom_row_line(temp_df, i+1)
            # for j in range(0, len(temp_df.columns)):
            #     # print(df.iloc[i, j] )
            #     row = row + "|" + str(temp_df.iloc[temp_df.shape[0] - i, j])

            if len(row) > len(large_row):
                large_row = row
            if len(row) < len(small_row):
                small_row = row
            # 如果这一行很短，说明这是合计行
            if len(small_row) <= len(large_row) / 2:
                print("发现末行：", row)
                skipbottom1 = i - 1
                break
            # print("跟踪1 small_row：",small_row)
            # print("跟踪2 large_row：",large_row)
            # print("跟踪3 row：",row)

            # 前两列都是空格，说明这是合计行
            # 如果前三列，有2列为空，则表示这是一个空行
            left_value=""
            l1 = 0
            l2 = 0
            l3 = 0
            l4 = 0
            l5 = 0
            if len(temp_df.columns)>=1:
                value_1 = str(temp_df.iloc[temp_df.shape[0] - i-1, 0])
                if len(value_1) > 0:
                    l1 = 1
                left_value=left_value+"|"+value_1
            if len(temp_df.columns)>=2:
                value_2 = str(temp_df.iloc[temp_df.shape[0] - i-1, 1])
                if len(value_2) > 0:
                    l2 = 1
                left_value = left_value + "|" + value_2
            if len(temp_df.columns)>=3:
                value_3 = str(temp_df.iloc[temp_df.shape[0] - i-1, 2])
                if len(value_3) > 0:
                    l3 = 1
                left_value = left_value + "|" + value_3
            if len(temp_df.columns)>=4:
                value_4 = str(temp_df.iloc[temp_df.shape[0] - i-1, 3])
                if len(value_4) > 0:
                    l4 = 1
                left_value = left_value + "|" + value_4
            if len(temp_df.columns)>=5:
                value_5 = str(temp_df.iloc[temp_df.shape[0] - i-1, 4])
                if len(value_5) > 0:
                    l5 = 1
                left_value = left_value + "|" + value_5

            # 如果发现合计
            if left_value.find("合计")>=0:
                print("发现合计 ",left_value)
                skipbottom1 = i + 1
            elif l1 + l2 + l3+ l4+ l5 <= 4:
            # 如果前三列，有2列为空，则表示这是一个空行
                print("如果前5列，有2列为空，则表示这是一个空行")
                # print(value_1)
                # print(value_2)
                # print(value_3)
                skipbottom1 = i + 1

            # print(row)

        col_1 = "".join(temp_df.columns[0])
        col_2 = "".join(temp_df.columns[1])
        # for i in range(1, min(temp_df.shape[0] - 1, 15)):
        #     row = self.get_bottom_row_line(temp_df, i)
        #     # 整行内容合并成一个字符串
        #     # for j in range(0, len(temp_df.columns)):
        #     #     # print(df.iloc[i, j] )
        #     #     row = row + "|" + str(temp_df.iloc[temp_df.shape[0] - i, j])
        #
        #     # 用关键词定位合计行
        #     # if ((row.find("合计") > 0) | (row.find("总收入") > 0)):
        #
        #     # 优先级汇总
        #     # if  row.find("结算汇总") > 0 :
        #     #     print("发现合计行：", row)
        #     #     skipbottom2 = i
        #     #     break
        #
        #     # if  row.find(bottom_key) > 0 :
        #     #     print("发现合计行：", row)
        #     #     skipbottom2 = i
        #     #     break
        #
        #
        #     max_skipbottom=0
        #     break_flag = False
        #     if len(bottom_key) > 0:
        #         bottom_key = bottom_key.replace(" ", ",").replace("，", ",").replace(",,", ",").replace(",,", ",")
        #         # print("bottom_key:",bottom_key)
        #         for key in bottom_key.split(","):
        #             if not break_flag:
        #                 if len(key) > 0:
        #                     if row.find(key) > 0:
        #                         # 末尾通常是合计行
        #                         # print("发现合计行2：", key, " in:  ", row)
        #                         skipbottom2 = i
        #                         if skipbottom2>max_skipbottom:
        #                             skipbottom2 = max_skipbottom
        #
        #                         break_flag = True
        #
        #     # skipbottom2=max_skipbottom
        #
        #
        #
        #     # blank_temp_df = temp_df[[col_1, col_2]].copy()
        #     # blank_temp_df[col_1] = blank_temp_df[col_1].astype(str)
        #     # blank_temp_df[col_2] = blank_temp_df[col_2].astype(str)
        #     # temp_df = temp_df[
        #     #     (blank_temp_df[col_1].str.strip().str.len() > 0) & (blank_temp_df[col_2].str.strip().str.len() > 0)]

        # 倒数15行，从上往下查

        # skipbottom3=0
        row2=temp_df.shape[0]-1
        row1=row2-min(temp_df.shape[0] - 1, 15)


        if row1>row2:
            row1=row2

        print("倒数15行，从上往下查!", row1, row2,bottom_key)

        max_skipbottom = 0
        row_min,row_max=self.get_key_row( temp_df, row1, row2, bottom_key)
        print("row_min,row_max=",row_min,row_max)
        if row_min>0:
            skipbottom2=temp_df.shape[0]-row_min
        elif row_max>0:
            skipbottom2=temp_df.shape[0]-row_max

        # for i in range(row1, row2):
        #     row=""
        #     row=self.get_row_line(temp_df,i)
        #     # for j in range(0, len(temp_df.columns)):
        #     #     row = row + "|" + str(temp_df.iloc[i, j])
        #
        #     print("row3:",row)
        #     print(bottom_key)
        #     break_flag = False
        #     if len(bottom_key) > 0:
        #         bottom_key = bottom_key.replace(" ", ",").replace("，", ",").replace(",,", ",").replace(",,", ",")
        #         # print("bottom_key:",bottom_key)
        #         for key in bottom_key.split(","):
        #             if not break_flag:
        #                 if len(key) > 0:
        #                     if row.find(key) > 0:
        #                         # 末尾通常是合计行
        #                         print("发现合计行3：", key, " in:  ", row)
        #                         skipbottom3 = i
        #                         print(skipbottom3)
        #                         if skipbottom3>max_skipbottom:
        #                             skipbottom3 = max_skipbottom
        #                             break
        #
        #                         break_flag = True


        print("忽略尾巴:",skipbottom1, skipbottom2)
        skipbottom = max(skipbottom1, skipbottom2)
        print(skipbottom1, skipbottom2,skipbottom)
        # skipbottom = max(skipbottom, skipbottom3)
        # print(skipbottom1, skipbottom2,skipbottom3, skipbottom)

        # print("最后一次用#来匹配")
        # skipbottom3=0
        # for i in range(1, 8):
        #     row = self.get_bottom_row_line(temp_df, i)
        #     row=row.replace(" ","")
        #     print("bottom_row===", row)
        #     if row.find("#") >= 0:  # 如果打头是 # ，表示本行为注释
        #         print(row+" bottom发现了#")
        #         skipbottom3 = i   # 扣掉本行，还有标题行，共2行
        #     else:
        #         print(row )
        #         print("本行没有#，尾部忽略行数:{} ".format(skipbottom3))
        #         break
        #         # return skiptop
        #
        # skipbottom = max(skipbottom, skipbottom3)
        # print("末尾 忽略1:{}，忽略2:{}，忽略3:{}，skipbottom结果={}".format(skipbottom1, skipbottom2,skipbottom3, skipbottom))
        print("末尾 忽略1:{}，忽略2:{}， skipbottom结果={}".format(skipbottom1, skipbottom2, skipbottom))
        return skipbottom

    # def del_bottom_blank(self, temp_df):

    def add_blankcolumns(self,temp_df,new_columns):
        # 仍然缺少的字段，默认也给补上
        if  len(new_columns.strip())>0:
            for col in new_columns.split(","):
                if col in temp_df.columns:
                    pass
                else:
                    temp_df[col] = ""

            print("返回取的字段")
            # new_columns="'平台','店铺名称','开始时间','订单金额'"
            # print(new_columns)
            print(temp_df.head(2).to_markdown())
            newcolumns_series = new_columns.split(",")

            # print("构建字段series")
            # print(newcolumns_series)

            # temp_df2=temp_df[[newcolumns_series]]
            temp_df2 = temp_df[newcolumns_series]
            return temp_df2
        else:
            return temp_df

    def default_column_set(self,temp_df,default_columns):
        # 给一些字段设置默认值
        # print("默认值:")
        # print(default_columns)
        for key_value in default_columns:
            # print(key_value)
            # print(type(key_value))  # dict
            for key in key_value:
                value = key_value[key]
                # print(str(key) + ' 设置默认值为： ' + str(value))
                if value in temp_df.columns:
                    pass
                else:
                    temp_df[key] = value
        return temp_df

    def rename_column(self,temp_df,column_rename):
        # 字段重命名
        for key_value in column_rename:
            # print(key_value)
            # print(type(key_value))  # dict
            for key in key_value:
                value = key_value[key]

                # key=key_value.split(":")[0]
                # value=key_value.split(":")[1]
                # print(str(key) + '=' + str(value))
                if value in temp_df.columns:
                    pass
                else:
                    if key in temp_df.columns:
                        # print(key + ':' + key_value[key])
                        temp_df.rename(columns={key: value}, inplace=True)
        return temp_df

    def strtoint(self,str):
        if str=="":
            return 0
        else:
            return int(str)

    #

    # 智能读取数据表
    def read_worksheet_with_policy(self, filename,sheet_name_def, item):
        # sheet_name, ignoretop, title_key, bottom_key, valid_columns,new_columns,column_rename,default_columns
        plat, sheet_name, title, ignoretop, skiptop, skipbottom, title_key, bottom_key, default_columns, new_columns, column_rename, default_columns = \
            item[0], item[1], item[2], item[3], item[4], item[5], item[6], item[7], item[8], item[9], item[10], item[11]

        sheet_name=sheet_name_def
        print("读取参数配置:", item)

        skiptop = str(self.strtoint(skiptop))
        skipbottom = str(self.strtoint(skipbottom))

        if len(title.strip()) == 0:
            title = sheet_name
            if len(title.strip()) == 0:
                title = plat

        print("过滤策略1：", self.strtoint(skiptop), self.strtoint(skipbottom), self.strtoint(ignoretop))
        if (self.strtoint(skiptop) + self.strtoint(skipbottom) + self.strtoint(ignoretop) == 0):
            # 无需加工，直接读取
            print(filename, " read no filter")
            with warnings.catch_warnings(record=True):
                warnings.simplefilter("always")
                if os.path.splitext(filename)[1].find("xls") >= 0:
                    if len(sheet_name) > 0:
                        data_xls = pd.read_excel(filename, sheet_name=None, dtype=str)
                        for key in data_xls:
                            if key == sheet_name:
                                # print(" sheet_name 策略匹配成功:",sheet_name)
                                print(" {}@{} 匹配 {} 策略成功(code:1) ".format(sheet_name, filename, title))
                                temp_df = pd.read_excel(filename, sheet_name=sheet_name, dtype=str)
                    else:
                        print(" sheet_name为空 策略匹配成功(code:2) ")
                        temp_df = pd.read_excel(filename, dtype=str)
                else:
                    # if os.path.splitext(filename)[1].find("xls") < 0:
                    if os.path.splitext(filename)[1].find("csv") >= 0:
                        if len(sheet_name) == 0:  # csv 没有 sheetname
                            print(" {} 匹配 {} 策略成功(code:3) ".format(filename, title))
                            try:
                                temp_df = pd.read_csv(filename, encoding="gbk", dtype=str,
                                                      on_bad_lines='skip').reset_index()  # ,decodeing="utf-8"
                                # temp_df = pd.read_csv(filename,encoding="gb18030", dtype=str,error_bad_lines=False).reset_index()  # ,decodeing="utf-8"
                            except Exception as  err:
                                # print(filename, " 异常:", err)
                                # print(filename, "是空表")
                                try:
                                    temp_df = pd.read_csv(filename, encoding="gbk", dtype=str,
                                                          error_bad_lines=False, engine="python").reset_index()
                                except Exception as  err:
                                    print(filename, " 异常:", err)
                                    # print(filename, "是空表")

                                    return pd.DataFrame(columns=["iid"]).head(0)
                        else:
                            return pd.DataFrame(columns=["iid"]).head(0)
            # 字段重命名
            temp_df = self.rename_column(temp_df, column_rename)
            # 给一些字段设置默认值
            temp_df = self.default_column_set(temp_df, default_columns)
            # 补充空字段，并且按照新的字段标准返回数据
            temp_df = self.add_blankcolumns(temp_df, new_columns)
            return temp_df

        elif ((self.strtoint(skiptop) > 0) and (self.strtoint(skipbottom) > 0) and (self.strtoint(ignoretop) > 0)):
            # 直接按照约定过滤标准执行 要求过滤无效行，并且指定忽略的头部行以及尾部行
            print(filename, " read with fiter ", skipbottom)
            if os.path.splitext(filename)[1].find("xls") >= 0:
                # 读取excel
                if len(sheet_name) == 0:  # 无需指定 sheetname
                    print(" {} 匹配 {} 策略成功(code:4) ".format(filename, title))
                    temp_df = pd.read_excel(filename, dtype=str, skiprows=self.strtoint(skiptop),
                                            skipfooter=self.strtoint(skipbottom))
                else:
                    data_xls = pd.read_excel(filename, sheet_name=None, dtype=str)
                    for key in data_xls:
                        if key == sheet_name:  # 如果有指定excel
                            # print(" sheet_name 策略匹配成功:",sheet_name)
                            print(" {}@{} 匹配 {} 策略成功(code:5) ".format(sheet_name, filename, title))
                            temp_df = pd.read_excel(filename, sheet_name=sheet_name, dtype=str)
                        else:
                            print(" sheet_name为空 策略匹配失败(code:6) ")
                            return pd.DataFrame(columns=["iid"]).head(0)

            else:

                if len(sheet_name) == 0:  # csv 没有 sheetname
                    print(" {} 匹配 {} 策略成功 (code:7)".format(filename, title))
                    try:
                        temp_df = pd.read_csv(filename, encoding="gbk", dtype=str, on_bad_lines='skip',
                                              skiprows=skiptop, skipfooter=skipbottom)
                    except Exception as err:
                        try:
                            temp_df = pd.read_csv(filename, encoding="gbk", dtype=str, skiprows=self.strtoint(skiptop),
                                                  skipfooter=self.strtoint(skipbottom),
                                                  error_bad_lines=False, engine="python").reset_index()
                        except Exception as  err:
                            print(filename, " 异常:", err)
                            # print(filename, "是空表")

                            return pd.DataFrame(columns=["iid"]).head(0)
                else:
                    # print(" {} 匹配 {} 策略成功 (code:7)".format(filename, title))
                    return pd.DataFrame(columns=["iid"]).head(0)

            # 字段重命名
            temp_df = self.rename_column(temp_df, column_rename)
            # 给一些字段设置默认值
            temp_df = self.default_column_set(temp_df, default_columns)
            # 补充空字段，并且按照新的字段标准返回数据
            temp_df = self.add_blankcolumns(temp_df, new_columns)
            return temp_df

        else:
            # 先试着读取下
            print(filename, " read with policy2")
            skiptop = self.strtoint(skiptop)
            skipbottom = self.strtoint(skipbottom)

            with warnings.catch_warnings(record=True):
                warnings.simplefilter("always")
                if os.path.splitext(filename)[1].find("xls") >= 0:
                    if len(sheet_name) > 0:
                        data_xls = pd.read_excel(filename, sheet_name=None, dtype=str)
                        for key in data_xls:
                            if key == sheet_name:
                                print(" {} 匹配 {} 策略成功 (code:83)".format(filename, title))
                                temp_df = pd.read_excel(filename, sheet_name=sheet_name, dtype=str)
                    else:
                        temp_df = pd.read_excel(filename, dtype=str)
                else:
                    # if os.path.splitext(filename)[1].find("xls") < 0:
                    if os.path.splitext(filename)[1].find("csv") >= 0:
                        print(" {} 匹配 {} 策略成功 (code:9)".format(filename, title))
                        try:
                            temp_df = pd.read_csv(filename, encoding="gbk", dtype=str,
                                                  on_bad_lines='skip').reset_index()  # ,decodeing="utf-8"
                            # temp_df = pd.read_csv(filename,encoding="gb18030", dtype=str,error_bad_lines=False).reset_index()  # ,decodeing="utf-8"
                        except Exception as  err:
                            # print(filename, " 异常:", err)
                            # print(filename, "是空表")
                            try:
                                temp_df = pd.read_csv(filename, encoding="gbk", dtype=str,
                                                      error_bad_lines=False, engine="python").reset_index()
                            except Exception as  err:
                                print(filename, " 异常:", err)
                                # print(filename, "是空表")

                                return pd.DataFrame(columns=["iid"]).head(0)

            # 去除末尾合计行，通常第一列，第二列总有一列为空
            cnt1 = temp_df.shape[0]

            print("提前抽查下数据1：", plat, filename)
            print(temp_df.head(3).to_markdown())

            col_1 = "".join(temp_df.columns[0])
            col_2 = "".join(temp_df.columns[1])
            # print("判断为空的两列:", col_1, col_2)
            temp_df.dropna(axis='index', how='all', subset=[col_1, col_2])

            cnt2 = temp_df.shape[0]
            if cnt2 - cnt1 > 0:
                print("末尾有{}个空行".format(cnt2 - cnt1))

            # 自动识别并忽略头部和尾部，如果有必要，进行二次读取
            if ignoretop:
                # if int("0".join(ignoretop)) > 0:
                # temp_df_rows=temp_df.copy()
                temp_df.fillna("", inplace=True)
                # print("抽查：", filename)
                # print(temp_df.head(3).to_markdown())  #
                skiptop = 0
                skipbottom = 0
                skipbottom1 = 0
                skipbottom2 = 0
                shopname = ""

                if len(title_key) > 0:
                    skiptop = self.find_top(temp_df, title_key)

                if len(bottom_key) > 0:
                    skipbottom = self.find_bottom(temp_df, bottom_key)

                # print(row)

                print("总计{}行,忽略头部{}，尾部{}".format(temp_df.shape[0], skiptop, skipbottom))
                #  按照正确的列名重新读取csv文件
                # skipfooter=skipbottom  ,  error_bad_lines=False, engine="python"  不同版本支持不同的写法
                if int(skiptop) + int(skipbottom) > 0:  # 避免无效的两次读取
                    print(filename, " 二次读取!")
                    if os.path.splitext(filename)[1].find("xls") >= 0:
                        # temp_df = pd.read_excel(filename, dtype=str, skiprows=skiptop,skipfooter=skipbottom)
                        temp_df = pd.read_excel(filename, dtype=str, skiprows=skiptop)
                        # temp_df=temp_df[:-skipbottom]

                        skipbottom = 0
                        row2 = temp_df.shape[0] - 1
                        row1 = row2 - min(temp_df.shape[0] - 1, 15)
                        if row1 > row2:
                            row1 = row2
                        row_min, row_max = self.get_key_row(temp_df, row1, row2, bottom_key)
                        print("row_min,row_max=", row_min, row_max)
                        if row_min > 0:
                            skipbottom = row_min
                        elif row_max > 0:
                            skipbottom = row_max

                        # 最后一次清理 合计
                        print("最后一次清理 合计", skipbottom)
                        if skipbottom > 0:
                            temp_df = temp_df[:-skipbottom]

                        print("1.读取{}行".format(temp_df.shape[0]))
                    else:
                        try:
                            temp_df = pd.read_csv(filename, encoding="gbk", dtype=str, on_bad_lines='skip',
                                                  skiprows=skiptop, skipfooter=skipbottom)
                            print("2.读取{}行".format(temp_df.shape[0]))
                        except Exception as err:
                            try:
                                temp_df = pd.read_csv(filename, encoding="gbk", dtype=str, skiprows=skiptop,
                                                      skipfooter=skipbottom,
                                                      error_bad_lines=False, engine="python").reset_index()
                            except Exception as  err:
                                print(filename, " 异常:", err)
                                # print(filename, "是空表")

                                return pd.DataFrame(columns=["iid"]).head(0)
                else:
                    print(filename, " 一次读取!")

                # 去除末尾合计行，通常第一列，第二列总有一列为空
                col_1 = "".join(temp_df.columns[0])
                col_2 = "".join(temp_df.columns[1])
                # print("判断为空的两列:", col_1, col_2)
                temp_df.dropna(axis='index', how='all', subset=[col_1, col_2])

                print("清空合计")
                print(temp_df.head(15).to_markdown())

                temp_df.fillna("", inplace=True)
                print("忽略末尾行：", temp_df.shape[0], skipbottom)

                for col in temp_df.columns:
                    # print("col=",col)
                    if col.find("Unnamed:") >= 0:
                        del temp_df[col]

                print("文件： {} ".format(filename))
                print(" skiptop:{} skipbottom:{} 表行数：{}".format(skiptop, skipbottom, temp_df.shape[0]))

            # 去掉空行
            # if not temp_df.empty:

            # if temp_df.shape[0]<=3:
            #     print("追踪...")
            #     print(temp_df.head(10).to_markdown())
            if 'temp_df' in vars():
                # print(temp_df)
                if len(temp_df.columns) > 0:
                    # print(temp_df.columns.to_list())
                    column1 = np.array(temp_df.columns.to_list())[0]
                    column2 = np.array(temp_df.columns.to_list())[1]
                    # print("列名：", column1, column2)
                    temp_df[column1] = temp_df[column1].astype(str)
                    temp_df[column2] = temp_df[column2].astype(str)
                    temp_df = temp_df[(temp_df[column1].str.len() > 0) | (temp_df[column2].str.len() > 0)]

                temp_df = self.rename_column(temp_df, column_rename)

                # 给一些字段设置默认值
                # print("默认值:")
                # print(default_columns)
                temp_df = self.default_column_set(temp_df, default_columns)

                temp_df = self.add_blankcolumns(temp_df, new_columns)

                print("二次筛选的结果1：")
                print(temp_df.head(5).to_markdown())

                return temp_df

            # 返回空表
            print(filename, "是空表")
            return pd.DataFrame(columns=["iid"]).head(0)

    def clear_df(self,filename,temp_df,title_key,bottom_key):
        # 把头尾清洗掉，然后重新读取
        # if int("0".join(ignoretop)) > 0:
        # temp_df_rows=temp_df.copy()
        temp_df.fillna("", inplace=True)
        # print("抽查：", filename)
        # print(temp_df.head(3).to_markdown())  #
        skiptop = 0
        skipbottom = 0
        skipbottom1 = 0
        skipbottom2 = 0
        shopname = ""

        sheetname=temp_df["sheetname"].iloc[0]

        if len(title_key) > 0:
            skiptop = self.find_top(temp_df, title_key)

        if len(bottom_key) > 0:
            skipbottom = self.find_bottom(temp_df, bottom_key)

        # print(row)

        print(sheetname+" 总计{}行,忽略头部{}，尾部{}".format(temp_df.shape[0], skiptop, skipbottom))
        #  按照正确的列名重新读取csv文件
        # skipfooter=skipbottom  ,  error_bad_lines=False, engine="python"  不同版本支持不同的写法
        if int(skiptop) + int(skipbottom) > 0:  # 避免无效的两次读取
            print(filename, " 二次读取!")
            if os.path.splitext(filename)[1].find("xls") >= 0:
                # temp_df = pd.read_excel(filename, dtype=str, skiprows=skiptop,skipfooter=skipbottom)
                temp_df = pd.read_excel(filename, dtype=str, skiprows=skiptop)
                #
                if skipbottom<temp_df.shape[0]:
                    print("切除尾巴:",skipbottom)
                    temp_df = temp_df[:-skipbottom]
                else:
                    skipbottom = 0
                    row2 = temp_df.shape[0] - 1
                    row1 = row2 - min(temp_df.shape[0] - 1, 15)
                    if row1 > row2:
                        row1 = row2
                    row_min, row_max = self.get_key_row(temp_df, row1, row2, bottom_key)
                    print("row_min,row_max=", row_min, row_max)
                    if row_min > 0:
                        skipbottom = row_min
                    elif row_max > 0:
                        skipbottom = row_max

                    # 最后一次清理 合计
                    print("最后一次清理 合计", skipbottom)
                    if skipbottom > 0:
                        temp_df = temp_df[:-skipbottom]
                    print("抽查111:")
                    print(temp_df.head(10).to_markdown())

                print("1.读取{}行".format(temp_df.shape[0]))
            else:
                try:
                    temp_df = pd.read_csv(filename, encoding="gbk", dtype=str, on_bad_lines='skip',
                                          skiprows=skiptop, skipfooter=skipbottom)
                    print("2.读取{}行".format(temp_df.shape[0]))
                except Exception as err:
                    try:
                        temp_df = pd.read_csv(filename, encoding="gbk", dtype=str, skiprows=skiptop,
                                              skipfooter=skipbottom,
                                              error_bad_lines=False, engine="python").reset_index()
                    except Exception as  err:
                        print(filename, " 异常:", err)
                        # print(filename, "是空表")

                        return pd.DataFrame(columns=["iid"]).head(0)
        else:
            print(filename, " 一次读取!")

        # 去除末尾合计行，通常第一列，第二列总有一列为空
        col_1 = "".join(temp_df.columns[0])
        col_2 = "".join(temp_df.columns[1])
        # print("判断为空的两列:", col_1, col_2)
        temp_df.dropna(axis='index', how='all', subset=[col_1, col_2])


        print("清空合计")
        print(temp_df.head(15).to_markdown())

        temp_df.fillna("", inplace=True)
        print("忽略末尾行：", temp_df.shape[0], skipbottom)

        for col in temp_df.columns:
            # print("col=",col)
            if col.find("Unnamed:") >= 0:
                del temp_df[col]

        temp_df["sheetname"] = sheetname

        print("文件： {} ".format(filename))
        print(" skiptop:{} skipbottom:{} 表行数：{}".format(skiptop, skipbottom, temp_df.shape[0]))
        return temp_df


    def read_table_with_policy(self,filename, item):
        df_list=[]
        # sheet_name, ignoretop, title_key, bottom_key, valid_columns,new_columns,column_rename,default_columns
        plat, sheet_name,title, ignoretop, skiptop,skipbottom,title_key, bottom_key, default_columns, new_columns, column_rename, default_columns = \
        item[0], item[1], item[2], item[3], item[4], item[5], item[6], item[7], item[8], item[9], item[10], item[11]

        print("读取参数配置:",item)

        skiptop=str(self.strtoint(skiptop))
        skipbottom=str(self.strtoint(skipbottom))

        if  len(title.strip())==0:
            title=sheet_name
            if len(title.strip()) == 0:
                title = plat

        print("过滤策略2：",self.strtoint(skiptop),self.strtoint(skipbottom),self.strtoint(ignoretop))
        if (self.strtoint(skiptop) + self.strtoint(skipbottom) + self.strtoint(ignoretop)== 0):
            # 无需加工，直接读取
            print(filename," read no filter")
            with warnings.catch_warnings(record=True):
                warnings.simplefilter("always")
                if os.path.splitext(filename)[1].find("xls") >= 0:
                    if len(sheet_name) > 0:
                        data_xls = pd.read_excel(filename, sheet_name=None, dtype=str)
                        for key in data_xls:
                            if key == sheet_name:
                                # print(" sheet_name 策略匹配成功:",sheet_name)
                                print(" {}@{} 匹配 {} 策略成功(code:1) ".format(sheet_name,filename, title))
                                temp_df = pd.read_excel(filename, sheet_name=sheet_name, dtype=str)
                                temp_df["sheetname"] = key
                                df_list.append(temp_df)
                    else:
                        print(" sheet_name为空 策略匹配成功(code:2) ")
                        # temp_df = pd.read_excel(filename, dtype=str)
                        data_xls = pd.read_excel(filename, sheet_name=None, dtype=str)
                        for key in data_xls:
                            temp_df = self.read_worksheet_with_policy(filename,key, item)
                            temp_df["sheetname"] = key
                            df_list.append(temp_df)
                else:
                    # if os.path.splitext(filename)[1].find("xls") < 0:
                    if os.path.splitext(filename)[1].find("csv") >= 0:
                        if len(sheet_name)==0:  # csv 没有 sheetname
                            print(" {} 匹配 {} 策略成功(code:3) ".format(filename,title))
                            try:
                                temp_df = pd.read_csv(filename, encoding="gbk", dtype=str,
                                                      on_bad_lines='skip').reset_index()  # ,decodeing="utf-8"
                                # temp_df = pd.read_csv(filename,encoding="gb18030", dtype=str,error_bad_lines=False).reset_index()  # ,decodeing="utf-8"
                                print("读取异常")
                                df_list.append(temp_df)
                            except Exception as  err:
                                # print(filename, " 异常:", err)
                                # print(filename, "是空表")
                                try:
                                    temp_df = pd.read_csv(filename, encoding="gbk", dtype=str,
                                                          error_bad_lines=False, engine="python").reset_index()
                                    print("读取异常")
                                    df_list.append(temp_df)
                                except Exception as  err:
                                    print(filename, " 异常:", err)
                                    # print(filename, "是空表")
                                    df_list.append(pd.DataFrame(columns=["iid"]).head(0))

                                    return df_list
                        else:
                            df_list.append(pd.DataFrame(columns=["iid"]).head(0))
                            return df_list
            # 字段重命名
            for temp_df in df_list:
                temp_df = self.rename_column(temp_df, column_rename)
                # 给一些字段设置默认值
                temp_df = self.default_column_set(temp_df, default_columns)
                # 补充空字段，并且按照新的字段标准返回数据
                temp_df = self.add_blankcolumns(temp_df, new_columns)
            return df_list

        elif ((self.strtoint(skiptop)>0) and (self.strtoint(skipbottom)>0) and (self.strtoint(ignoretop)>0)):
            #直接按照约定过滤标准执行 要求过滤无效行，并且指定忽略的头部行以及尾部行
            print(filename, " read with fiter ",skipbottom)
            if os.path.splitext(filename)[1].find("xls") >= 0:
                # 读取excel
                if len(sheet_name) == 0:  # 无需指定 sheetname
                    print(" {} 匹配 {} 策略成功(code:4) ".format(filename, title))
                    data_xls = pd.read_excel(filename, sheet_name=None, dtype=str)
                    for key in data_xls:
                        # temp_df = pd.read_excel(filename, sheet_name=key, dtype=str)
                        temp_df = pd.read_excel(filename,  sheet_name=key,dtype=str, skiprows=self.strtoint(skiptop), skipfooter=self.strtoint(skipbottom))
                        df_list.append(temp_df)
                        # temp_df = pd.read_excel(filename, dtype=str, skiprows=self.strtoint(skiptop), skipfooter=self.strtoint(skipbottom))
                else:
                    data_xls = pd.read_excel(filename, sheet_name=None, dtype=str)
                    for key in data_xls:
                        if key == sheet_name:  # 如果有指定excel
                            # print(" sheet_name 策略匹配成功:",sheet_name)
                            print(" {}@{} 匹配 {} 策略成功(code:5) ".format(sheet_name, filename, title))
                            temp_df = pd.read_excel(filename, sheet_name=sheet_name, dtype=str)
                            temp_df["sheetname"] = key
                            df_list.append(temp_df)
                        else:
                            print(" sheet_name为空 策略匹配失败(code:6) ")
                            return pd.DataFrame(columns=["iid"]).head(0)

            else:

                if len(sheet_name) == 0:  # csv 没有 sheetname
                    print(" {} 匹配 {} 策略成功 (code:7)".format(filename,  title))
                    try:
                        temp_df = pd.read_csv(filename, encoding="gbk", dtype=str, on_bad_lines='skip',
                                              skiprows=skiptop, skipfooter=skipbottom)
                        df_list.append(temp_df)
                    except Exception as err:
                        try:
                            temp_df = pd.read_csv(filename, encoding="gbk", dtype=str, skiprows= self.strtoint(skiptop) ,
                                                  skipfooter=self.strtoint(skipbottom),
                                                  error_bad_lines=False, engine="python").reset_index()
                            print("读取异常")
                            df_list.append(temp_df)
                        except Exception as  err:
                            print(filename, " 异常:", err)
                            # print(filename, "是空表")

                            # return pd.DataFrame(columns=["iid"]).head(0)
                            df_list.append(pd.DataFrame(columns=["iid"]).head(0))
                            return df_list
                else:
                    # print(" {} 匹配 {} 策略成功 (code:7)".format(filename, title))
                    return pd.DataFrame(columns=["iid"]).head(0)

            # 字段重命名
            for temp_df in df_list:
                temp_df = self.rename_column(temp_df, column_rename)
                # 给一些字段设置默认值
                temp_df = self.default_column_set(temp_df, default_columns)
                # 补充空字段，并且按照新的字段标准返回数据
                temp_df = self.add_blankcolumns(temp_df, new_columns)
            return df_list

        else:
            # 先试着读取下
            print(filename, " read with policy1")
            skiptop =   self.strtoint(skiptop)
            skipbottom =   self.strtoint(skipbottom)

            with warnings.catch_warnings(record=True):
                warnings.simplefilter("always")
                if os.path.splitext(filename)[1].find("xls") >= 0:
                    if len(sheet_name) > 0:
                        data_xls = pd.read_excel(filename, sheet_name=None, dtype=str)
                        for key in data_xls:
                            if key == sheet_name:
                                print(" {} 匹配 {} 策略成功 (code:81)".format(filename, title))
                                temp_df = pd.read_excel(filename, sheet_name=sheet_name, dtype=str)
                                temp_df["sheetname"] = key
                                df_list.append(temp_df)
                    else:
                        # temp_df = pd.read_excel(filename, dtype=str)
                        data_xls = pd.read_excel(filename, sheet_name=None, dtype=str)
                        for key in data_xls:
                            print(" {} 匹配 {} 策略成功 (code:82) @sheet= {}".format(filename, title,key))
                            temp_df = pd.read_excel(filename, sheet_name=key, dtype=str)
                            temp_df["sheetname"]=key
                            df_list.append(temp_df)

                else:
                    # if os.path.splitext(filename)[1].find("xls") < 0:
                    if os.path.splitext(filename)[1].find("csv") >= 0:
                        print(" {} 匹配 {} 策略成功 (code:9)".format(filename, title))
                        try:
                            temp_df = pd.read_csv(filename, encoding="gbk", dtype=str,
                                                  on_bad_lines='skip').reset_index()  # ,decodeing="utf-8"
                            # temp_df = pd.read_csv(filename,encoding="gb18030", dtype=str,error_bad_lines=False).reset_index()  # ,decodeing="utf-8"
                            df_list.append(temp_df)
                        except Exception as  err:
                            # print(filename, " 异常:", err)
                            # print(filename, "是空表")
                            try:
                                temp_df = pd.read_csv(filename, encoding="gbk", dtype=str,
                                                      error_bad_lines=False,engine="python").reset_index()
                                print("读取异常")
                                df_list.append(temp_df)
                            except Exception as  err:
                                print(filename, " 异常:", err)
                                # print(filename, "是空表")
                                df_list.append(pd.DataFrame(columns=["iid"]).head(0))

                                # return pd.DataFrame(columns=["iid"]).head(0)
                                return df_list

            df_list_new=[]
            sheet_index=1
            for temp_df in df_list:
                # 去除末尾合计行，通常第一列，第二列总有一列为空
                cnt1=temp_df.shape[0]

                print("提前抽查下数据{}/{}(cnt={})：".format(sheet_index,len(df_list),temp_df.shape[0]),plat,filename)
                print(temp_df.head(3).to_markdown())

                col_1 = "".join(temp_df.columns[0])
                col_2 = "".join(temp_df.columns[1])
                # print("判断为空的两列:", col_1, col_2)
                temp_df.dropna(axis='index', how='all', subset=[col_1, col_2])

                cnt2 = temp_df.shape[0]
                if cnt2-cnt1>0:
                    print("末尾有{}个空行".format(cnt2-cnt1))

                # 自动识别并忽略头部和尾部，如果有必要，进行二次读取
                if ignoretop:
                    temp_df=self.clear_df( filename, temp_df, title_key, bottom_key)

                print("debug22")
                print(temp_df.head(10).to_markdown())
                # 修改字段名
                if len(temp_df.columns) > 0:
                    # print(temp_df.columns.to_list())
                    column1 = np.array(temp_df.columns.to_list())[0]
                    column2 = np.array(temp_df.columns.to_list())[1]
                    # print("列名：", column1, column2)
                    temp_df[column1] = temp_df[column1].astype(str)
                    temp_df[column2] = temp_df[column2].astype(str)
                    temp_df = temp_df[(temp_df[column1].str.len() > 0) | (temp_df[column2].str.len() > 0)]

                temp_df=self.rename_column( temp_df, column_rename)

                print("debug23")
                print(temp_df.head(3).to_markdown())


                # 给一些字段设置默认值
                # print("默认值:")
                # print(default_columns)
                temp_df=self.default_column_set( temp_df, default_columns)

                print("debug24")
                print(temp_df.head(3).to_markdown())

                temp_df=self.add_blankcolumns( temp_df, new_columns)

                print("二次筛选的结果2：")
                print(temp_df.head(5).to_markdown())
                df_list_new.append(temp_df)

                sheet_index = sheet_index + 1

                # 去掉空行
            # if not temp_df.empty:

            # if temp_df.shape[0]<=3:
            #     print("追踪...")
            #     print(temp_df.head(10).to_markdown())
            # if 'temp_df' in vars():
            #     # print(temp_df)
            #     return temp_df

            # 返回空表
            # print(filename, "是空表")
            # return pd.DataFrame(columns=["iid"]).head(0)

            return df_list_new

    def list_all_files(self,rootdir, filekey_list):

        # filekey_list="2020|2019"  or
        # filekey_list="2019(.*)海外旗舰"  and
        # filekey_list = "(?!京东)"

        if len(filekey_list) > 0:
            filekey_list = filekey_list.replace(",", " ").replace("，", " ")
            filekey = filekey_list.split(" ")
            pass
        else:
            filekey = ''

        _files = []
        list = os.listdir(rootdir)  # 列出文件夹下所有的目录与文件
        for i in range(0, len(list)):
            path = os.path.join(rootdir, list[i])
            if os.path.isdir(path):
                _files.extend(self.list_all_files(path, filekey_list))
            if os.path.isfile(path):
                if (path.find("~") < 0) and (path.find(".DS_Store") < 0) and (path.find("._") < 0):  # 带~符号表示临时文件，不读取
                    # if len(filekey_list) > 0:
                    #     t=re.search(filekey_list,path)
                    #     if t:  # 如果匹配成功
                    #         _files.append(path)

                    if len(filekey) > 0:
                        break_flag = False
                        for key in filekey:
                            if not break_flag:
                                # print(path)

                                # 简化版的不包含(类似正则表达式)  !京东 = 不包含京东
                                if ((len(key.replace("!", "")) + 1 == len(key)) and (key.find("?") < 0) and (
                                        key.find(".") < 0) and (key.find("(") < 0) and (key.find(")") < 0)):
                                    if path.find(key.replace("!", "")) >= 0:
                                        # 要求不包含，结果找到了！
                                        print("要求不包含{}，结果找到了！".format(key.replace("!", "")))
                                        break_flag = True
                                else:
                                    t = re.search(key, path)
                                    if t:  # 如果匹配成功
                                        pass
                                    else:
                                        # 只要有一项匹配不成功，则自动退出，认为不符合条件
                                        break_flag = True

                        if not break_flag:
                            _files.append(path)

                    else:
                        _files.append(path)

        # print(_files)
        return _files

    def get_files_df(self,rootdir, filekey):
        filelist = self.list_all_files(rootdir, filekey)
        # print(filelist)
        if len(filelist) > 0:
            mySeries = pd.Series(filelist)
            df = pd.DataFrame(mySeries)
            df.columns = ["filename"]
            # df["filename"] = df["filename"].apply(lambda x: x.strip().replace(" ", ""))

            # print(df.to_markdown())
            print(df.to_markdown())
            return df
        else:
            print("没有发现符合条件的文件！")
            # 创建一个空表
            return pd.DataFrame(columns=["filename"])

    #
    def df_fit_template(self,df_shop,  item,  filename, dd,  searchcolumn,  searchword):

        # df_box=[]
        model_index=0

        plat, sheetname, title, ignoretop, skiptop, skipbottom, title_key, bottom_key, default_columns, new_columns, column_rename, default_columns = \
        item[0], item[1], item[2], item[3], item[4], item[5], item[6], item[7], item[8], item[9], item[10], item[11]

        # 重新确认sheet_name
        sheetname=dd["sheetname"].iloc[0]

        if len(searchcolumn)>0:
            print("搜索字段：", searchcolumn)
            foundcolvalue = False
            for selcol in searchcolumn.split(","):
                if selcol in dd.columns:
                    print("搜索字段...", selcol)
                    if not foundcolvalue:
                        foundcolvalue = True
                        # print("发现字段!", selcol,searchword)
                        dd = dd[dd[selcol].str.contains(searchword, na=False)]
                        if dd.shape[0] > 0:
                            print("发现关键词:", searchword, " 文件：", filename)

            if not foundcolvalue:
                dd = dd.head(0)

        # 如果不是空表
        if dd.shape[0] > 0:
            dd["filename"] = filename
            dd["sheetname"] = sheetname
            dd["title"] = title
            dd["policy"] = plat

            if "店铺名称" in dd.columns:
                print("更新店铺名称")
                dd["店铺名称"].fillna("", inplace=True)
                shop_name = get_shopname(df_shop, filename)[1]
                print("更新店铺名称:", shop_name)
                # dd[dd["店铺名称"].str.len()==0]["店铺名称"]= get_shopname(df_shop,file["filename"])[1]
                dd["店铺名称"] = shop_name


            # 如果行数为0，则忽略，不需要关心格式
            if dd.shape[0] == 0:
                print("忽略空表：", filename)
            elif filename.find("错误") > 0:
                print("忽略错误的表格：", filename)
            else:
                # if dd.shape[0] > 0:
                # 将列名转换成list
                if_exist = False
                iseq = 0
                print("字段列表:",dd.columns.to_list())
                print("检查数据:",dd.head(10).to_markdown())
                dd_columns = dd.columns.to_list()
                for r in self.file_columns_list:
                    # 如果相同，已经存在
                    # print("对比1:",dd_columns)
                    # print("对比2:",r[1:])

                    if operator.eq(dd_columns, r[1:]):
                        # if dd_columns==r[1:]:
                        if_exist = True
                        break

                    # 模板不存在
                    iseq = iseq + 1

                # 模板已经存在
                if if_exist:
                    # print("模板已经存在！：", file["filename"])
                    # print("模板1:", dd_columns)
                    # print("模板2:",iseq)
                    # for t in file_columns_list:
                    #     print(t)

                    if iseq == 0:
                        if 'df0' in vars():
                            df0 = df0.append(dd)
                        else:
                            df0 = dd
                    elif iseq == 1:
                        if 'df1' in vars():
                            df1 = df1.append(dd)
                        else:
                            df1 = dd
                    elif iseq == 2:
                        if 'df2' in vars():
                            df2 = df2.append(dd)
                        else:
                            df2 = dd
                    elif iseq == 3:
                        if 'df3' in vars():
                            df3 = df3.append(dd)
                        else:
                            df3 = dd
                    elif iseq == 4:
                        if 'df4' in vars():
                            df4 = df4.append(dd)
                        else:
                            df4 = dd
                    elif iseq == 5:
                        if 'df5' in vars():
                            df5 = df5.append(dd)
                        else:
                            df5 = dd
                    elif iseq == 6:
                        if 'df6' in vars():
                            df6 = df6.append(dd)
                        else:
                            df6 = dd
                    elif iseq == 7:
                        if 'df7' in vars():
                            df7 = df7.append(dd)
                        else:
                            df7 = dd
                    elif iseq == 8:
                        if 'df8' in vars():
                            df8 = df8.append(dd)
                        else:
                            df8 = dd
                    elif iseq == 9:
                        if 'df9' in vars():
                            df9 = df9.append(dd)
                        else:
                            df9 = dd
                    elif iseq == 10:
                        if 'df10' in vars():
                            df10 = df10.append(dd)
                        else:
                            df10 = dd
                    elif iseq == 11:
                        if 'df11' in vars():
                            df11 = df11.append(dd)
                        else:
                            df11 = dd
                    elif iseq == 12:
                        if 'df12' in vars():
                            df12 = df12.append(dd)
                        else:
                            df12 = dd
                    elif iseq == 13:
                        if 'df13' in vars():
                            df13 = df13.append(dd)
                        else:
                            df13 = dd


                else:
                    #     print("新增字段模板：",file["filename"])
                    #     print("模板1:", dd_columns)
                    #     print("模板2:")
                    #     for t in file_columns_list:
                    #         print(t)

                    print("新增模板：", dd_columns)

                    # 从第0列插入字段
                    dd_columns.insert(0, filename)
                    # 从第1列插入字段
                    # dd_columns.insert(1,sheetname)

                    self.file_columns_list.append(dd_columns)

                    iseq = len(self.file_columns_list) - 1
                    if iseq == 0:
                        df0 = dd
                    elif iseq == 1:
                        df1 = dd
                    elif iseq == 2:
                        df2 = dd
                    elif iseq == 3:
                        df3 = dd
                    elif iseq == 4:
                        df4 = dd
                    elif iseq == 5:
                        df5 = dd
                    elif iseq == 6:
                        df6 = dd
                    elif iseq == 7:
                        df7 = dd
                    elif iseq == 8:
                        df8 = dd
                    elif iseq == 9:
                        df9 = dd
                    elif iseq == 10:
                        df10 = dd
                    elif iseq == 11:
                        df11 = dd
                    elif iseq == 12:
                        df12 = dd
                    elif iseq == 13:
                        df13 = dd



                if iseq == 0:
                      return iseq, df0
                elif iseq == 1:
                    return iseq, df1
                elif iseq == 2:
                    return iseq, df2
                elif iseq == 3:
                    return iseq, df3
                elif iseq == 4:
                    return iseq, df4
                elif iseq == 5:
                    return iseq, df5
                elif iseq == 6:
                    return iseq, df6
                elif iseq == 7:
                    return iseq, df7
                elif iseq == 8:
                    return iseq, df8
                elif iseq == 9:
                    return iseq, df9
                elif iseq == 10:
                    return iseq, df10
                elif iseq == 11:
                    return iseq, df11
                elif iseq == 12:
                    return iseq, df12
                elif iseq == 13:
                    return iseq, df13


                    # print(file["filename"], dd.shape[0])
                    # print("进度表：{}/{}  文件{}，行数{}".format(index + 1, df_files.shape[0], file["filename"], dd.shape[0]))

        # if 'df0' in vars():
        #     df_box.append(df0)
        # if 'df1' in vars():
        #     df_box.append(df1)
        # if 'df2' in vars():
        #     df_box.append(df2)
        # if 'df3' in vars():
        #     df_box.append(df3)
        # if 'df4' in vars():
        #     df_box.append(df4)
        # if 'df5' in vars():
        #     df_box.append(df5)
        # if 'df6' in vars():
        #     df_box.append(df6)
        # if 'df7' in vars():
        #     df_box.append(df7)
        # if 'df8' in vars():
        #     df_box.append(df8)
        # if 'df9' in vars():
        #     df_box.append(df9)
        # if 'df10' in vars():
        #     df_box.append(df10)
        # if 'df11' in vars():
        #     df_box.append(df11)
        # if 'df12' in vars():
        #     df_box.append(df12)
        # if 'df13' in vars():
        #     df_box.append(df13)

        # return df0,df1,df2,df3,df4,df5,df6,df7,df8,df9,df10,df11,df12,df13
        return iseq, dd

    def read_files_table_byword(self,rootdir, filekey,searchcolumn,searchword):
        # 下载店铺数据
        df_shop=read_shop()

        # title_key,bottom_key
        df_files = self.get_files_df(rootdir, filekey)
        df_box = []
        files_count=df_files.shape[0]
        for index, file in df_files.iterrows():
            # 根据文件名，匹配不同的政策
            k=0
            policy_list= self.get_policy(file["filename"])
            print("读取政策1")
            # print(file["filename"] )
            # print( policy_list)
            for item in policy_list:
                # 循环匹配每一个sheet
                # column_rename = sheet["rename"]
                # sheetname, ignoretop, title_key, bottom_key, default_columns=policy["sheetname"], policy["ignoretop"], policy["title_key"], policy["bottom_key"],policy["default_columns"]
                plat,sheetname,title, ignoretop,skiptop,skipbottom, title_key, bottom_key, default_columns,new_columns,column_rename,default_columns = item[0], item[1], item[2], item[3], item[4], item[5], item[6], item[7],item[8],item[9],item[10],item[11]

                if ignoretop == "1":
                    ignoretop = True
                else:
                    ignoretop = False

                # 文件只有匹配成功才能被读取
                # print("读取文件:", file["filename"])
                k=k+1
                # dd = self.read_table_with_policy(file["filename"], sheetname, ignoretop, title_key, bottom_key, default_columns,new_columns,column_rename,default_columns)

                # df = Xiaohongshu("/Users/lichunlei/Downloads/小红书测试/小红书Dentyl Active旗舰店201902-账单.xlsx").get_table()
                # 小红书特殊处理
                if (file["filename"].find("小红书")>=0) or (file["filename"].find("XHS")>=0):
                    dd = Xiaohongshu(file["filename"]).get_table()
                    filename = file["filename"]
                    iseq, dd = self.df_fit_template( df_shop, item, filename, dd, searchcolumn, searchword)

                    if iseq == 0:
                        if 'df0' in vars():
                            df0 = df0.append(dd)
                        else:
                            df0 = dd
                    elif iseq == 1:
                        if 'df1' in vars():
                            df1 = df1.append(dd)
                        else:
                            df1 = dd
                    elif iseq == 2:
                        if 'df2' in vars():
                            df2 = df2.append(dd)
                        else:
                            df2 = dd
                    elif iseq == 3:
                        if 'df3' in vars():
                            df3 = df3.append(dd)
                        else:
                            df3 = dd
                    elif iseq == 4:
                        if 'df4' in vars():
                            df4 = df4.append(dd)
                        else:
                            df4 = dd
                    elif iseq == 5:
                        if 'df5' in vars():
                            df5 = df5.append(dd)
                        else:
                            df5 = dd
                    elif iseq == 6:
                        if 'df6' in vars():
                            df6 = df6.append(dd)
                        else:
                            df6 = dd
                    elif iseq == 7:
                        if 'df7' in vars():
                            df7 = df7.append(dd)
                        else:
                            df7 = dd
                    elif iseq == 8:
                        if 'df8' in vars():
                            df8 = df8.append(dd)
                        else:
                            df8 = dd
                    elif iseq == 9:
                        if 'df9' in vars():
                            df9 = df9.append(dd)
                        else:
                            df9 = dd
                    elif iseq == 10:
                        if 'df10' in vars():
                            df10 = df10.append(dd)
                        else:
                            df10 = dd
                    elif iseq == 11:
                        if 'df11' in vars():
                            df11 = df11.append(dd)
                        else:
                            df11 = dd
                    elif iseq == 12:
                        if 'df12' in vars():
                            df12 = df12.append(dd)
                        else:
                            df12 = dd
                    elif iseq == 13:
                        if 'df13' in vars():
                            df13 = df13.append(dd)
                        else:
                            df13 = dd
                elif (file["filename"].find("考拉") >= 0)  :
                    dd = Kaola(file["filename"]).get_table()
                    filename = file["filename"]
                    iseq, dd = self.df_fit_template( df_shop, item, filename, dd, searchcolumn, searchword)

                    if iseq == 0:
                        if 'df0' in vars():
                            df0 = df0.append(dd)
                        else:
                            df0 = dd
                    elif iseq == 1:
                        if 'df1' in vars():
                            df1 = df1.append(dd)
                        else:
                            df1 = dd
                    elif iseq == 2:
                        if 'df2' in vars():
                            df2 = df2.append(dd)
                        else:
                            df2 = dd
                    elif iseq == 3:
                        if 'df3' in vars():
                            df3 = df3.append(dd)
                        else:
                            df3 = dd
                    elif iseq == 4:
                        if 'df4' in vars():
                            df4 = df4.append(dd)
                        else:
                            df4 = dd
                    elif iseq == 5:
                        if 'df5' in vars():
                            df5 = df5.append(dd)
                        else:
                            df5 = dd
                    elif iseq == 6:
                        if 'df6' in vars():
                            df6 = df6.append(dd)
                        else:
                            df6 = dd
                    elif iseq == 7:
                        if 'df7' in vars():
                            df7 = df7.append(dd)
                        else:
                            df7 = dd
                    elif iseq == 8:
                        if 'df8' in vars():
                            df8 = df8.append(dd)
                        else:
                            df8 = dd
                    elif iseq == 9:
                        if 'df9' in vars():
                            df9 = df9.append(dd)
                        else:
                            df9 = dd
                    elif iseq == 10:
                        if 'df10' in vars():
                            df10 = df10.append(dd)
                        else:
                            df10 = dd
                    elif iseq == 11:
                        if 'df11' in vars():
                            df11 = df11.append(dd)
                        else:
                            df11 = dd
                    elif iseq == 12:
                        if 'df12' in vars():
                            df12 = df12.append(dd)
                        else:
                            df12 = dd
                    elif iseq == 13:
                        if 'df13' in vars():
                            df13 = df13.append(dd)
                        else:
                            df13 = dd
                else:
                    ddd = self.read_table_with_policy(file["filename"], item)
                    # 不同的dd对应不同的模板
                    for dd in ddd:
                        iseq,dd =  self.df_fit_template( df_shop, item, file["filename"], dd, searchcolumn, searchword)

                        if iseq == 0:
                            if 'df0' in vars():
                                df0 = df0.append(dd)
                            else:
                                df0 = dd
                        elif iseq == 1:
                            if 'df1' in vars():
                                df1 = df1.append(dd)
                            else:
                                df1 = dd
                        elif iseq == 2:
                            if 'df2' in vars():
                                df2 = df2.append(dd)
                            else:
                                df2 = dd
                        elif iseq == 3:
                            if 'df3' in vars():
                                df3 = df3.append(dd)
                            else:
                                df3 = dd
                        elif iseq == 4:
                            if 'df4' in vars():
                                df4 = df4.append(dd)
                            else:
                                df4 = dd
                        elif iseq == 5:
                            if 'df5' in vars():
                                df5 = df5.append(dd)
                            else:
                                df5 = dd
                        elif iseq == 6:
                            if 'df6' in vars():
                                df6 = df6.append(dd)
                            else:
                                df6 = dd
                        elif iseq == 7:
                            if 'df7' in vars():
                                df7 = df7.append(dd)
                            else:
                                df7 = dd
                        elif iseq == 8:
                            if 'df8' in vars():
                                df8 = df8.append(dd)
                            else:
                                df8 = dd
                        elif iseq == 9:
                            if 'df9' in vars():
                                df9 = df9.append(dd)
                            else:
                                df9 = dd
                        elif iseq == 10:
                            if 'df10' in vars():
                                df10 = df10.append(dd)
                            else:
                                df10 = dd
                        elif iseq == 11:
                            if 'df11' in vars():
                                df11 = df11.append(dd)
                            else:
                                df11 = dd
                        elif iseq == 12:
                            if 'df12' in vars():
                                df12 = df12.append(dd)
                            else:
                                df12 = dd
                        elif iseq == 13:
                            if 'df13' in vars():
                                df13 = df13.append(dd)
                            else:
                                df13 = dd

                print("进度表：  {}/{} [sheet {}/{}]  文件:{} ，行数{}".format(index + 1, files_count, k,
                                                                      len(policy_list), filename,
                                                                      dd.shape[0]))

        if 'df0' in vars():
            df_box.append(df0)
        if 'df1' in vars():
            df_box.append(df1)
        if 'df2' in vars():
            df_box.append(df2)
        if 'df3' in vars():
            df_box.append(df3)
        if 'df4' in vars():
            df_box.append(df4)
        if 'df5' in vars():
            df_box.append(df5)
        if 'df6' in vars():
            df_box.append(df6)
        if 'df7' in vars():
            df_box.append(df7)
        if 'df8' in vars():
            df_box.append(df8)
        if 'df9' in vars():
            df_box.append(df9)
        if 'df10' in vars():
            df_box.append(df10)
        if 'df11' in vars():
            df_box.append(df11)
        if 'df12' in vars():
            df_box.append(df12)
        if 'df13' in vars():
            df_box.append(df13)

        print("最终的字段列表：")
        print(self.file_columns_list)
        print("字段模板共有：{} 个".format(len(self.file_columns_list)))

        return df_box


    def read_files_table(self,rootdir, filekey):
        # title_key,bottom_key
        df_shop = read_shop()
        df_files = self.get_files_df(rootdir, filekey)
        df_box = []
        files_count=df_files.shape[0]
        for index, file in df_files.iterrows():
            # 根据文件名，匹配不同的政策
            k=0
            policy_list= self.get_policy(file["filename"])
            print("读取政策2")
            # print(file["filename"] )
            # print( policy_list)
            for item in policy_list:
                # 循环匹配每一个sheet
                # column_rename = sheet["rename"]
                # sheetname, ignoretop, title_key, bottom_key, default_columns=policy["sheetname"], policy["ignoretop"], policy["title_key"], policy["bottom_key"],policy["default_columns"]
                plat,sheetname, ignoretop,skiptop,skipbottom, title_key, bottom_key, default_columns,new_columns,column_rename,default_columns = item[0], item[1], item[2], item[3], item[4], item[5], item[6], item[7],item[8],item[9],item[10]

                if ignoretop == "1":
                    ignoretop = True
                else:
                    ignoretop = False

                # 文件只有匹配成功才能被读取
                # print("读取文件:", file["filename"])
                k=k+1
                # dd = self.read_table_with_policy(file["filename"], sheetname, ignoretop, title_key, bottom_key, default_columns,new_columns,column_rename,default_columns)
                # df = Xiaohongshu("/Users/lichunlei/Downloads/小红书测试/小红书Dentyl Active旗舰店201902-账单.xlsx").get_table()
                # 小红书特殊处理
                if (file["filename"].find("小红书")>=0) or (file["filename"].find("XHS")>=0):
                    dd = Xiaohongshu(file["filename"]).get_table()
                    if iseq == 0:
                        if 'df0' in vars():
                            df0 = df0.append(dd)
                        else:
                            df0 = dd
                    elif iseq == 1:
                        if 'df1' in vars():
                            df1 = df1.append(dd)
                        else:
                            df1 = dd
                    elif iseq == 2:
                        if 'df2' in vars():
                            df2 = df2.append(dd)
                        else:
                            df2 = dd
                    elif iseq == 3:
                        if 'df3' in vars():
                            df3 = df3.append(dd)
                        else:
                            df3 = dd
                    elif iseq == 4:
                        if 'df4' in vars():
                            df4 = df4.append(dd)
                        else:
                            df4 = dd
                    elif iseq == 5:
                        if 'df5' in vars():
                            df5 = df5.append(dd)
                        else:
                            df5 = dd
                    elif iseq == 6:
                        if 'df6' in vars():
                            df6 = df6.append(dd)
                        else:
                            df6 = dd
                    elif iseq == 7:
                        if 'df7' in vars():
                            df7 = df7.append(dd)
                        else:
                            df7 = dd
                    elif iseq == 8:
                        if 'df8' in vars():
                            df8 = df8.append(dd)
                        else:
                            df8 = dd
                    elif iseq == 9:
                        if 'df9' in vars():
                            df9 = df9.append(dd)
                        else:
                            df9 = dd
                    elif iseq == 10:
                        if 'df10' in vars():
                            df10 = df10.append(dd)
                        else:
                            df10 = dd
                    elif iseq == 11:
                        if 'df11' in vars():
                            df11 = df11.append(dd)
                        else:
                            df11 = dd
                    elif iseq == 12:
                        if 'df12' in vars():
                            df12 = df12.append(dd)
                        else:
                            df12 = dd
                    elif iseq == 13:
                        if 'df13' in vars():
                            df13 = df13.append(dd)
                        else:
                            df13 = dd

                    if 'df0' in vars():
                        df_box.append(df0)
                    if 'df1' in vars():
                        df_box.append(df1)
                    if 'df2' in vars():
                        df_box.append(df2)
                    if 'df3' in vars():
                        df_box.append(df3)
                    if 'df4' in vars():
                        df_box.append(df4)
                    if 'df5' in vars():
                        df_box.append(df5)
                    if 'df6' in vars():
                        df_box.append(df6)
                    if 'df7' in vars():
                        df_box.append(df7)
                    if 'df8' in vars():
                        df_box.append(df8)
                    if 'df9' in vars():
                        df_box.append(df9)
                    if 'df10' in vars():
                        df_box.append(df10)
                    if 'df11' in vars():
                        df_box.append(df11)
                    if 'df12' in vars():
                        df_box.append(df12)
                    if 'df13' in vars():
                        df_box.append(df13)

                elif (file["filename"].find("考拉") >= 0)  :
                    dd = Kaola(file["filename"]).get_table()
                    if iseq == 0:
                        if 'df0' in vars():
                            df0 = df0.append(dd)
                        else:
                            df0 = dd
                    elif iseq == 1:
                        if 'df1' in vars():
                            df1 = df1.append(dd)
                        else:
                            df1 = dd
                    elif iseq == 2:
                        if 'df2' in vars():
                            df2 = df2.append(dd)
                        else:
                            df2 = dd
                    elif iseq == 3:
                        if 'df3' in vars():
                            df3 = df3.append(dd)
                        else:
                            df3 = dd
                    elif iseq == 4:
                        if 'df4' in vars():
                            df4 = df4.append(dd)
                        else:
                            df4 = dd
                    elif iseq == 5:
                        if 'df5' in vars():
                            df5 = df5.append(dd)
                        else:
                            df5 = dd
                    elif iseq == 6:
                        if 'df6' in vars():
                            df6 = df6.append(dd)
                        else:
                            df6 = dd
                    elif iseq == 7:
                        if 'df7' in vars():
                            df7 = df7.append(dd)
                        else:
                            df7 = dd
                    elif iseq == 8:
                        if 'df8' in vars():
                            df8 = df8.append(dd)
                        else:
                            df8 = dd
                    elif iseq == 9:
                        if 'df9' in vars():
                            df9 = df9.append(dd)
                        else:
                            df9 = dd
                    elif iseq == 10:
                        if 'df10' in vars():
                            df10 = df10.append(dd)
                        else:
                            df10 = dd
                    elif iseq == 11:
                        if 'df11' in vars():
                            df11 = df11.append(dd)
                        else:
                            df11 = dd
                    elif iseq == 12:
                        if 'df12' in vars():
                            df12 = df12.append(dd)
                        else:
                            df12 = dd
                    elif iseq == 13:
                        if 'df13' in vars():
                            df13 = df13.append(dd)
                        else:
                            df13 = dd

                    if 'df0' in vars():
                        df_box.append(df0)
                    if 'df1' in vars():
                        df_box.append(df1)
                    if 'df2' in vars():
                        df_box.append(df2)
                    if 'df3' in vars():
                        df_box.append(df3)
                    if 'df4' in vars():
                        df_box.append(df4)
                    if 'df5' in vars():
                        df_box.append(df5)
                    if 'df6' in vars():
                        df_box.append(df6)
                    if 'df7' in vars():
                        df_box.append(df7)
                    if 'df8' in vars():
                        df_box.append(df8)
                    if 'df9' in vars():
                        df_box.append(df9)
                    if 'df10' in vars():
                        df_box.append(df10)
                    if 'df11' in vars():
                        df_box.append(df11)
                    if 'df12' in vars():
                        df_box.append(df12)
                    if 'df13' in vars():
                        df_box.append(df13)
                else:
                    ddd = self.read_table_with_policy(file["filename"], item)
                    filename=file["filename"]
                    for dd in ddd:
                        print("数据预处理:")
                        print(dd.head(5).to_markdown())
                        iseq, dd = self.df_fit_template( df_shop, item,filename, dd, "", "")
                        print("sheet{} 记录数:".format(iseq+1),dd.shape[0])

                        if iseq == 0:
                            if 'df0' in vars():
                                df0 = df0.append(dd)
                            else:
                                df0 = dd
                        elif iseq == 1:
                            if 'df1' in vars():
                                df1 = df1.append(dd)
                            else:
                                df1 = dd
                        elif iseq == 2:
                            if 'df2' in vars():
                                df2 = df2.append(dd)
                            else:
                                df2 = dd
                        elif iseq == 3:
                            if 'df3' in vars():
                                df3 = df3.append(dd)
                            else:
                                df3 = dd
                        elif iseq == 4:
                            if 'df4' in vars():
                                df4 = df4.append(dd)
                            else:
                                df4 = dd
                        elif iseq == 5:
                            if 'df5' in vars():
                                df5 = df5.append(dd)
                            else:
                                df5 = dd
                        elif iseq == 6:
                            if 'df6' in vars():
                                df6 = df6.append(dd)
                            else:
                                df6 = dd
                        elif iseq == 7:
                            if 'df7' in vars():
                                df7 = df7.append(dd)
                            else:
                                df7 = dd
                        elif iseq == 8:
                            if 'df8' in vars():
                                df8 = df8.append(dd)
                            else:
                                df8 = dd
                        elif iseq == 9:
                            if 'df9' in vars():
                                df9 = df9.append(dd)
                            else:
                                df9 = dd
                        elif iseq == 10:
                            if 'df10' in vars():
                                df10 = df10.append(dd)
                            else:
                                df10 = dd
                        elif iseq == 11:
                            if 'df11' in vars():
                                df11 = df11.append(dd)
                            else:
                                df11 = dd
                        elif iseq == 12:
                            if 'df12' in vars():
                                df12 = df12.append(dd)
                            else:
                                df12 = dd
                        elif iseq == 13:
                            if 'df13' in vars():
                                df13 = df13.append(dd)
                            else:
                                df13 = dd

                        if 'df0' in vars():
                            df_box.append(df0)
                        if 'df1' in vars():
                            df_box.append(df1)
                        if 'df2' in vars():
                            df_box.append(df2)
                        if 'df3' in vars():
                            df_box.append(df3)
                        if 'df4' in vars():
                            df_box.append(df4)
                        if 'df5' in vars():
                            df_box.append(df5)
                        if 'df6' in vars():
                            df_box.append(df6)
                        if 'df7' in vars():
                            df_box.append(df7)
                        if 'df8' in vars():
                            df_box.append(df8)
                        if 'df9' in vars():
                            df_box.append(df9)
                        if 'df10' in vars():
                            df_box.append(df10)
                        if 'df11' in vars():
                            df_box.append(df11)
                        if 'df12' in vars():
                            df_box.append(df12)
                        if 'df13' in vars():
                            df_box.append(df13)

                print("进度表：  {}/{} [sheet {}/{}]  文件:{} ，行数{}".format(index + 1, files_count, k,
                                                                      len(policy_list), filename,
                                                                      dd.shape[0]))


                # # 如果不是空表
                #     if dd.shape[0] > 0:
                #         dd["filename"] = file["filename"]
                #         dd["sheetname"] = sheetname
                #         dd["policy"] = plat
                #
                #         if "店铺名称" in dd.columns:
                #             print("更新店铺名称")
                #             dd["店铺名称"].fillna("",inplace=True)
                #             shop_name=get_shopname(df_shop,file["filename"])[1]
                #             platform=get_shopname(df_shop,file["filename"])[0]
                #             print("更新店铺名称:",shop_name)
                #             # dd[dd["店铺名称"].str.len()==0]["店铺名称"]= get_shopname(df_shop,file["filename"])[1]
                #
                #             if len(shop_name.strip())>0:
                #                 dd["店铺名称"]= shop_name
                #                 if "平台" in dd.columns:
                #                     dd["平台"]= platform
                #
                #
                #         print("进度表：  {}/{} [sheet {}/{}]  文件:{} ，行数{}".format(index + 1, df_files.shape[0], k,
                #                                                               len(policy_list), file["filename"],
                #                                                               dd.shape[0]))
                #
                #         # 如果行数为0，则忽略，不需要关心格式
                #         if dd.shape[0] == 0:
                #             print("忽略空表：", file["filename"])
                #         elif file["filename"].find("错误") > 0:
                #             print("忽略错误的表格：", file["filename"])
                #         else:
                #             # if dd.shape[0] > 0:
                #             # 将列名转换成list
                #             if_exist = False
                #             iseq = 0
                #             dd_columns = dd.columns.to_list()
                #             for r in self.file_columns_list:
                #                 # 如果相同，已经存在
                #                 # print("对比1:",dd_columns)
                #                 # print("对比2:",r[1:])
                #
                #                 if operator.eq(dd_columns, r[1:]):
                #                     # if dd_columns==r[1:]:
                #                     if_exist = True
                #                     break
                #                 iseq = iseq + 1
                #
                #             # 模板已经存在
                #             if if_exist:
                #                 # print("模板已经存在！：", file["filename"])
                #                 # print("模板1:", dd_columns)
                #                 # print("模板2:",iseq)
                #                 # for t in file_columns_list:
                #                 #     print(t)
                #
                #                 if iseq == 0:
                #                     if 'df0' in vars():
                #                         df0 = df0.append(dd)
                #                     else:
                #                         df0 = dd
                #                 elif iseq == 1:
                #                     if 'df1' in vars():
                #                         df1 = df1.append(dd)
                #                     else:
                #                         df1 = dd
                #                 elif iseq == 2:
                #                     if 'df2' in vars():
                #                         df2 = df2.append(dd)
                #                     else:
                #                         df2 = dd
                #                 elif iseq == 3:
                #                     if 'df3' in vars():
                #                         df3 = df3.append(dd)
                #                     else:
                #                         df3 = dd
                #                 elif iseq == 4:
                #                     if 'df4' in vars():
                #                         df4 = df4.append(dd)
                #                     else:
                #                         df4 = dd
                #                 elif iseq == 5:
                #                     if 'df5' in vars():
                #                         df5 = df5.append(dd)
                #                     else:
                #                         df5 = dd
                #                 elif iseq == 6:
                #                     if 'df6' in vars():
                #                         df6 = df6.append(dd)
                #                     else:
                #                         df6 = dd
                #                 elif iseq == 7:
                #                     if 'df7' in vars():
                #                         df7 = df7.append(dd)
                #                     else:
                #                         df7 = dd
                #                 elif iseq == 8:
                #                     if 'df8' in vars():
                #                         df8 = df8.append(dd)
                #                     else:
                #                         df8 = dd
                #                 elif iseq == 9:
                #                     if 'df9' in vars():
                #                         df9 = df9.append(dd)
                #                     else:
                #                         df9 = dd
                #                 elif iseq == 10:
                #                     if 'df10' in vars():
                #                         df10 = df10.append(dd)
                #                     else:
                #                         df10 = dd
                #                 elif iseq == 11:
                #                     if 'df11' in vars():
                #                         df11 = df11.append(dd)
                #                     else:
                #                         df11 = dd
                #                 elif iseq == 12:
                #                     if 'df12' in vars():
                #                         df12 = df12.append(dd)
                #                     else:
                #                         df12 = dd
                #                 elif iseq == 13:
                #                     if 'df13' in vars():
                #                         df13 = df13.append(dd)
                #                     else:
                #                         df13 = dd
                #
                #
                #             else:
                #                 #     print("新增字段模板：",file["filename"])
                #                 #     print("模板1:", dd_columns)
                #                 #     print("模板2:")
                #                 #     for t in file_columns_list:
                #                 #         print(t)
                #
                #                 print("新增模板：", dd_columns)
                #
                #                 # 从第0列插入字段
                #                 dd_columns.insert(0, file["filename"])
                #                 # 从第1列插入字段
                #                 # dd_columns.insert(1,sheetname)
                #
                #                 self.file_columns_list.append(dd_columns)
                #
                #                 iseq = len(self.file_columns_list) - 1
                #                 if iseq == 0:
                #                     df0 = dd
                #                 elif iseq == 1:
                #                     df1 = dd
                #                 elif iseq == 2:
                #                     df2 = dd
                #                 elif iseq == 3:
                #                     df3 = dd
                #                 elif iseq == 4:
                #                     df4 = dd
                #                 elif iseq == 5:
                #                     df5 = dd
                #                 elif iseq == 6:
                #                     df6 = dd
                #                 elif iseq == 7:
                #                     df7 = dd
                #                 elif iseq == 8:
                #                     df8 = dd
                #                 elif iseq == 9:
                #                     df9 = dd
                #                 elif iseq == 10:
                #                     df10 = dd
                #                 elif iseq == 11:
                #                     df11 = dd
                #                 elif iseq == 12:
                #                     df12 = dd
                #                 elif iseq == 13:
                #                     df13 = dd

                                # print(file["filename"], dd.shape[0])
                                # print("进度表：{}/{}  文件{}，行数{}".format(index + 1, df_files.shape[0], file["filename"], dd.shape[0]))



        print("最终的字段列表：")
        print(self.file_columns_list)
        print("字段模板共有：{} 个".format(len(self.file_columns_list)))

        return df_box

    def combine_table(self,filedir,keyword):
        # df_box = self.get_table_box(filepath,keyword)
        print('请选择要合并的文件目录（请确定所有excel的表头都是相同的哦！）:')

        if len(filedir) > 0:
            pass
        else:
            filedir = filedialog.askdirectory()  # 获取文件夹
            print("你选择的路径是：", filedir)

        if len(filedir) == 0:
            print("你没有输入任何目录 :(")
            sys.exit()
            return

        global default_dir
        default_dir = filedir

        if len(keyword) > 0:
            pass
        else:
            # print('请输入要合并的文件名称或者后缀匹配符:（比如都必须包含 序时账 三个字，那么就请输入  "序时账" ， 不输入就表示所有excel都要合并！）')
            print(
                '筛选文件的规则:  \r\n1、京东 csv 表示选择文件完整路径中包含 "京东"和"csv"的文件  \r\n2、比如 淘宝 !天猫  表示只要淘宝，不要天猫  \r\n3、淘宝|天猫 表示 包含淘宝或者天猫    \r\n4、空格中间是and关系，每个项目都支持正则表达式 比如：2019(.*)海外旗舰  \r\n4、什么都不输入，表示默认选择目录下所有文件! \r\n请输入:')
            keyword = input()

            if len(filedir) == 0:
                print("你没有输入任何关键词 :(")
                keyword = ''
                # sys.exit()
                # return

        print("你希望在'{}'目录下找到所有的包含“{}”文件，然后合并。".format(filedir, keyword))

        df_box = self.read_files_table(filedir, keyword)

        # df0,df1,df2,df3,df4,df5,df6,df7,df8,df9,df10,df11,df12,df13
        index = 0
        for df in df_box:
            print("第{}个表格,记录数:{}".format(index, df.shape[0]))
            print(df.head(10).to_markdown())
            # df.to_excel(r"work/合并表格_test.xlsx")

            # plat=df_box.head(1)["policy"]
            # plat = df.iloc[0]["policy"]
            # 取指定列的第一行数据
            plat = df["policy"].iloc[0]
            # plat="".join(plat)
            print("plat:", plat)

            # 每张表格最大的行数
            pagecount=500000
            pagecnt= int(df.shape[0] / pagecount) + 1
            for i in range(0, int(df.shape[0] / pagecount) + 1):
                # print("分页：{}  from:{} to:{}".format(i+1, i * 500000, (i + 1) * 500000))
                print("分页：{}  from:{} to:{} ，记录数: {} ".format(i+1, i * pagecount, (i + 1) * pagecount,df[i*pagecount:(i+1)*pagecount].shape[0]))
                # df.iloc[i*500000:(i+1)*500000].to_csv(default_dir + "\{}_合并表格_{}.csv".format(index,i))
                # df.iloc[i * 500000:(i + 1) * 500000].to_excel(default_dir + "\{}_合并表格_{}.xlsx".format(index, i))
                # df[i*500000:(i+1)*500000].to_csv( "work\{}_合并表格_{}.csv".format(index,i))
                print("合并生成:",filedir+os.sep+"{}_合并表格_{}.{}.{}.xlsx".format(plat,index,pagecnt,i+1) )
                df[i*pagecount:(i+1)*pagecount].to_excel( filedir+os.sep+"{}_合并表格_{}.{}.{}.xlsx".format(plat,index,pagecnt,i+1))

            index = index + 1

        # print("生成完毕，现在关闭吗？yes/no")
        # byebye = input()
        # print('bybye:', byebye)


    def search_text(self,filedir,keyword,searchcolumn,searchword):
        # df_box = self.get_table_box(filepath,keyword)
        print('请选择要合并的文件目录（请确定所有excel的表头都是相同的哦！）:')

        if len(filedir) > 0:
            pass
        else:
            filedir = filedialog.askdirectory()  # 获取文件夹
            print("你选择的路径是：", filedir)

        if len(filedir) == 0:
            print("你没有输入任何目录 :(")
            sys.exit()
            return

        global default_dir
        default_dir = filedir

        if len(keyword) > 0:
            pass
        else:
            # print('请输入要合并的文件名称或者后缀匹配符:（比如都必须包含 序时账 三个字，那么就请输入  "序时账" ， 不输入就表示所有excel都要合并！）')
            print(
                '筛选文件的规则:  \r\n1、京东 csv 表示选择文件完整路径中包含 "京东"和"csv"的文件  \r\n2、比如 淘宝 !天猫  表示只要淘宝，不要天猫  \r\n3、淘宝|天猫 表示 包含淘宝或者天猫    \r\n4、空格中间是and关系，每个项目都支持正则表达式 比如：2019(.*)海外旗舰  \r\n4、什么都不输入，表示默认选择目录下所有文件! \r\n请输入:')
            keyword = input()

            if len(filedir) == 0:
                print("你没有输入任何关键词 :(")
                keyword = ''
                # sys.exit()
                # return

        print("你希望在'{}'目录下找到所有的包含“{}”文件，然后合并。".format(filedir, keyword))

        # df_box = self.read_files_table(filedir, keyword)
        df_box = self.read_files_table_byword(filedir, keyword,searchcolumn,searchword)

        # df0,df1,df2,df3,df4,df5,df6,df7,df8,df9,df10,df11,df12,df13
        index = 0
        for df in df_box:
            print("第{}个表格,记录数:{}".format(index, df.shape[0]))
            print(df.head(10).to_markdown())
            # df.to_excel(r"work/合并表格_test.xlsx")

            # plat=df_box.head(1)["policy"]
            # plat = df.iloc[0]["policy"]
            # 取指定列的第一行数据
            plat = df["policy"].iloc[0]
            # plat="".join(plat)
            print("plat:", plat)

            pagecnt = int(df.shape[0] / 500000) + 1
            for i in range(0, int(df.shape[0] / 500000) + 1):
                # print("分页：{}  from:{} to:{}".format(i+1, i * 500000, (i + 1) * 500000))
                print("分页：{}  from:{} to:{} ，记录数: {} ".format(i + 1, i * 500000, (i + 1) * 500000,
                                                              df[i * 500000:(i + 1) * 500000].shape[0]))
                # df.iloc[i*500000:(i+1)*500000].to_csv(default_dir + "\{}_合并表格_{}.csv".format(index,i))
                # df.iloc[i * 500000:(i + 1) * 500000].to_excel(default_dir + "\{}_合并表格_{}.xlsx".format(index, i))
                # df[i*500000:(i+1)*500000].to_csv( "work\{}_合并表格_{}.csv".format(index,i))
                print("合并生成:", filedir + os.sep + "{}_合并表格_{}.{}.{}.xlsx".format(plat, index, pagecnt, i + 1))
                df[i * 500000:(i + 1) * 500000].to_excel(
                    filedir + os.sep + "{}_合并表格_{}.{}.{}.xlsx".format(plat, index, pagecnt, i + 1))

            index = index + 1

    def call_opt(self,filename):
        # 读取 csv 转 excel

        # df = pd.read_csv(filename)
        # try:
        #     df = pd.read_csv(filename, encoding="gbk", dtype=str,
        #                           on_bad_lines='skip').reset_index()  # ,decodeing="utf-8"
        #     # temp_df = pd.read_csv(filename,encoding="gb18030", dtype=str,error_bad_lines=False).reset_index()  # ,decodeing="utf-8"
        # except Exception as  err:
        #     # print(filename, " 异常:", err)
        #     # print(filename, "是空表")
        #     try:
        #         df = pd.read_csv(filename, encoding="gbk", dtype=str,
        #                               error_bad_lines=False, engine="python").reset_index()
        #     except Exception as  err:
        #         print(filename, " 异常:", err)
        #         # print(filename, "是空表")
        #
        #         # return pd.DataFrame(columns=["iid"]).head(0)
        #
        #
        # excel_filename = filename.replace(".csv", ".xlsx")
        # print("生成文件:", excel_filename)
        # df.to_excel(excel_filename, index=False)
        #
        # return

        policy_list = self.get_policy(filename)
        for item in policy_list:
            # 循环匹配每一个sheet
            # [plat, sheetname, title, ignoretop, skiptop, skipbottom, title_key, bottom_key, _valid_columns,
            #  _new_columns, _rename, _default_columns])

            platform, sheetname, title,ignoretop, skiptop,skipbottom,title_key, bottom_key, valid_columns, new_columns, _rename, default_columns = item[0], item[1], item[2], item[
                3], item[4] ,item[5],item[6],item[7],item[8],item[9],item[10],item[11],item[12]

            if ignoretop == "1":
                ignoretop = True
            else:
                ignoretop = False

            print("读取文件:", filename)
            # k = k + 1
            df = self.read_table_with_policy(filename, item)
            # df=pd.read_csv(filename)
            # 如果不是空表
            if df.shape[0] > 0:
                # 读取到数据 ，  以下内容可以自定义
                excel_filename = filename.replace(".csv", ".xlsx")
                if filename.find("海外") < 0:
                    df.to_excel(excel_filename, sheet_name="月账单", index=False)
                else:
                    df.to_excel(excel_filename, index=False)
                print("生成文件:", excel_filename)


    def call_opt_reduce(self,filename):
        #  xlsx 转 xls
        if filename.find(".xlsx")>0:
            df =  pd.read_excel(filename)
        elif filename.find(".csv")>0:
            df =  pd.read_csv(filename)
        else  :
            df = pd.read_excel(filename)

        excel_filename= filename.replace(".xlsx",".xls")
        df.to_excel(excel_filename, index=False)
        print("生成文件:", excel_filename)




    def opt_all_excel(self, rootdir, filekey, call_opt):
        # title_key,bottom_key
        df_files = self.get_files_df(rootdir, filekey)
        df_box = []
        for index, file in df_files.iterrows():
            call_opt(file["filename"])
            print("进度表：  {}/{}   文件:{} ".format(index + 1, df_files.shape[0], file["filename"]))


    def convert_excel(self, filedir, keyword):
        global default_dir

        default_dir = ""
        if len(filedir) == 0:
            print('请选择要操作的文件目录:')
            default_dir = filedialog.askdirectory()  # 获取文件夹
            print("你选择的路径是：", default_dir)
        else:
            default_dir = filedir

        if len(default_dir) == 0:
            print("你没有输入任何目录 :(")
            sys.exit()
            return

        if len(keyword) == 0:
            print(
                '筛选文件的规则:  \r\n1、京东 csv 表示选择文件完整路径中包含 "京东"和"csv"的文件  \r\n2、比如 淘宝 !天猫  表示只要淘宝，不要天猫  \r\n3、淘宝|天猫 表示 包含淘宝或者天猫    \r\n4、空格中间是and关系，每个项目都支持正则表达式 比如：2019(.*)海外旗舰  \r\n4、什么都不输入，表示默认选择目录下所有文件! \r\n请输入:')
            filekey = input()
        else:
            filekey = keyword

        if len(filekey) == 0:
            print("你没有输入任何关键词 :(")
            filekey = ''
            # sys.exit()
            # return

        print("你希望在'{}'目录下找到所有的  {}   文件，批量操作...".format(default_dir, filekey))

        self.opt_all_excel(default_dir, filekey, self.call_opt)

    def sum_bill_fn(self, filedir, keyword):
        global default_dir

        default_dir = ""
        if len(filedir) == 0:
            print('请选择要操作的文件目录:')
            default_dir = filedialog.askdirectory()  # 获取文件夹
            print("你选择的路径是：", default_dir)
        else:
            default_dir = filedir

        if len(default_dir) == 0:
            print("你没有输入任何目录 :(")
            sys.exit()
            return

        if len(keyword) == 0:
            print(
                '筛选文件的规则:  \r\n1、京东 csv 表示选择文件完整路径中包含 "京东"和"csv"的文件  \r\n2、比如 淘宝 !天猫  表示只要淘宝，不要天猫  \r\n3、淘宝|天猫 表示 包含淘宝或者天猫    \r\n4、空格中间是and关系，每个项目都支持正则表达式 比如：2019(.*)海外旗舰  \r\n4、什么都不输入，表示默认选择目录下所有文件! \r\n请输入:')
            filekey = input()
        else:
            filekey = keyword

        if len(filekey) == 0:
            print("你没有输入任何关键词 :(")
            filekey = ''
            # sys.exit()
            # return

        print("你希望在'{}'目录下找到所有的  {}   文件，批量操作...".format(default_dir, filekey))

        # df_sum_wjl=pd.DataFrame(columns=["公司主体", "平台", "店铺名称", "回款金额", "退款金额"],index=[0])
        # self.opt_all_excel(default_dir, filekey, self.call_opt)
        df_files = self.get_files_df(default_dir, filekey)
        df_box = []
        for index, file in df_files.iterrows():
            filename=file["filename"]
            if filename[-10:].find("订单")>0:
                # call_opt(file["filename"])
                df = pd.read_excel(filename)
                # print(df.dtypes)
                if "回款日期" in df.columns:
                    yearmonth=df["回款日期"].iloc[0]
                    year=yearmonth[0:4]
                    month=yearmonth[5:7]

                    print("filename",filename,year,month)
                    # print(df.head(3).to_markdown())
                    df_2 = df[["公司主体", "平台", "店铺名称", "回款金额", "退款金额"]].copy()
                    df_2["回款金额"].fillna(0, inplace=True)
                    df_2["退款金额"].fillna(0, inplace=True)

                    df_2["回款金额"] = df_2["回款金额"].astype("float64")
                    df_2["退款金额"] = df_2["退款金额"].astype("float64")

                    df_3 = df_2.groupby(["公司主体", "平台", "店铺名称"]).agg({"回款金额": np.sum, "退款金额": np.sum})
                    df_3 = pd.DataFrame(df_3).reset_index()
                    df_3.columns = ["主体", "平台", "店铺名称", "回款金额", "退款金额"]
                    df_3["year"]=year
                    df_3["month"]=int(month)

                    # print(df.head(3).to_markdown())
                    # df_sum_wjl.append(df_3)

                    if "df_sum_wjl" in vars():
                        df_sum_wjl=df_sum_wjl.append(df_3)
                    else:
                        df_sum_wjl=df_3.copy()

                else:
                    print("文件异常",filename,year,month)
                    print(df.head(3).to_markdown())
                    print(df.dtypes)

                # print(df_sum_wjl.to_markdown())

        print("进度表：  {}/{}   文件:{} ".format(index + 1, df_files.shape[0], file["filename"]))

        df_sum_wjl.to_excel(r"d:\kkkk.xls")


    def sum_it_billmatchorder(self, filedir, keyword):
        # 读取 账单匹配订单
        global default_dir

        default_dir = ""
        if len(filedir) == 0:
            print('请选择要操作的文件目录:')
            default_dir = filedialog.askdirectory()  # 获取文件夹
            print("你选择的路径是：", default_dir)
        else:
            default_dir = filedir

        if len(default_dir) == 0:
            print("你没有输入任何目录 :(")
            sys.exit()
            return

        if len(keyword) == 0:
            print(
                '筛选文件的规则:  \r\n1、京东 csv 表示选择文件完整路径中包含 "京东"和"csv"的文件  \r\n2、比如 淘宝 !天猫  表示只要淘宝，不要天猫  \r\n3、淘宝|天猫 表示 包含淘宝或者天猫    \r\n4、空格中间是and关系，每个项目都支持正则表达式 比如：2019(.*)海外旗舰  \r\n4、什么都不输入，表示默认选择目录下所有文件! \r\n请输入:')
            filekey = input()
        else:
            filekey = keyword

        if len(filekey) == 0:
            print("你没有输入任何关键词 :(")
            filekey = ''
            # sys.exit()
            # return

        print("你希望在'{}'目录下找到所有的  {}   文件，批量操作...".format(default_dir, filekey))

        # df_sum_wjl=pd.DataFrame(columns=["公司主体", "平台", "店铺名称", "回款金额", "退款金额"],index=[0])
        # self.opt_all_excel(default_dir, filekey, self.call_opt)
        df_files = self.get_files_df(default_dir, filekey)
        df_box = []
        for index, file in df_files.iterrows():
            filename=file["filename"]
            print("判断文件 ",filename)
            # if filename[-30:].find("回款总表")>0:   # 文件名后10位含有 订单字样
            if True:
                print("找到回款总表",filename)
                # print(filename[-10:].find("回款总表"))
                # call_opt(file["filename"])
                df = pd.read_excel(filename)
                # print(df.dtypes)
                print(df.head(10).to_markdown())
                if df.shape[0]>0:
                    if "回款时间" in df.columns:
                        try:
                            yearmonth = df["回款时间"].iloc[0]
                            year = yearmonth[0:4]
                            month = yearmonth[5:7]

                            print("filename", filename, year, month)
                            # print(df.head(3).to_markdown())
                            df_2 = df[["主体", "平台名称", "店铺名称", "回款金额", "退款金额","回款时间"]].copy()
                            df_2["回款金额"].fillna(0, inplace=True)
                            df_2["退款金额"].fillna(0, inplace=True)

                            df_2["回款金额"] = df_2["回款金额"].astype("float64")
                            df_2["退款金额"] = df_2["退款金额"].astype("float64")

                            df_2.columns = ["主体", "平台", "店铺名称", "回款金额", "退款金额","回款时间"]
                            df_2["year"] = df_2["回款时间"].apply(lambda x: int(x.split('-')[0]))
                            df_2["month"] =df_2["回款时间"].apply(lambda x: int(x.split('-')[1]))

                            # df_3 = df_2.groupby(["主体", "平台名称", "店铺名称"]).agg({"回款金额": np.sum, "退款金额": np.sum})
                            # df_3 = pd.DataFrame(df_3).reset_index()
                            # df_3.columns = ["主体", "平台", "店铺名称", "回款金额", "退款金额"]
                            # df_3["year"] = year
                            # df_3["month"] = int(month)

                            # print(df.head(3).to_markdown())
                            # df_sum_wjl.append(df_3)

                            if "df_sum_wjl" in vars():
                                df_sum_wjl = df_sum_wjl.append(df_2)
                            else:
                                df_sum_wjl = df_2.copy()

                        except Exception as  e:
                                print("文件读取错误：" + str(e))

                    else:
                        print("文件异常",filename,year,month)
                        print(df.head(3).to_markdown())
                        print(df.dtypes)
            else:
                print("没有找到回款总表", filename)
                # print(df_sum_wjl.to_markdown())

        print("进度表：  {}/{}   文件:{} ".format(index + 1, df_files.shape[0], file["filename"]))

        if "df_sum_wjl" in vars():
            df_sum_wjl.to_excel(r"C:\Users\mega\Desktop\生成总表\IT口径店铺账单匹配订单.xlsx")

        return  df_sum_wjl

    def sum_it_billsubject(self, filedir, keyword):
        # 读取 账单主体表
        global default_dir
        default_dir = ""
        if len(filedir) == 0:
            # print('请选择要操作的文件目录:')
            default_dir = filedialog.askdirectory()  # 获取文件夹
            # print("你选择的路径是：", default_dir)
        else:
            default_dir = filedir

        if len(default_dir) == 0:
            print("你没有输入任何目录 :(")
            sys.exit()
            return

        if len(keyword) == 0:
            print(
                '筛选文件的规则:  \r\n1、京东 csv 表示选择文件完整路径中包含 "京东"和"csv"的文件  \r\n2、比如 淘宝 !天猫  表示只要淘宝，不要天猫  \r\n3、淘宝|天猫 表示 包含淘宝或者天猫    \r\n4、空格中间是and关系，每个项目都支持正则表达式 比如：2019(.*)海外旗舰  \r\n4、什么都不输入，表示默认选择目录下所有文件! \r\n请输入:')
            filekey = input()
        else:
            filekey = keyword

        if len(filekey) == 0:
            print("你没有输入任何关键词 :(")
            filekey = ''
            # sys.exit()
            # return
        print("你希望在'{}'目录下找到所有的  {}   文件，批量操作...".format(default_dir, filekey))

        df_files = self.get_files_df(default_dir, filekey)
        df_box = []
        for index, file in df_files.iterrows():
            filename = file["filename"]
            print("判断文件 ", filename)
            if (filename.find("账单主体")>0) & (filename.find("回款日期")>0):
            # if True:
                print("找到账单主体",filename)
                df = pd.read_excel(filename)
                # print(df.head(10).to_markdown())
                if df.shape[0]>0:
                    if "总回款" in df.columns:
                        try:
                            df_2 = df[["平台", "账单主体", "账单店铺" , "总回款", "总退款", "支付日期"]].copy()
                            df_2["总回款"].fillna(0, inplace=True)
                            df_2["总退款"].fillna(0, inplace=True)
                            df_2["总回款"] = df_2["总回款"].astype("float64")
                            df_2["总退款"] = df_2["总退款"].astype("float64")
                            df_2.columns = ["平台", "账单主体", "账单店铺", "总回款", "总退款", "支付日期"]
                            if "df_sum_wjl" in vars():
                                df_sum_wjl = df_sum_wjl.append(df_2)
                            else:
                                df_sum_wjl = df_2.copy()
                        except Exception as  e:
                                print("文件读取错误：" + str(e))
                    else:
                        print("文件异常",filename,year,month)
                        print(df.head(3).to_markdown())
                        print(df.dtypes)
            else:
                # pass
                print("没有找到账单主体表", filename)
                # print(df_sum_wjl.to_markdown())
        print("进度表：  {}/{}   文件:{} ".format(index + 1, df_files.shape[0], file["filename"]))
        #
        # if "df_sum_wjl" in vars():
        #     df_sum_wjl.to_excel(r"C:\Users\mega\Desktop\生成总表\订单主体.xlsx")
        return  df_sum_wjl

    def sum_bill_it(self, filedir, keyword):
        # 读取 账单总表
        global default_dir

        default_dir = ""
        if len(filedir) == 0:
            print('请选择要操作的文件目录:')
            default_dir = filedialog.askdirectory()  # 获取文件夹
            print("你选择的路径是：", default_dir)
        else:
            default_dir = filedir

        if len(default_dir) == 0:
            print("你没有输入任何目录 :(")
            sys.exit()
            return

        if len(keyword) == 0:
            print(
                '筛选文件的规则:  \r\n1、京东 csv 表示选择文件完整路径中包含 "京东"和"csv"的文件  \r\n2、比如 淘宝 !天猫  表示只要淘宝，不要天猫  \r\n3、淘宝|天猫 表示 包含淘宝或者天猫    \r\n4、空格中间是and关系，每个项目都支持正则表达式 比如：2019(.*)海外旗舰  \r\n4、什么都不输入，表示默认选择目录下所有文件! \r\n请输入:')
            filekey = input()
        else:
            filekey = keyword

        if len(filekey) == 0:
            print("你没有输入任何关键词 :(")
            filekey = ''
            # sys.exit()
            # return

        print("你希望在'{}'目录下找到所有的  {}   文件，批量操作...".format(default_dir, filekey))

        # df_sum_wjl=pd.DataFrame(columns=["公司主体", "平台", "店铺名称", "回款金额", "退款金额"],index=[0])
        # self.opt_all_excel(default_dir, filekey, self.call_opt)
        df_files = self.get_files_df(default_dir, filekey)
        df_box = []
        for index, file in df_files.iterrows():
            filename=file["filename"]
            # if filename[-10:].find("账单总表")>0:   # 文件名后10位含有 订单字样
            if True:
                # call_opt(file["filename"])
                df = pd.read_excel(filename)  # ,skiprows=1
                # print(df.dtypes)
                print("抽查2:",filename)
                print(df.head(10).to_markdown())
                if df.shape[0]>0:
                    # if "回款日期" in df.columns:
                    if True:
                        try:
                            # yearmonth = df["回款日期"].iloc[0]
                            # year = yearmonth[0:4]
                            # month = yearmonth[5:7]
                            # 摘要	公司	平台	店铺	交易方式  收入	支出	收入	支出	收入	支出	收入	支出	收入	支出	收入	支出	收入	支出	收入	支出	收入	支出	收入	支出	收入	支出	收入	支出	收入	支出
                            # df.columns=["摘要","公司","平台","店铺","交易方式","收入1","支出1","收入2","支出2","收入3","支出3","收入4","支出4","收入5","支出5","收入6","支出6","收入7","支出7","收入8","支出8","收入9","支出9","收入10","支出10","收入11","支出11","收入12","支出12","收入合计","支出合计" ]
                            # df.rename(columns={"收入":"收入1","支出":"支出1",
                            #            "收入.1":"收入2","支出.1":"支出2",
                            #            "收入.2":"收入3","支出.2":"支出3",
                            #            "收入.3":"收入4","支出.3":"支出4",
                            #            "收入.4":"收入5","支出.4":"支出5",
                            #            "收入.5":"收入6","支出.5":"支出6",
                            #            "收入.6":"收入7","支出.6":"支出7",
                            #            "收入.7":"收入8","支出.7":"支出8",
                            #            "收入.8":"收入9","支出.8":"支出9",
                            #            "收入.9":"收入10","支出.9":"支出10",
                            #            "收入.10":"收入11","支出.10":"支出11",
                            #            "收入.11":"收入12","支出.11":"支出12",
                            #            "收入.12":"收入合计","支出.12":"支出合计"},inplace=True)

                            df.rename(columns={"1月": "收入1", "1月.1": "支出1",
                                               "2月": "收入2", "2月.1": "支出2",
                                               "3月": "收入3", "3月.1": "支出3",
                                               "4月": "收入4", "4月.1": "支出4",
                                               "5月": "收入5", "5月.1": "支出5",
                                               "6月": "收入6", "6月.1": "支出6",
                                               "7月": "收入7", "7月.1": "支出7",
                                               "8月": "收入8", "8月.1": "支出8",
                                               "9月": "收入9", "9月.1": "支出9",
                                               "10月": "收入10", "10月.1": "支出10",
                                               "11月": "收入11", "11月.1": "支出11",
                                               "12月": "收入12", "12月.1": "支出12",
                                               "汇总": "收入合计", "汇总.1": "支出合计"}, inplace=True)


                            df=df[1:]
                            print("test1111:")
                            print(df.head(10).to_markdown())

                            # df_sum_wjl = df[["主体", "平台", "店铺", "收入1", "支出1"]]
                            # df_sum_wjl = ["主体", "平台", "店铺", "回款金额", "退款金额"]
                            # df_sum_wjl["month"]=1
                            for i in range(1,13):
                                if "收入{}".format(i) in df.columns:
                                    df_temp=df[["摘要", "公司", "平台", "店铺", "收入{}".format(i), "支出{}".format(i)]]
                                    df_temp.columns=["摘要", "主体", "平台", "店铺名称", "回款金额", "退款金额"]
                                    df_temp["month"] = i

                                    print("test2:")
                                    print(df_temp.head(10).to_markdown())

                                    if "df_sum_wjl" in vars():
                                        df_sum_wjl = df_sum_wjl.append(df_temp)
                                    else:
                                        df_sum_wjl = df_temp.copy()


                        except Exception as  e:
                                print("文件读取错误：" + str(e))

                    else:
                        print("文件异常",filename,year,month)
                        print(df.head(3).to_markdown())
                        print(df.dtypes)

                # print(df_sum_wjl.to_markdown())

        print("进度表：  {}/{}   文件:{} ".format(index + 1, df_files.shape[0], file["filename"]))

        df_sum_wjl.columns = ["摘要","主体", "平台名称", "店铺名称", "回款金额", "退款金额", "month"]
        # 根据摘要删除不参与统计的记录
        print("debug_1")
        print(df_sum_wjl.head(10).to_markdown())
        df_sum_wjl["摘要"]=df_sum_wjl["摘要"].astype(str)
        df_sum_wjl["回款金额"]=df_sum_wjl.apply(lambda x: 0 if ( x["摘要"].find("佣金")>0  ) else x["回款金额"]  ,axis=1)
        print("debug_2")
        print(df_sum_wjl.head(10).to_markdown())
        df_sum_wjl["回款金额"]=df_sum_wjl.apply(lambda x: 0 if (   x["摘要"].find("随单送的京豆")>0 ) else x["回款金额"]  ,axis=1)
        print("debug_3")
        print(df_sum_wjl.head(10).to_markdown())
        df_sum_wjl = df_sum_wjl.groupby(["主体", "平台名称", "店铺名称", "month"]).agg({"回款金额": np.sum, "退款金额": np.sum})
        df_sum_wjl = pd.DataFrame(df_sum_wjl).reset_index()
        df_sum_wjl.columns = ["主体", "平台", "店铺名称", "month", "回款小计", "退款小计"]
        print("debug_4")
        print(df_sum_wjl.head(10).to_markdown())



        df_sum_wjl.to_excel(r"C:\Users\mega\Desktop\生成总表\IT口径店铺账单收入.xlsx")
        return  df_sum_wjl


if __name__ == "__main__":
    print("执行开始:", time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()))
    start = time.time()
    # 汇总比对财务和IT账单匹配订单统计结果
    policy = {
    "默认":{
        "filekey":"",
        "sheets":[
            {
               "sheetname":"",
               "title":"",
                "ignoretop":"1",
                "skiptop": "1",
                "skipbottom": "",
                "title_key":"数量",
                "bottom_key":"合计",
                "columns":"",
                "newcolumns":"",
                "rename":[],
                "default_columns":[]
            }
        ]
    }}

    fasttable=Fasttable()
    # 设置读取规则
    fasttable.set_policy(policy)
    # 读取csv，转excel
    # fasttable.convert_excel("/Users/mac/Downloads/未命名文件夹 2", ".csv  海外")

    # 读取账单主体_回款金额表数据
    it_path=r"Y:\it审计处理需求\OMS导出数据\2020导出数据(账单匹配订单2022-1-12)18pm"
    df_it_dianpu = fasttable.sum_it_billsubject(it_path, "回款日期")
    # df_it_dianpu["支付日期"].fillna("1999",inplace=True)
    df_it_dianpu["支付日期"] = df_it_dianpu["支付日期"].astype(str)
    df_it_dianpu["2019年发出商品金额"] = df_it_dianpu.apply(lambda x: x["总回款"] if x["支付日期"].find("2019")>=0 else 0, axis=1)
    df_it_dianpu = df_it_dianpu.groupby([ "账单主体", "平台", "账单店铺"]).agg({"总回款": np.sum, "总退款": np.sum, "2019年发出商品金额":np.sum})
    df_it_dianpu = pd.DataFrame(df_it_dianpu).reset_index()
    df_it_dianpu.columns = [ "主体", "平台", "店铺名称" , "回款金额_IT", "退款金额_IT","2019年发出商品金额"]



    df_it_dianpu.to_excel(r"C:\Users\mega\Desktop\生成总表\账单主体回款时间汇总.xlsx")

    # 读取财务汇总表数据
    df_fn = pd.read_excel(r"C:\Users\mega\Desktop\财务总数据\汇总2019-2021店铺收入-税率4.xlsx", sheet_name="2020年")
    df_fn = df_fn[["主体", "平台", "OMS店铺名称",  "回款小计", "退款小计"]]
    df_fn["平台"] = df_fn["平台"].astype(str)
    df_fn["平台"] = df_fn["平台"].apply(lambda x: x.replace("货到付款", "").strip())
    df_fn = pd.DataFrame(df_fn.groupby(["主体", "平台", "OMS店铺名称"]).agg({ "回款小计": np.sum, "退款小计": np.sum})).reset_index()
    df_fn.columns = ["主体", "平台", "店铺名称", "回款金额_财务", "退款金额_财务"]
    df_fn.to_excel(r"C:\Users\mega\Desktop\生成总表\财务数据汇总.xlsx")


    print("财务匹配IT")
    df_fn_more = df_fn.merge(df_it_dianpu, how="left", on=["主体", "平台", "店铺名称"])
    df_fn_more["回款金额_IT"].fillna(0, inplace=True)
    df_fn_more["退款金额_IT"].fillna(0, inplace=True)
    df_fn_more["回款金额_财务"].fillna(0, inplace=True)
    df_fn_more["退款金额_财务"].fillna(0, inplace=True)
    df_fn_more["2019年发出商品金额"].fillna(0, inplace=True)
    # print(df_fn_more.head(10).to_markdown())
    df_fn_more["回款金额_差异"] = df_fn_more["回款金额_财务"] - df_fn_more["回款金额_IT"]
    df_fn_more["退款金额_差异"] = df_fn_more["退款金额_财务"] - df_fn_more["退款金额_IT"]
    # df_fn_more["实际回款金额_财务"] = df_fn_more["回款金额_财务"] + df_fn_more["退款金额_财务"]
    # df_fn_more["实际回款金额_IT"] = df_fn_more["回款金额_IT"] + df_fn_more["退款金额_IT"]
    df_fn_more["收入差异"] = df_fn_more["回款金额_差异"] + df_fn_more["退款金额_差异"]
    df_fn_more["回款金额_差异"] = round(df_fn_more["回款金额_差异"], 2)
    df_fn_more["退款金额_差异"] = round(df_fn_more["退款金额_差异"], 2)
    df_fn_more["收入差异"] = round(df_fn_more["收入差异"], 2)

    df_fn_more[["主体", "平台", "店铺名称", "回款金额_财务", "退款金额_财务", "回款金额_IT", "退款金额_IT", "2019年发出商品金额", "回款金额_差异", "退款金额_差异", "收入差异"]]\
        .to_excel(r"C:\Users\mega\Desktop\生成总表\2020财务匹配IT_11号19点.xlsx")



    print("执行正确，结束时间:", time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()))
    end = time.time()
    print("执行时间",'%.2f' %(end - start), "秒")