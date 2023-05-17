# coding=utf-8

import sys
import os
import pulp
import decimal

'''
MyProbLP = pulp.LpProblem("LPProbDemo1", sense=pulp.LpMaximize)
x1 = pulp.LpVariable('x1', lowBound=0, upBound=1, cat='Binary')
x2 = pulp.LpVariable('x2', lowBound=0, upBound=1, cat='Binary')
x3 = pulp.LpVariable('x3', lowBound=0, upBound=1, cat='Binary')
x4 = pulp.LpVariable('x4', lowBound=0, upBound=1, cat='Binary')
x5 = pulp.LpVariable('x5', lowBound=0, upBound=1, cat='Binary')

MyProbLP += 1*x1 + 2*x2 + 3*x3 +4*x4 + 5*x5  	# 设置目标函数
MyProbLP += (1*x1 + 2*x2 + 3*x3 +4*x4 + 5*x5  <= 8)  # 不等式约束
# MyProbLP += (x1 + 3*x2 + x3 <= 12)  # 不等式约束
# MyProbLP += (x1 + x2 + x3 == 7)  # 等式约束
MyProbLP.solve()
print("Status:", pulp.LpStatus[MyProbLP.status]) # 输出求解状态
for v in MyProbLP.variables():
    print(v.name, "=", v.varValue)  # 输出每个变量的最优值
print("F(x) = ", pulp.value(MyProbLP.objective))  #输出最优解的目标函数值
'''



# 使用字典定义
# 1. 建立问题
AlloyModel = pulp.LpProblem("发票金额的数组中若干数字相加等于指定收款数字", sense=pulp.LpMinimize)

# 2. 建立变量
print("请输入要组合的金额列表")
amnt_str=input()
if amnt_str=='':
    sys.exit()

# 2. 建立变量
# 金额序列
amnt_list=[]
for x in amnt_str.replace("，",",").split(','):
    # k= decimal.Decimal(x)
    k=float(x)
    amnt_list.append(k)


# 设置目标值:收款金额
print("请输入要凑的目标数字:")
goal=input()
if goal=="":
    sys.exit()

if goal=="0":
    sys.exit()

payment= float(goal)
# 532620.87


# 定义每个变量的权重，即将每个金额当成一个权重
# 构造数据字典
amnt_dict = {}
for i in range(1,len(amnt_list)+1):
    # 从金额列表取值
    item="X{}".format(i)
    amnt_dict[item]=amnt_list[i-1]

print(amnt_dict)

# 所有的变量:X1,X2,X3,....X540 ，每个变量取值0或者1
#字典中的key转换为列表
var_percent = list(amnt_dict.keys())
print("0/权重变量")
print(var_percent)

mass = pulp.LpVariable.dicts("发票", var_percent, lowBound=0, cat='Binary')

# 3. 设置目标函数,施加约束
# 发票金额

AlloyModel += pulp.lpSum([amnt_dict[item] * mass[item] for item in var_percent]), "求最小发票总金额"
AlloyModel += pulp.lpSum([amnt_dict[item] * mass[item] for item in var_percent]) >= payment, "发票总金额必须大于目标"
# 5. 求解
AlloyModel.solve()
# 6. 打印结果
print(AlloyModel)  # 输出问题设定参数和条件
print("最佳组合为:", pulp.LpStatus[AlloyModel.status])
for v in AlloyModel.variables():
    if v.varValue!=0:
        # print(v.name, "=", v.varValue)
        s=v.name.replace("发票_","")
        print(amnt_dict[s])
print("总值 = ", pulp.value(AlloyModel.objective))
