#!/usr/bin/env python
# -*- coding: utf-8 -*-
# @Author : Torre Yang Edit with Python3.6
# @Email  : klyweiwei@163.com
# @Time   : 2018/11/9 15:31

import xlwt
import xlrd
from xlutils.copy import copy

# 测试元素文档的路径(当前路径下)
files = ['./勋章条件.xlsx']
medalCondition = {}
medal = []  # 勋章list
condition = []  # 获得条件List
for file in files:
    excel = xlrd.open_workbook(file)
    for sheet in excel.sheets():
        medal = sheet.col_values(0)  # 勋章list
        condition = sheet.col_values(1)  # 获得条件

# 两列List(一一对应)组成字典: 勋章：获取条件
medalCondition = dict(zip(medal, condition))

# 复制一份需要写入用例的模板到excelNew
targetExcel = './APP端测试用例20181107.xlsx'  # 模板文档
excel = xlrd.open_workbook(targetExcel)
excelNew = copy(excel)
targetSheet = excelNew.get_sheet(0)

casesss = []
i = 1  # 从第二行开始写入 目标测试用例文档
for md, con in medalCondition.items():
    # 测试案例模板, 此部分可以以excel文档模板形式准备
    case1 = "对比接口返回的'{0}'勋章状态字段state=1,前端勋章是否点亮".format(md)
    case2 = "对比接口返回的'{0}'勋章状态字段state=0,前端勋章是否点亮".format(md)
    case3 = "'{0}'勋章已点亮，点击{0}勋章是否弹出获得勋章的事件信息".format(md)
    case4 = "'{0}'勋章已点亮，点击{0}勋章是否弹出你参与'{1}'后,将会获得'{0}'称号".format(md, con)
    case5 = "'{0}'勋章未点亮，点击{0}勋章是否弹出你参与'{1}'后,将会获得'{0}'称号".format(md, con)
    case6 = "'{0}'勋章未点亮，点击{0}勋章是否弹出获得勋章的事件信息".format(md)
    case7 = "当'{0}'事件触发，用户进入里程碑页面是否弹出勋章获得提醒".format(md)
    case8 = "当'{0}'事件触发，点击分享按钮是否成功".format(md)
    case9 = "当'{0}'事件触发，点击关闭按钮是否成功".format(md)
    cases = [case1, case2, case3, case4, case5, case6, case7, case8, case9]    # 测试案例
    expectRes = ['是', '否', '是', '否', '是', '否', '是', '是', '是']   # 预期结果
    casesExpect = dict(zip(cases, expectRes))  # 将用例和结果组成字典

# 写入模块
    targetSheet.write(i, 0, md)  # 写入模块
    for case, expect in casesExpect.items():
        targetSheet.write(i, 2, case)     # 写入测试案例的列
        targetSheet.write(i, 3, expect)   # 写入期望结果的列
        i += 1
# 保存生成的用例到new.xls文档
excelNew.save('new.xls')
