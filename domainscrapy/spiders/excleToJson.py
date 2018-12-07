# -*- coding: utf-8 -*-
# author  : bunuo
# datetime   : 2018/12/5 16:28
# filename   : excleToJson.py

import xlrd
from collections import OrderedDict
import json
import codecs

def listIsEmpty(list):
    if not list:
        return True

    length = len(list)
    k = 0
    for i in list:
        if i == '' :
            k = k+1

    if length == k:
        return True

    return False

wb = xlrd.open_workbook('domainInfo.xls')

# convert_list = []
# sheet = wb.sheet_by_index(0)
# for rownum in range(sheet.nrows):
#     rowvalue = sheet.row_values(rownum)
#     print(rowvalue)
#     if not listIsEmpty(rowvalue):
#         tmp = {}
#         for item in rowvalue:
#             tmp[item] = item
#
#         convert_list.append(tmp)

convert_list = []
sh = wb.sheet_by_index(0)
title = sh.row_values(0)
print(title)
convert_list.append(title)
for rownum in range(1, sh.nrows):
    rowvalue = sh.row_values(rownum)
    single = OrderedDict()
    if not listIsEmpty(rowvalue):
        for colnum in range(0, len(rowvalue)):
            print(title[colnum], rowvalue[colnum])
            single[title[colnum]] = rowvalue[colnum]
        convert_list.append(single)

j = json.dumps(convert_list)

with codecs.open('file.json', "w", encoding='utf-8') as f:
    f.write(j)







