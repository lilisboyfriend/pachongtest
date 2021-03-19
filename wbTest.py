from openpyxl import Workbook
from openpyxl import load_workbook
import time
import random
from openpyxl.styles import colors, Font, Fill, NamedStyle
from openpyxl.styles import PatternFill, Border, Side, Alignment
from openpyxl.styles import Font, Alignment  #设置单元格格式
from openpyxl.utils import get_column_letter, column_index_from_string
from openpyxl.styles import PatternFill, Border, Side, Alignment, Protection, Font
from copy import copy, deepcopy

print("开始增加数据......HW")
url1 = "D:\日报\烽火日报2021{}{}.xlsx".format(time.strftime("%m", time.localtime()),time.strftime("%d", time.localtime(time.time()-86400)))
url2 = "D:\日报\烽火并发日报2021-{}-{}.xlsx".format(time.strftime("%m",time.localtime()),time.strftime("%d",time.localtime()))
print(url1)
print(url2)
wb1 = load_workbook(url1)
wb2 = load_workbook(url2)

ws1 = wb1.active
print(ws1.title)
ws2 = wb2.active
print(ws2.title)

rmax1 = ws1.max_row
lmax1 = ws1.max_column
print(rmax1,lmax1)
rmax2 = ws2.max_row
lmax2 = ws2.max_column
print(rmax2,lmax2)

ws1.cell(lmax1+4,1).value = time.strftime("%y/%m/%d", time.localtime())

for i in range(1,rmax2+1):
    for j in range(1,lmax2+1):
        ws1.cell(j+lmax1+5,i).value = ws2.cell(j,i).value

wb1.save(url1)
wb2.save(url2)


# wb = load_workbook("D:\日报\日常巡检表.xlsx")
# ws = wb.active
# print(ws.title)
#
#
# print(ws)
# print(ws.max_column)
# print(ws.max_row)
# rmax = ws.max_row
# lmax = ws.max_column
# #
# # so = ws.cell(1,lmax)
# # tar = ws.cell(1,lmax+1)
# #
# # ws.cell(4,lmax+1).value = ws.cell(4,lmax).value+1
# #
# # if so.has_style:
# #     tar._style = copy(so._style)
# #     tar.fill = copy(so.fill)
# #
# ws.cell(1,lmax+1).value = time.strftime("%y.%m.%d", time.localtime())
# for i in range(2,rmax+1):
#     c = ws.cell(i,lmax).value
#     ws.cell(i,lmax+1).value = c
#     so = ws.cell(i,lmax)
#     tar = ws.cell(i,lmax+1)
#     if so.has_style:   #复杂单元格格式
#         tar._style = copy(so._style)
#         tar.fill = copy(so.fill)
#         tar.border = copy(so.border)
#         tar.alignment = copy(so.alignment)
#
# # ws.merge_cells(start_row=77, start_column=lmax+1, end_row=78, end_column=lmax+1)
# # ws.merge_cells(start_row=79, start_column=lmax+1, end_row=80, end_column=lmax+1)
#
#
# # ws.cell(9,lmax+1).value = 0.52+random.randint(3,6)*0.001+random.randint(0,9)*0.0001
# #
# wb.save("D:\日报\日常巡检表.xlsx")
# # print("修改完成。")
# #
# #
# # #ws.cell(1,lmax+1).value = time.strftime("%m/%d", time.localtime())
# for i in range(1,rmax):
#     c = ws.cell(i,lmax+1).value
#     print(c)