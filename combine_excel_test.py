import os
import xlrd
import xlsxwriter

path = r'C:\Users\Enter\Desktop\python_test\combine_test'  # 文件夹路径

work = xlsxwriter.Workbook('output.xlsx')  # 合并后excel的路径和文件名
sheet = work.add_worksheet('combine')  # 新建一个sheet

file_list = os.listdir(path)  # 读取文件列表
file_list.sort()  # 排序

file_name = ''
x1 = 1
x2 = 1
fileNum = len(file_list)
print("在该目录下有%d个xlsx文件" % fileNum)
for file in file_list:
    if '.xlsx' in file:  # 遍历以.xlsx为后缀的文件，注意此处需要修改！
        file_name = os.path.join(path, file)
    else:
        continue

    workbook = xlrd.open_workbook(file_name)
    sheet_name = workbook.sheet_names()

    for file_1 in sheet_name:
        table = workbook.sheet_by_name(file_1)
        rows = table.nrows
        clos = table.ncols

        for i in range(rows):
            sheet.write_row('A' + str(x1), table.row_values(i))
            x1 += 1

    print('正在合并第%d个文件 ' % x2)
    print('已完成 ' + file_name)
    x2 += 1

print("已将%d个文件合并完成" % fileNum)
work.close()

# import xlwt
#
# # 文件列表
# xlxs_list = ["1/11.xlsx", "1/12.xlsx", "1/13.xlsx"]
# # 创建合并后的文件
# workbook = xlwt.Workbook(encoding='ascii')
# worksheet = workbook.add_sheet('Sheet1')
#
# # 竖
# # 行数
# count = 0
# # 表头（只写入第一个xlsx的表头）
# bt = 0
# for name in xlxs_list:
#     wb = xlrd.open_workbook(name)
#     # 按工作簿定位工作表
#     sh = wb.sheet_by_name('Sheet1')
#     # 遍历excel，打印所有数据
#     if count > 1:
#         bt = 1
#         for i in range(bt, sh.nrows):
#             k = sh.row_values(i)
#             # 遍历每一行中的每一列
#             for j in range(0, len(k)):
#                 worksheet.write(count, j, label=str(k[j]))
#                 count = count + 1
#                 workbook.save('1/合并.xlsx')
#
# # 横
# # 列数
# col = 0
# for name in xlxs_list:
#     wb = xlrd.open_workbook(name)
#
#     # 按工作簿定位工作表
#     sh = wb.sheet_by_name('Sheet1')
#     # 遍历excel，打印所有数据
#     for i in range(0, sh.nrows):
#         k = sh.row_values(i)
#         # 遍历每一行中的每一列
#         for j in range(0, len(k)):
#             worksheet.write(i, col + j, label=str(k[j]))
#             col = col + len(k)
#             workbook.save('2/合并.xlsx')
