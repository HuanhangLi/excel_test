import xlrd

work_book = xlrd.open_workbook('test.xlsx')

# print(work_book)  # 一个对象
# print(work_book.nsheets)  # 工作簿sheet表的数量

sheets = work_book.sheets()  # 获取所有sheet表对象
# print(len(sheets))

sheets_name = work_book.sheet_names()  # 获取工作簿中所有sheet表对象的名称
# print(sheets_name)

sheet_1 = work_book.sheet_by_index(0)  # 按索引获取sheet对象
# print(sheet_1)  # 打印sheet_1对象
# print(sheet_1.name)   # 打印sheet_1名称

sheet_2 = work_book.sheet_by_name('Sheet2')   # 按sheet表名称获取sheet对象，名称区分大小写
# print(sheet_2)   # 打印sheet_2对象
# print(sheet_2.name)   # 打印sheet_2名称

cell_0 = sheet_1.cell(0, 0)   # 获取sheet表单元格对象，单元格数据类型：单元格值
# print(cell_0)

cell_0_value = sheet_1.cell_value(0, 0)   # 获取sheet表单元格值
# print(cell_0_value)

cell_0_type = sheet_1.cell_type(0, 0)   # 获取单元格类型
# print(cell_0_type)

row_sum = sheet_1.nrows   # 获取sheet表对象的有效行数
# print(row_sum)

row_len = sheet_1.row_len(0)   # 获取sheet表某一行长度
# print(row_len)

row_0 = sheet_1.row(0)   # 获取sheet表某一行所有数据类型及值  结果：[number:123.0, text:'hello']
# print(row_0)

row_0_s = sheet_1.row_slice(0, 1, 2)   # 获取某一行数据类型、值，可指定开始结束列
"def row_slice(self, rowx, start_colx=0, end_colx=None):"
# print(row_0_s)

row_0_type = sheet_1.row_types(0)   # 获取sheet表对象某一行数据类型，返回一个数组对象
# print(row_0_type)

row_0_value = sheet_1.row_values(0)
# print(row_0_value)

rows = sheet_1.get_rows()   # 获取sheet对象所有行对象生成器
# print(rows)
# for row in rows:
#     print(row)


col_sum = sheet_1.ncols   # 获取sheet表有效列数
# print(col_sum)

col_0 = sheet_1.col_slice(0, 0, 10)   # 获取列对象
# print(col_0)

col_0_value = sheet_1.col_values(0)   # 获取某一列的值
# print(col_0_value)

col_0_type = sheet_1.col_types(0)   # 获取某一列的数据类型
# print(col_0_type)

# 按行读取sheet对象中所有数据
data_row = []
for row in range(sheet_1.nrows):
    data_row.append(sheet_1.row_values(row))   # 每一行是列表，放在data_row这个大列表里面
# print(data_row)

# 按列读取sheet对象中所有数据
data_col = [sheet_1.col_values(i) for i in range(sheet_1.ncols)]
# print(data_col)


# 按行读取excel文件所有sheet表数据
# 构造一个字典，按行读取文件所有数据
all_data = {}
for i, sheet_obj in enumerate(work_book.sheets()):
    all_data[i] = [sheet_obj.row_values(row)
                   for row in range(sheet_obj.nrows)]
print(all_data)

