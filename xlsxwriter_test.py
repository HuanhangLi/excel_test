import xlsxwriter   # XlsxWriter 只能创建文件，不能读取或修改现有的文件

# 测试 xlsxwriter模块
# workbook = xlsxwriter.Workbook('hello.xlsx')
# worksheet = workbook.add_worksheet()
#
# worksheet.write('A1', 'Hello world')
#
# workbook.close()


# 合并单元格
# book = xlsxwriter.Workbook('./combine_cell_test.xlsx')
# sheet = book.add_worksheet('sheet 1')
# # 单元格合并后居中
# fmt = book.add_format({'align': 'center', 'valign': 'vcenter'})
# # sheet.merge_range(x1, y1, x2, y2, value, cell_format=None)
# sheet.merge_range(0, 0, 1, 10, 'hello', cell_format=fmt)
# book.close()

