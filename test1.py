from pyautocad import Autocad, APoint

acad = Autocad(create_if_not_exists=True)

# 画一个A3的外框
A3_size = APoint(420, 297)
start_point = APoint(0, 0)
acad.model.AddLine(start_point, APoint(A3_size.x, 0))
acad.model.AddLine(APoint(A3_size.x, 0), APoint(A3_size.x, A3_size.y))
acad.model.AddLine(APoint(A3_size.x, A3_size.y), APoint(0, A3_size.y))
acad.model.AddLine(APoint(0, A3_size.y), start_point)

# # 添加一段文本
# text_point = APoint(0, 15)
# text_string = "Hello, AutoCAD!"
# acad.model.AddText(text_string, text_point, 2.5)

# # 画一个表格然后填入hello world，，需要有网格线
# table_point = APoint(0, 30)
# table_width = 100
# table_height = 20
# table_rows = 3
# table_columns = 3



# 使用pandas读取表格，将第一行的数据填入cad中，并画出表格线
import pandas as pd

df = pd.read_excel('E:/万合结构/1项目/WHJ82蒸发波导诊断系统/框图明细.xlsx')
# for i in range(len(data)):
#     text_point = APoint(0, 30 + i * 10)
#     text_string = data.iloc[i, 0]
#     acad.model.AddText(text_string, text_point, 2.5)
#     if i % 2 == 0:
#         line_start = APoint(0, 30 + i * 10)
#         line_end = APoint(100, 30 + i * 10)
#         acad.model.AddLine(line_start, line_end)
#     else:
#         line_start = APoint(0, 30 + i * 10)
#         line_end = APoint(100, 30 + i * 10)
#         acad.model.AddLine(line_start, line_end)


header_row = df.iloc[0]
# 设置表格起始点和行高、列间距
start_point = APoint(100, 100)
row_height = 5
col_spacing = 0  # 列间距

# 计算列宽
col_widths = [len(str(value)) * 2 for value in header_row]  # 假设每个字符宽度为2

# 绘制表格边框和内部线条
current_point = start_point
for i, width in enumerate(col_widths):
    # 绘制竖直线
    acad.model.AddLine(current_point, APoint(current_point.x, current_point.y - row_height))
    current_point.x += width + col_spacing

# 绘制底部横线
acad.model.AddLine(start_point, APoint(current_point.x, start_point.y))

# 填充数据
current_point = start_point
for i, value in enumerate(header_row):
    # 计算文本插入点
    text_point = APoint(current_point.x + 2, current_point.y - row_height / 2)
    # 添加文本
    acad.model.AddText(str(value), text_point, 3)
    current_point.x += col_widths[i] + col_spacing





