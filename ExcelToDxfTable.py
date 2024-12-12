import ezdxf
from ezdxf.enums import TextEntityAlignment
from ezdxf.math import Vec2
import pandas as pd
import os
from FastMX import format_data, classify_part, import_bom_data, select_file, fix_numbers, get_folder_path, update_parent_info, add_sorting_column


# 读取Excel文件
file_path = select_file()

if not file_path:
    print("未选择文件")
    exit()

# 导入BOM数据
bom_data = import_bom_data(file_path)

if bom_data is None:
    print("导入BOM数据失败")
    exit()
# 获取文件夹路径
folder_path = get_folder_path(file_path)
# print(f'文件路径为：{folder_path}')

grouped = bom_data.groupby('父件的代号')

# 创建一个新的DXF文档
doc = ezdxf.new('R2010')
msp = doc.modelspace()

# 设置文本样式
doc.styles.new('Arial', dxfattribs={'font': 'arial.ttf'})

# 函数：绘制表格
def draw_table(msp, data, start_point, cell_width=10, cell_height=5):
    x, y = start_point
    for index, row in data.iterrows():
        # 绘制单元格
        msp.add_polyline2d([
            (x, y),
            (x + cell_width * 4, y),
            (x + cell_width * 4, y - cell_height),
            (x, y - cell_height),
            (x, y)
        ])
        
        # 添加文本
        msp.add_text(
            row['零件代号'], 
            height=1.5, 
            dxfattribs={
                'style': 'Arial', 
                'insert': Vec2(x + cell_width / 2, y - cell_height / 2),
                'align': TextEntityAlignment.CENTER
            }
        )
        
        msp.add_text(
            row['零件名称'], 
            height=1.5, 
            dxfattribs={
                'style': 'Arial', 
                'insert': Vec2(x + cell_width * 2.5, y - cell_height / 2),
                'align': TextEntityAlignment.CENTER
            }
        )
        
        msp.add_text(
            str(row['数量']), 
            height=1.5, 
            dxfattribs={
                'style': 'Arial', 
                'insert': Vec2(x + cell_width * 3.5, y - cell_height / 2),
                'align': TextEntityAlignment.CENTER
            }
        )
        
        y -= cell_height
# 绘制每个分组的表格
start_x, start_y = 0, 0
for name, group in grouped:
    draw_table(msp, group, (start_x, start_y))
    start_y -= (len(group) + 1) * 5  # 更新起始点，为下一个表格留出空间

dxf_file_path = os.path.join(folder_path, 'output.dxf')
# 保存DXF文件
doc.saveas(dxf_file_path)