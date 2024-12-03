from pyautocad import Autocad, APoint
import openpyxl

# 连接到AutoCAD
acad = Autocad(create_if_not_exists=True)

# 加载Excel文件
file_path = r'G:\资料模板\bom.xlsx'
wb = openpyxl.load_workbook(file_path)
sheet = wb.active

# 定义一个函数来创建块
def create_block(acad, name, position, attributes):
    # 使用正确的方法创建块
    block = acad.doc.ModelSpace.AddBlock(name, position, 1, 1, 1)
    for key, value in attributes.items():
        block.SetAttributeValue(key, str(value))  # 确保属性值是字符串
    return block

# 遍历Excel中的行
for row in sheet.iter_rows(min_row=2, values_only=True):
    序号, 阶层, 零件代号, 零件名称, 规格, 数量, 材料, 单重, 总重, 备注 = row
    # 将阶层转换为字符串
    阶层 = str(阶层)

    # 创建块的属性
    attributes = {
        'PartNumber': 零件代号,
        'PartName': 零件名称,
        'Specification': 规格,
        'Quantity': 数量,
        'Material': 材料,
        'Weight': 单重
    }
    # 根据阶层确定位置
    if 阶层.startswith('3'):
        position = APoint(0, 100)  # 例如，阶层3的零件放在(0, 100)
    elif 阶层.startswith('2'):
        position = APoint(0, 200)  # 例如，阶层2的零件放在(0, 200)
    else:
        position = APoint(0, 0)  # 其他阶层默认位置

    # 创建块
    create_block(acad, 零件代号, position, attributes)

# 保存AutoCAD文档
acad.doc.Save()