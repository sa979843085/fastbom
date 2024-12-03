import pandas as pd
import os

# 定义导入BOM数据的函数

def import_bom_data(file_path):
    if file_path.endswith('.xlsx'):
        bom_data = pd.read_excel(file_path)
    elif file_path.endswith('.csv'):
        bom_data = pd.read_csv(file_path)
    else:
        print("不支持的文件格式")
        return None
    return bom_data

# 文件路径为D:\万合光电\WHJ82蒸发波导诊断系统\BOM.xlsx
bom_data = import_bom_data('D:\万合光电\WHJ82蒸发波导诊断系统\BOM.xlsx')
if bom_data is not None:
    # 将BOM数据前几行打印出来
    print(bom_data.head())  
    print("BOM数据已成功导入")
else:
    print("BOM数据导入失败")


# 1. 增加名为“文件名称”的列
if '文件名称' not in bom_data.columns:
    bom_data['文件名称'] = ''



# 2. 删除名为“规格”的列
if '规格' in bom_data.columns:
    bom_data = bom_data.drop(columns=['规格'])



# 3. 删除备注为“连接器自带电缆”的行
if bom_data is not None:
    bom_data = bom_data[bom_data['备注'] != '连接器自带电缆']



# 4. 新建列存储“父件的名称”、“父件的数量”、“总数量”
bom_data['父件的名称'] = ''
bom_data['父件的数量'] = '1'
bom_data['总数量'] = bom_data["数量"] * bom_data["父件的数量"]
# 根据零件的阶层判断其父件的名称和数量
# 如果零件的阶层为正整数，其父件是阶层为0的零件，如果阶层为1.*，其父件是阶层为1的零件，如果阶层为1.1.*，其父件是阶层为1.1的零件，以此类推
bom_data['父件的名称'] = bom_data['阶层'].apply(lambda x: x.split('.')[0])
bom_data['父件的数量'] = bom_data['阶层'].apply(lambda x: x.split('.')[1] if len(x.split('.')) > 1 else '1')
# 将父件的数量转换为整数
bom_data['父件的数量'] = bom_data['父件的数量'].astype(int)
# 计算总数量
bom_data['总数量'] = bom_data['数量'] * bom_data['父件的数量']


print(bom_data.head())  # 打印修改后的BOM数据

if bom_data is not None:
    # 将BOM数据前几行打印出来
    print(bom_data.head())
else:
    print("BOM数据修改失败")



