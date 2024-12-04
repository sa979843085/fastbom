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
bom_data = import_bom_data('E:/万合结构/1项目/WHJ82蒸发波导诊断系统/总BOM.xlsx')
if bom_data is not None:
    # # 将BOM数据前几行打印出来
    # print(bom_data.head())  
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
bom_data['父件的数量'] = ''



# 将“阶层”列转换为字符串格式
bom_data['阶层'] = bom_data['阶层'].astype(str)


# 遍历BOM数据修改“父件的名称”和“父件的数量”
for index, row in bom_data.iterrows():
    level = row['阶层']
    if level == '0':
        bom_data.at[index, '父件的名称'] = ''
        bom_data.at[index, '父件的数量'] = ''
    elif '.' not in level:
        # 如果没有点，父件是阶层为0的零件
        bom_data.at[index, '父件的名称'] = bom_data.loc[bom_data['阶层'] == '0', '零件名称'].values[0]
        bom_data.at[index, '父件的数量'] = bom_data.loc[bom_data['阶层'] == '0', '数量'].values[0]
    else:
        # 如果存在点，去除最后一个点及后面的部分，剩下的字符串就是父件的阶层
        parent_level = level.rsplit('.', 1)[0]
        matches_name = bom_data.loc[bom_data['阶层'] == parent_level, '零件名称']

        # 查找父件的名称和数量
        if not matches_name.empty:
            bom_data.at[index, '父件的名称'] = matches_name.values[0]
        else:
            bom_data.at[index, '父件的名称'] = ''
        
        matches_quantity = bom_data.loc[bom_data['阶层'] == parent_level, '数量']
        if not matches_quantity.empty:
            bom_data.at[index, '父件的数量'] = matches_quantity.values[0]
        else:
            bom_data.at[index, '父件的数量'] = ''

# bom_data['总数量'] = (bom_data["数量"].astype(int) * bom_data["父件的数量"].astype(int))

if bom_data is not None:
    # 在同目录下生成新的excel
    bom_data.to_excel('E:/万合结构/1项目/WHJ82蒸发波导诊断系统/BOM数据修改.xlsx', index=False)
    print("BOM数据修改成功")
else:
    print("BOM数据修改失败")







