import pandas as pd
import os
import re

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
# bom_data = import_bom_data('D:/万合光电/WHJ82蒸发波导诊断系统/BOM.xlsx')
if bom_data is not None:
    # # 将BOM数据前几行打印出来
    # print(bom_data.head())  
    print("BOM数据已成功导入")
else:
    print("BOM数据导入失败")



# 删除名为“规格”的列
if '规格' in bom_data.columns:
    bom_data = bom_data.drop(columns=['规格'])



# 3. 删除备注为“连接器自带电缆”的行
if bom_data is not None:
    bom_data = bom_data[bom_data['备注'] != '连接器自带电缆']


bom_data = bom_data.reset_index(drop=True)


# 新建列存储“父件的名称”、“父件的数量”、“总数量”
bom_data['父件的名称'] = ''
bom_data['父件的代号'] = ''
bom_data['父件的数量'] = ''
bom_data['总数量'] = ''



# 将“阶层”列转换为字符串格式
bom_data['阶层'] = bom_data['阶层'].astype(str)


for i in range(1, len(bom_data)):
    current_value = bom_data.loc[i, '阶层']
    previous_value = bom_data.loc[i-1, '阶层']
    # 检查当前值是否符合“n.1”的模式
    if current_value.endswith('.1') and current_value.count('.') == 1:
        # 去掉“.1”得到n
        stripped_value = current_value[:-2]
        
        # 检查上一行的值是否是去掉“.1”后的整数n
        if previous_value != stripped_value:
            # 在当前值的末尾补一个“0”
            bom_data.at[i, '阶层'] = f"{current_value}0"


# 遍历BOM数据修改“父件的名称”和“父件的数量”
for index, row in bom_data.iterrows():
    level = row['阶层']
    if level == '0':
        bom_data.at[index, '父件的名称'] = ''
        bom_data.at[index, '父件的代号'] = ''
        bom_data.at[index, '父件的数量'] = ''
    elif '.' not in level:
        # 如果没有点，父件是阶层为0的零件
        bom_data.at[index, '父件的名称'] = bom_data.loc[bom_data['阶层'] == '0', '零件名称'].values[0]
        bom_data.at[index, '父件的代号'] = bom_data.loc[bom_data['阶层'] == '0', '零件代号'].values[0]
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
        
        matches_code = bom_data.loc[bom_data['阶层'] == parent_level, '零件代号']
        if not matches_code.empty:
            bom_data.at[index, '父件的代号'] = matches_code.values[0]
        else:
            bom_data.at[index, '父件的代号'] = ''
        
        matches_quantity = bom_data.loc[bom_data['阶层'] == parent_level, '数量']
        if not matches_quantity.empty:
            bom_data.at[index, '父件的数量'] = matches_quantity.values[0]
        else:
            bom_data.at[index, '父件的数量'] = ''

# 根据数量乘以父件的数量更改总数量

bom_data['数量'] = pd.to_numeric(bom_data['数量'], errors='coerce')
bom_data['父件的数量'] = pd.to_numeric(bom_data['父件的数量'], errors='coerce')

# 可以选择填充 nan 值
bom_data['数量'] = bom_data['数量'].fillna(0)
bom_data['父件的数量'] = bom_data['父件的数量'].fillna(0)

bom_data['总数量'] = bom_data['数量'].astype(int) * bom_data['父件的数量'].astype(int) 


# 将父件代号转换为字符串
bom_data['父件的代号'] = bom_data['父件的代号'].astype(str)

# 删除没有父件代号的行
bom_data = bom_data[bom_data['父件的代号'] != 'nan']
bom_data = bom_data.reset_index(drop=True)

# 创建"分类"列
bom_data['分类'] = ''

# 将"零件代号"列转换为字符串格式
bom_data['零件代号'] = bom_data['零件代号'].astype(str)
# # 打印bom_data的零件代号列前几行
# print(bom_data['零件代号'].head())
# 定义分类函数
def classify_part(part_number):
    if re.match(r'WH.*0000$', part_number):
        return '0成品'
    if re.match(r'WH.*000$', part_number):
        return '1组件'
    elif re.match(r'WH.*00$', part_number):
        return '2分组件'
    elif re.match(r'WH.*0$', part_number):
        return '3部件'
    elif re.match(r'WH[^-]*$', part_number):
        return '4分部件'
    elif re.match(r'WH.*-.*$', part_number):
        return '5零件'
    elif 'GB' in part_number:
        return '6标准件'
    elif part_number == 'nan':
        return '7外购件'
    else:
        return '8未分类'


# 遍历数据框，为"分类"列赋值
for index, row in bom_data.iterrows():
    part_number = row['零件代号']
    bom_data.at[index, '分类'] = classify_part(part_number)

# 根据分类进行pair排序,顺序未成品>组件>分组件>部件>分部件>标准件>外购件
bom_data['分类'] = bom_data['分类'].astype(str)

bom_data = bom_data.sort_values(['分类', '文件名称', '零件名称'], ascending=[True, True, True])


# 初始化一个空的DataFrame用于存储最终结果
processed_bom_data = pd.DataFrame(columns=bom_data.columns)

# 初始化一个空列表用于存储当前文件名称的行
current_file_rows = []
current_file_name = None

for index, row in bom_data.iterrows():
    # 检查是否遇到新的文件名称或到达最后一行
    if row['文件名称'] != current_file_name and current_file_name is not None:
        # 计算总数量
        total_quantity = sum([r['总数量'] for r in current_file_rows])
        # 创建汇总行
        summary_row = current_file_rows[0].copy()
        summary_row['总数量'] = total_quantity
        # 将汇总行添加到结果DataFrame中
        processed_bom_data = processed_bom_data.append(summary_row, ignore_index=True)
        # 清空原行的内容（除了父件相关列和总数量）
        for r in current_file_rows:
            r[[col for col in bom_data.columns if col not in ['父件的名称', '父件的代号', '父件的数量', '总数量']]] = ''
            processed_bom_data = processed_bom_data.append(r, ignore_index=True)
        # 重置当前文件名称的行列表
        current_file_rows = []
    
    # 将当前行添加到当前文件名称的行列表中
    current_file_rows.append(row)
    current_file_name = row['文件名称']

# 处理最后一组行
if current_file_rows:
    total_quantity = sum([r['总数量'] for r in current_file_rows])
    summary_row = current_file_rows[0].copy()
    summary_row['总数量'] = total_quantity
    processed_bom_data = processed_bom_data.append(summary_row, ignore_index=True)
    for r in current_file_rows:
        r[[col for col in bom_data.columns if col not in ['父件的名称', '父件的代号', '父件的数量', '总数量']]] = ''
        processed_bom_data = processed_bom_data.append(r, ignore_index=True)


if bom_data is not None:
    # 在同目录下生成新的excel
    bom_data.to_excel('E:/万合结构/1项目/WHJ82蒸发波导诊断系统/BOM数据修改.xlsx', index=False)
    # bom_data.to_excel('D:/万合光电/WHJ82蒸发波导诊断系统/BOM数据修改.xlsx', index=False)
    print("BOM数据修改成功")
else:
    print("BOM数据修改失败")