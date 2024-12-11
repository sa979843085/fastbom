import pandas as pd
import re
import tkinter as tk
from tkinter import filedialog
import os

root = tk.Tk()
root.withdraw()

######################函数块######################


def select_file():
    """文件选择对话框"""
    root = tk.Tk()
    root.withdraw() # 隐藏窗口
    
    file_path = filedialog.askopenfilename(
        title="选择BOM数据文件",  
        filetypes=(("Excel files", "*.xlsx"), ("CSV files", "*.csv"), ("All files", "*.*"))
    )
    
    root.destroy() # 关闭窗口
    return file_path

# 定义导入BOM数据的函数
def import_bom_data(file_path):
    """导入BOM数据"""
    if file_path.endswith('.xlsx'):
        bom_data = pd.read_excel(file_path)
    elif file_path.endswith('.csv'):
        bom_data = pd.read_csv(file_path)
    else:
        print("不支持的文件格式")
        return None
    return bom_data


# 定义获取文件夹路径的函数
def get_folder_path(file_name):
    """获取文件夹路径"""
    return os.path.dirname(os.path.abspath(file_name))


# 定义数据格式转换函数
def format_data(bom_data):
    """定义每列的数据格式"""
    bom_data = bom_data.astype({
        '序号': int,
        '阶层': str,
        '零件名称': str,
        '零件代号': str,
        '文件名称': str,
        '数量': int,
        '材料': str,
        '单重(kg)': float,
        '总重(kg)': float,
        '备注': str
    })
    return bom_data



def delete_gg_dl(bom_data):
    """删除“规格”列和“连接器自带电缆”行"""
    if bom_data is not None and '规格' in bom_data.columns:
        bom_data = bom_data.drop(columns=['规格'])
    if bom_data is not None :
        bom_data = bom_data[bom_data['备注'] != '连接器自带电缆']
    bom_data = bom_data.reset_index(drop=True)
    return bom_data




def fix_hierarchy_values(bom_data):
    """
    修正阶层为“n.1”的情况,在当前值的末尾补一个“0”。
    """
    if bom_data.empty or '阶层' not in bom_data.columns:
        raise ValueError("数据为空或'阶层'列不存在")
    
    for i in range(1, len(bom_data)):
        current_value = str(bom_data.loc[i, '阶层'])
        previous_value = str(bom_data.loc[i-1, '阶层'])
        
        # 检查当前值是否符合“n.1”的模式
        if current_value.endswith('.1') and current_value.count('.') == 1:
            # 去掉“.1”得到n
            stripped_value = current_value[:-2]
            
            # 检查上一行的值是否是去掉“.1”后的整数n
            if previous_value != stripped_value:
                # 在当前值的末尾补一个“0”
                bom_data.at[i, '阶层'] = f"{current_value}0"
    
    return bom_data


def update_parent_info(bom_data):
    """
    更新BOM数据中的“父件的名称”、“父件的代号”和“父件的数量”。

    参数:
    bom_data (pd.DataFrame): 包含层级数据的DataFrame，必须包含'阶层'、'零件名称'、'零件代号'和'数量'列。

    返回:
    pd.DataFrame: 更新后的DataFrame。
    """
    if bom_data.empty or not {'阶层', '零件名称', '零件代号', '数量'}.issubset(bom_data.columns):
        raise ValueError("数据为空或缺少必要的列")
    
    # 将阶层列转换为字符串
    bom_data['阶层'] = bom_data['阶层'].astype(str)
    
    for index, row in bom_data.iterrows():
        level = row['阶层']
        
        if level == '0':
            bom_data.at[index, '父件的名称'] = ''
            bom_data.at[index, '父件的代号'] = ''
            bom_data.at[index, '父件的数量'] = ''
        elif '.' not in level:
            # 如果没有点，父件是阶层为0的零件
            parent_row = bom_data.loc[bom_data['阶层'] == '0'].iloc[0]
            bom_data.at[index, '父件的名称'] = parent_row['零件名称']
            bom_data.at[index, '父件的代号'] = parent_row['零件代号']
            bom_data.at[index, '父件的数量'] = parent_row['数量']
        else:
            # 如果存在点，去除最后一个点及后面的部分，剩下的字符串就是父件的阶层
            parent_level = level.rsplit('.', 1)[0]
            parent_row = bom_data.loc[bom_data['阶层'] == parent_level]
            
            if not parent_row.empty:
                parent_row = parent_row.iloc[0]
                bom_data.at[index, '父件的名称'] = parent_row['零件名称']
                bom_data.at[index, '父件的代号'] = parent_row['零件代号']
                bom_data.at[index, '父件的数量'] = parent_row['数量']
            else:
                bom_data.at[index, '父件的名称'] = ''
                bom_data.at[index, '父件的代号'] = ''
                bom_data.at[index, '父件的数量'] = ''
    
    return bom_data

# 定义总数量处理函数
def process_quantity(bom_data):

    bom_data['数量'] = pd.to_numeric(bom_data['数量'], errors='coerce') # 将数量列转换为数值类型
    bom_data['父件的数量'] = pd.to_numeric(bom_data['父件的数量'], errors='coerce') # 将父件的数量列转换为数值类型

    # 可以选择填充 nan 值
    bom_data['数量'] = bom_data['数量'].fillna(0)
    bom_data['父件的数量'] = bom_data['父件的数量'].fillna(0)


    # bom_data['总数量'] = ''
    bom_data['总数量'] = bom_data['数量'].astype(int) * bom_data['父件的数量'].astype(int) 
    return bom_data

# 定义格式化数量处理函数
# 如果总数量大于1，将“数量_格式化”填充为“数量”×“父件的数量”的文字，如果总数量小于等于1，则填充为“总数量”
def format_quantity(row):
    if row['总数量'] > 1:
        return f"{row['数量']}×{int(row['父件的数量'])}"
    else:
        return str(row['总数量'])





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
    

# 定义一个函数，用于在每个分类前增加一行
def add_category_row(group):
    # 获取分组的名称（即'分类'列的值）
    category = group.name
    # 如果分类不是“成品”，则添加新行
    if category != '0成品':
        category_row = pd.DataFrame([[''] * len(bom_data.columns)], columns=bom_data.columns)
        # 将category中的数字去除
        category = re.sub(r'\d+', '', category)
        category_row['零件名称'] = category
        return pd.concat([category_row, group], ignore_index=True)
    else:
        return group
    

# 定义主函数
def main():
    ##################处理部分###################
    # 选择文件
    file_path = select_file()
    if not file_path:
        print("未选择文件")
        exit()

    # 导入BOM数据
    bom_data = import_bom_data(file_path)
    if bom_data is None:
        print("BOM数据导入失败")
        exit()

    # 处理BOM数据

    format_data(bom_data)
    bom_data = bom_data.reset_index(drop=True)
    bom_data = fix_hierarchy_values(bom_data)

    bom_data['父件的名称'] = ''
    bom_data['父件的代号'] = ''
    bom_data['父件的数量'] = ''

    bom_data = update_parent_info(bom_data)
    bom_data['总数量'] =bom_data.apply(process_quantity , axis=1) # 计算总数量
    bom_data['数量_格式化'] = bom_data.apply(format_quantity, axis=1)# 格式化数量

    # 将父件代号转换为字符串
    bom_data['父件的代号'] = bom_data['父件的代号'].astype(str)

    # 删除没有父件代号的行
    bom_data = bom_data[bom_data['父件的代号'] != 'nan']
    bom_data = bom_data.reset_index(drop=True)

    ####### 分类 #######
    # 创建"分类"列
    bom_data['分类'] = ''

    # 将"零件代号"列转换为字符串格式
    bom_data['零件代号'] = bom_data['零件代号'].astype(str)

    # 遍历数据框，为"分类"列赋值
    for index, row in bom_data.iterrows():
        part_number = row['零件代号']
        bom_data.at[index, '分类'] = classify_part(part_number)




    # 根据分类进行pair排序,顺序：成品>组件>分组件>部件>分部件>标准件>外购件
    bom_data['分类'] = bom_data['分类'].astype(str)

    bom_data = bom_data.sort_values(['分类', '零件代号','零件名称'], ascending=[True, True, True]) # 按分类和零件代号升序排序

    bom_data = bom_data.drop(columns=['序号','父件的名称', '父件的数量'])

    # # 去除分类列里面的数字
    # bom_data['分类'] = bom_data['分类'].str.replace(r'\d+', '', regex=True)

    # 重新设置索引
    bom_data = bom_data.reset_index(drop=True)

    bom_data['整数部分'] = (bom_data['文件名称'] != bom_data['文件名称'].shift()).cumsum()
    # 小数部分：在每个文件名称组内进行计数
    bom_data['小数部分'] = bom_data.groupby('文件名称').cumcount() + 1

    # 组合整数部分和小数部分
    bom_data['排序辅助列'] = bom_data['整数部分'].astype(str) + '.' + bom_data['小数部分'].astype(str)

    # 删除辅助列
    bom_data = bom_data.drop(columns=['整数部分', '小数部分'])

    # 重置所有列的格式
    for col in bom_data.columns:
        bom_data[col] = bom_data[col].astype(str)



    ##################
    # 增加汇总行
    ##################
    # 根据文件名称分组
    grouped = bom_data.groupby('文件名称')
    # 处理大于一行的组
    new_bom_data = pd.DataFrame()  # 创建一个新的 DataFrame 用于存储修改后的数据


    for name, group in grouped:
        if len(group) > 1:
            # 增加汇总行
            total_quantity = group['总数量'].astype(int).sum().astype(str)
            total_mass = group['总重(kg)'].astype(float).sum().astype(str)
            new_row = group.iloc[0].copy()
            new_row['总数量'] = total_quantity
            new_row['单重(kg)'] = ''
            new_row['总重(kg)'] = total_mass
            new_row['父件的代号'] = ''
            new_row['数量_格式化']= ''
            # 汇总行的'排序辅助列'去掉.后面的部分
            new_row['排序辅助列'] = new_row['排序辅助列'].split('.')[0]
            # 将组内行的零件名称、零件代号、父件的代号、总数量、总重(kg)，然后设置为空

            group.loc[:, ['零件名称', '零件代号', '总数量', '总重(kg)','备注']] = ''
            
            # 将修改后的组数据添加到新的 DataFrame 中
            new_bom_data = pd.concat([new_bom_data, group], ignore_index=True)
            
            # 将新的行添加到结果 DataFrame 中
            new_bom_data = pd.concat([new_bom_data, new_row.to_frame().T], ignore_index=True)
        else:
            # 如果组内只有一个元素，直接添加到新的 DataFrame 中
            new_bom_data = pd.concat([new_bom_data, group], ignore_index=True)

    # 覆盖bom_data
    bom_data = new_bom_data

    # 排序辅助列转换为浮点数
    bom_data['排序辅助列'] = bom_data['排序辅助列'].astype(float)

    # 重新按照排序辅助列、数量进行排序
    bom_data = bom_data.sort_values(['排序辅助列', '数量'], ascending=[True, False])
    bom_data.reset_index(drop=True, inplace=True)

    # 删除辅助列
    bom_data.drop('排序辅助列', axis=1, inplace=True)

    # 清空所有字符串为nan的单元格
    bom_data = bom_data.replace('nan', '')
    # 将第一行的数量列内容赋值给第一行的父件数量列
    bom_data.loc[0, '总数量'] = bom_data.loc[0, '数量']
    # 清空第一行的数量_格式化列
    bom_data.loc[0, '数量_格式化'] = ''



    # 使用groupby并应用add_category_row函数
    bom_data = bom_data.groupby('分类', group_keys=False).apply(add_category_row, include_groups=False).reset_index(drop=True)



    # 删除阶层列、文件名称、数量列
    bom_data.drop('阶层', axis=1, inplace=True)
    bom_data.drop('文件名称', axis=1, inplace=True)
    bom_data.drop('数量', axis=1, inplace=True)



    # 重新整理列的顺序
    new_order = ['零件代号', '零件名称', '父件的代号', '数量_格式化', '总数量', '单重(kg)', '总重(kg)', '材料', '备注']
    bom_data = bom_data[new_order]


    folder_path = get_folder_path(file_path)

    if bom_data is not None:
        # 在folder_path目录下生成新的excel
        bom_data.to_excel(os.path.join(folder_path, 'BOM数据修改.xlsx'), index=False)
        # bom_data.to_excel('E:/万合结构/1项目/WHJ82蒸发波导诊断系统/BOM数据修改.xlsx', index=False)
        # bom_data.to_excel('D:/万合光电/WHJ82蒸发波导诊断系统/BOM数据修改.xlsx', index=False)
        # print("BOM数据修改成功")
        print(f"修改后文件路径:  {folder_path}/BOM数据修改.xlsx")
    else:
        print("BOM数据修改失败")


    ######################处理部分结束线######################


    
if __name__ == '__main__':
    main()

