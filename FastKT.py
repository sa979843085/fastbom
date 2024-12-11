from FastMX import import_bom_data, select_file, fix_hierarchy_values, get_folder_path, update_parent_info
import pandas as pd
pd.set_option('future.no_silent_downcasting', True)
import re
import tkinter as tk
from tkinter import filedialog
import os



#######################函数定义#######################
# 定义处理函数，删除除了阶层、零件名称、零件代号、数量以外的所有列
def process_bom_data(bom_data):
    # 删除除了阶层、零件名称、零件代号、数量以外的所有列
    bom_data = bom_data[['阶层', '零件名称', '零件代号', '数量']]
    return bom_data


########################处理部分########################

def main():
    file_path = select_file()

    if not file_path:
        print("未选择文件")
        exit()

    # 导入BOM数据
    bom_data = import_bom_data(file_path)


    if bom_data is None:
        print("导入BOM数据失败")
        exit()
        bom_data = bom_data.dropna(subset=['Part Number'])

    # 获取文件夹路径
    folder_path = get_folder_path(file_path)

    # 删除无用的BOM数据
    bom_data = process_bom_data(bom_data)

    # 修正阶层格式
    bom_data = fix_hierarchy_values(bom_data)


    # 增加父件信息列
    bom_data['父件的名称'] = ''
    bom_data['父件的代号'] = ''
    bom_data['父件的数量'] = ''

    # 更新父件信息
    bom_data = update_parent_info(bom_data)

    # 删除父件的名称和数量
    if '父件的名称' in bom_data.columns  and '父件的数量' in bom_data.columns:
        bom_data = bom_data.drop(['父件的名称', '父件的数量'], axis=1)

    # 根据父件的代号进行排序
    bom_data = bom_data.sort_values(['父件的代号'], ascending=[True])


    # 将bom_data输出到folder_path下KT.xlsx
    bom_data.to_excel(os.path.join(folder_path, 'KT.xlsx'), index= True)

    # 输出路径
    output_path = os.path.join(folder_path, 'KT.xlsx')
    print(f'文件已保存到{output_path}')


if __name__ == '__main__':
    main()









