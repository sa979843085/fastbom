from FastMX import import_bom_data, select_file, fix_hierarchy_values, get_folder_path, update_parent_info, add_sorting_column
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
    bom_data = bom_data[['阶层', '零件名称', '零件代号', '数量','备注']]
    return bom_data

def delete_dl(bom_data):
    if bom_data is not None :
        bom_data = bom_data[bom_data['备注'] != '连接器自带电缆']
    bom_data = bom_data.reset_index(drop=True)


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

    # 删除连接器自带电缆
    bom_data = delete_dl(bom_data)

    # 阶层格式定义为字符串
    bom_data['阶层'] = bom_data['阶层'].astype(str)
    # 修正阶层格式
    bom_data = fix_hierarchy_values(bom_data)


    # 增加父件信息列
    bom_data['父件的名称'] = ''
    bom_data['父件的代号'] = ''
    bom_data['父件的数量'] = ''

    # 更新父件信息
    bom_data = update_parent_info(bom_data)

    bom_data = bom_data[bom_data['父件的代号'] != 'nan']# 删除没有父件代号的行

    bom_data = add_sorting_column(bom_data)# 添加排序列
    bom_data['排序辅助列'] = bom_data['排序辅助列'].astype(float) # 转换为浮点数
    bom_data = bom_data.sort_values(['父件的名称', '排序辅助列'], ascending=[True, False])
    bom_data.reset_index(drop=True, inplace=True) # 重置索引

    # 重新设置列顺序，同时清除不需要的列
    new_order = ['零件代号', '零件名称', '数量', '阶层']
    bom_data = bom_data[new_order]

    # 将bom_data输出到folder_path下KT.xlsx
    bom_data.to_excel(os.path.join(folder_path, 'KT.xlsx'), index= True)

    # 输出路径
    output_path = os.path.join(folder_path, 'KT.xlsx')
    print(f'文件已保存到{output_path}')


if __name__ == '__main__':
    main()









