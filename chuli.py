import pandas as pd

def process_bom_data(bom_data):
    """
    处理BOM数据，在相同文件名称的行前面增加一行汇总行，并清空原行的内容（除了父件相关列和总数量）。
    """
    processed_bom_data = pd.DataFrame()
    current_file_rows = []

    for index, row in bom_data.iterrows():
        if current_file_rows and row['文件名称'] != current_file_rows[0]['文件名称']: # 如果当前行不是当前文件的第一行
            # 当文件名称变化时，处理当前文件名称的所有行
            total_quantity = sum([r['总数量'] for r in current_file_rows])
            new_row = current_file_rows[0].copy()
            new_row['总数量'] = total_quantity
            processed_bom_data = processed_bom_data.append(new_row, ignore_index=True)
            for r in current_file_rows:
                r[[col for col in bom_data.columns if col not in ['父件的名称', '父件的代号', '父件的数量', '总数量']]] = ''
                processed_bom_data = processed_bom_data.append(r, ignore_index=True)
            current_file_rows = []
        current_file_rows.append(row)

    # 处理最后一个文件名称的所有行
    if current_file_rows:
        total_quantity = sum([r['总数量'] for r in current_file_rows])
        new_row = current_file_rows[0].copy()
        new_row['总数量'] = total_quantity
        processed_bom_data = processed_bom_data.append(new_row, ignore_index=True)
        for r in current_file_rows:
            r[[col for col in bom_data.columns if col not in ['父件的名称', '父件的代号', '父件的数量', '总数量']]] = ''
            processed_bom_data = processed_bom_data.append(r, ignore_index=True)

    processed_bom_data.reset_index(drop=True, inplace=True)
    return processed_bom_data

