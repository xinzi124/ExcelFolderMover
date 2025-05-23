"""
脚本说明：
本脚本用于根据病人数据Excel表格中的信息，将符合特定条件的病人对应的文件夹从源文件夹移动到目标文件夹。
支持配置多个移动任务，每个任务可以指定不同的Excel文件、Sheet页、用于匹配的列、原文件夹和目标文件夹。
还可以选择根据Excel中某一列的值对病人进行过滤。
文件名格式假定为"编号-姓名"。
日志信息会保存到 move_file.log 文件中。
终端只打印成功移动的文件夹信息和任务总结。

依赖: pandas, openpyxl
"""

import os
import shutil
import pandas as pd

# 日志文件路径
LOG_FILE = 'move_file.log'

# 打开日志文件
try:
    log_file_handle = open(LOG_FILE, 'w', encoding='utf-8')
except IOError as e:
    print(f"错误: 无法打开日志文件 {LOG_FILE} 进行写入: {e}")
    log_file_handle = None # 如果打开失败，将句柄设为 None，后续不再尝试写入

def log_message(message):
    """将消息写入日志文件，如果文件句柄有效。"""
    if log_file_handle:
        try:
            log_file_handle.write(message + '\n')
        except Exception as e:
            # 即使写入日志失败，也不中断主流程，仅打印错误到终端
            print(f"错误: 写入日志文件失败: {e}")


# ========== 配置区 ==========
# move_tasks 是一个列表，每个元素代表一个独立的文件夹移动任务。
# 每个字典元素对应一组配置项：
# 'excel_path': 需要读取的Excel文件的完整路径。
# 'sheet_name': Excel文件中要读取的Sheet页。可以是Sheet名称的字符串，或者从0开始的Sheet索引（整数）。
# 'name_col': Excel文件中包含用于匹配文件夹名（编号或姓名）的列。可以是列名的字符串，或者从0开始的列索引（整数）。
# 'header': Excel文件中作为列头的行（从0开始计数）。例如，如果列头在Excel的第二行，header应设为1。
# 'source_path': 存放待移动文件夹的源文件夹路径。
# 'destination_path': 文件夹移动的目标文件夹路径。可以使用相对路径，相对于脚本执行时的当前工作目录。
# 'filter_col': (可选) 需要根据哪一列的值进行筛选。可以是列名的字符串，或者从0开始的列索引（整数）。如果不需要过滤，可以省略此项。
# 'filter_value': (可选) 需要筛选的目标值。只有当filter_col列的值等于此值时，该行数据才会被用于匹配文件夹。如果不需要过滤，可以省略此项。
#
# 请根据您的实际情况修改以下 move_tasks 配置列表：
move_tasks = [
    {
        'excel_path': '/path/to/your/excel_file_1.xlsx',
        'sheet_name': 0,
        'name_col': 'PatientID', # Example: column name 'PatientID'
        'header': 0,
        'source_path': '/path/to/your/source_folder_1',
        'destination_path': './processed_data_1',
        'filter_col': 'Diagnosis', # Example: column name 'Diagnosis'
        'filter_value': 'UA'
    },
    {
        'excel_path': '/path/to/your/excel_file_2.xlsx',
        'sheet_name': 'Sheet1', # Example: sheet name 'Sheet1'
        'name_col': 2,          # Example: column index 2 (3rd column)
        'header': 1,
        'source_path': '/path/to/your/source_folder_2',
        'destination_path': './processed_data_2',
        # filter_col and filter_value are optional
    },
    # Add more tasks as needed, following the structure above.
    # Remember to use either column names (strings) or column indices (integers, 0-based) consistently
    # for name_col and filter_col within each task.
]
# ========== 配置区结束 ==========


# 遍历每一个文件夹移动任务
for task in move_tasks:
    # 从当前任务字典中提取配置信息
    excel_path = task['excel_path']
    sheet_name = task['sheet_name']
    name_col = task['name_col']
    header = task['header']
    source_path = task['source_path']
    destination_path = task['destination_path']
    # 使用 .get() 方法获取可选的筛选配置，如果不存在则为 None
    filter_col = task.get('filter_col') # 获取筛选列名或索引
    filter_value = task.get('filter_value') # 获取筛选目标值

    log_message(f"\n--- 开始处理任务：Excel文件 '{excel_path}', Sheet '{sheet_name}', 源文件夹 '{source_path}', 目标文件夹 '{destination_path}' ---")
    if filter_col is not None:
        log_message(f"  筛选条件： 列 '{filter_col}' 的值为 '{filter_value}'")
    else:
        log_message("  无筛选条件，使用指定列的所有数据进行匹配。")


    # 构建usecols参数列表，用于pd.read_excel只读取必要的列以提高效率
    # usecols参数需要一个列名列表或列索引列表
    use_cols_list = [name_col]
    # 如果设置了筛选列，且筛选列与姓名列不同，则添加到要读取的列列表中
    # 确保 usecols 列表中的元素类型一致
    if filter_col is not None and filter_col != name_col:
         # 尝试将 filter_col 和 name_col 都转换为整数，如果失败则保留原始类型（可能是字符串列名）
        try:
            int_name_col = int(name_col)
            int_filter_col = int(filter_col)
            use_cols_list = [int_name_col]
            if int_filter_col != int_name_col: # 再次检查避免重复添加
                 use_cols_list.append(int_filter_col)
        except (ValueError, TypeError):
            # 如果转换失败，说明使用了列名，确保 usecols 是字符串列表
             use_cols_list = [str(name_col)]
             if filter_col is not None and str(filter_col) != str(name_col): # 再次检查避免重复添加
                 use_cols_list.append(str(filter_col))

    # 读取Excel文件中的数据
    try:
        # 读取指定sheet、header和列的数据
        df = pd.read_excel(excel_path, sheet_name=sheet_name, header=header, usecols=use_cols_list)
        log_message(f"成功读取Excel文件 '{excel_path}' 的 Sheet '{sheet_name}'。")
        log_message(f"成功读取Excel文件，Sheet '{sheet_name}' 共有 {len(df.columns)} 列")
        log_message(f"列名： {df.columns.tolist()}")

        #log_message("筛选前的DataFrame（前5行）：") # 调试打印，可以根据需要注释或删除
        #log_message(df.head().to_string()) # 调试打印
        if filter_col is not None and df.shape[1] > 1:
            #log_message(f"筛选列 '{df.columns[1]}' 的前5个值：") # 调试打印
            #log_message(str(df.iloc[:5, 1].tolist())) # 调试打印
            log_message(f"目标筛选值：'{filter_value}' (类型: {type(filter_value)})")


    except FileNotFoundError:
        log_message(f"错误: Excel文件未找到 - {excel_path}，跳过当前任务。")
        print(f"错误: Excel文件未找到 - {excel_path}，跳过当前任务。详细信息请查看日志文件 {LOG_FILE}。")
        continue # 文件未找到，跳过当前任务，继续下一个
    except ValueError as e:
        # 捕获由于 sheet_name, header, usecols 等参数错误导致的读取失败
        log_message(f"读取Excel文件失败: {excel_path} - {e}")
        log_message(f"请检查任务配置中 sheet_name ({sheet_name})、header ({header})、name_col ({name_col}) 和 filter_col ({filter_col}) 是否正确。")
        print(f"读取Excel文件失败: {excel_path} - {e}。请查看日志文件 {LOG_FILE} 获取更多详情。")
        continue # 读取失败，跳过当前任务，继续下一个
    except Exception as e:
        # 捕获其他可能的读取错误
        log_message(f"读取Excel文件时发生未知错误: {excel_path} - {e}，跳过当前任务。")
        print(f"读取Excel文件时发生未知错误: {excel_path} - {e}。请查看日志文件 {LOG_FILE} 获取更多详情。")
        continue


    # 根据是否设置了筛选条件，获取用于匹配文件夹的姓名/编号列表
    names_list = [] # 初始化姓名/编号列表
    if filter_col is not None:
        # === 执行筛选操作 ===
        # 确保 DataFrame 至少有两列，第一列用于匹配，第二列用于筛选
        if df.shape[1] < 2:
             log_message(f"错误: 读取的 Sheet '{sheet_name}' 列数不足 ({df.shape[1]}列)，无法进行筛选。请检查任务配置中 name_col ({name_col}) 和 filter_col ({filter_col}) 是否正确。跳过当前任务。")
             print(f"错误: 读取的 Sheet '{sheet_name}' 列数不足，无法进行筛选。请查看日志文件 {LOG_FILE}。")
             continue

        # 根据筛选条件过滤出符合条件的行，并提取姓名/编号列的数据
        try:
            # 使用 iloc[:, 1] 访问新的 DataFrame 的第二列（索引为 1），即原始 filter_col 对应的数据
            # 使用 iloc[:, 0] 访问新的 DataFrame 的第一列（索引为 0），即原始 name_col 对应的数据
            # 对筛选列的值进行转换为字符串并去空格处理，以提高匹配容错性
            df_filtered = df.loc[df.iloc[:, 1].astype(str).str.strip() == str(filter_value).strip(), df.columns[0]]
            # 将筛选后的数据转换为字符串列表，并去掉空值
            names_list = df_filtered.dropna().astype(str).tolist()
            log_message(f"根据筛选条件获取到 {len(names_list)} 个需要移动的姓名/编号。")

            #log_message("筛选后的DataFrame（前5行）：") # 调试打印，可以根据需要注释或删除
            #log_message(df_filtered.head().to_string()) # 调试打印
            #log_message(f"从筛选结果中获取的 names_list (前5个)：{names_list[:5]}") # 调试打印


        except Exception as e:
             # 捕获筛选过程中可能发生的错误
             log_message(f"在文件 '{excel_path}' 的 Sheet '{sheet_name}' 中根据筛选条件获取姓名列表时发生错误: {e}，跳过当前任务。")
             print(f"在文件 '{excel_path}' 的 Sheet '{sheet_name}' 中根据筛选条件获取姓名列表时发生错误，请查看日志。")
             continue

    else:
        # === 不进行筛选，使用指定列的所有数据 ===
        # 确保 DataFrame 至少有一列用于匹配
        if df.shape[1] < 1:
             log_message(f"错误: 读取的 Sheet '{sheet_name}' 没有列，无法获取匹配数据。请检查任务配置中 name_col ({name_col}) 是否正确。跳过当前任务。")
             print(f"错误: 读取的 Sheet '{sheet_name}' 没有列，无法获取匹配数据。请查看日志文件 {LOG_FILE}。")
             continue
        # 直接获取新的 DataFrame 的第一列（索引为 0）的所有数据
        names_list = df.iloc[:, 0].dropna().astype(str).tolist()
        log_message(f"未设置筛选条件，获取到 {len(names_list)} 个需要移动的姓名/编号。")


    # 移除 names_list 中条目的 '001-' 前缀（如果存在）
    processed_names_list = []
    removed_prefix_count = 0
    for name in names_list:
        # 确保是字符串类型再处理
        if isinstance(name, str) and name.strip().lower().startswith('001-'):
            processed_name = name.strip()[len('001-'):]
            processed_names_list.append(processed_name)
            removed_prefix_count += 1
            log_message(f"移除前缀 '001-': 原始 '{name}' -> 处理后 '{processed_name}'")
        else:
            processed_names_list.append(name)
    names_list = processed_names_list # 更新 names_list 为处理后的列表
    if removed_prefix_count > 0:
        log_message(f"成功从 {removed_prefix_count} 个条目中移除了 '001-' 前缀。")


    # === 遍历源文件夹，查找并移动匹配的文件夹 ===
    log_message(f"开始在源文件夹 '{source_path}' 中查找匹配的文件夹...")
    moved_count = 0 # 记录成功移动的文件夹数量
    skipped_count = 0 # 记录跳过的文件夹数量（不包含'-'或格式不正确）
    not_matched_count = 0 # 记录未在Excel列表中找到匹配项的文件夹数量

    # 检查源文件夹是否存在
    if not os.path.exists(source_path):
        log_message(f"错误: 源文件夹 '{source_path}' 不存在，无法执行移动操作，跳过当前任务。")
        print(f"错误: 源文件夹 '{source_path}' 不存在。请查看日志文件 {LOG_FILE}。")
        continue # 源文件夹不存在，跳过当前任务

    # 遍历源文件夹下的所有文件和文件夹
    for entry_name in os.listdir(source_path):
        source_entry_path = os.path.join(source_path, entry_name)
        # 只处理文件夹
        if os.path.isdir(source_entry_path):
            folder_name = entry_name # 文件夹名称

            # 判断文件夹名格式，假定格式为"编号-姓名"
            if '-' in folder_name:
                folder_name_split = folder_name.split('-', 1) # 只分割一次，以防姓名中包含'-'
                # 确保分割后有两部分
                if len(folder_name_split) == 2:
                    id_in_folder = folder_name_split[0].strip() # 获取编号部分并去空格
                    name_in_folder = folder_name_split[1].strip() # 获取姓名部分并去空格

                    # 检查Excel中的值（来自 names_list）是否匹配文件夹的编号或姓名部分
                    # 将Excel中的值和文件夹名部分都转换为小写进行不区分大小写的匹配
                    matched_excel_value = None # 用于记录匹配到的Excel值
                    for excel_value in names_list:
                        # 将Excel值转换为字符串并去空格、转小写
                        processed_excel_value = str(excel_value).strip().lower()
                        # 将文件夹编号和姓名转换为小写
                        lower_id_in_folder = id_in_folder.lower()
                        lower_name_in_folder = name_in_folder.lower()

                        if processed_excel_value == lower_id_in_folder or processed_excel_value == lower_name_in_folder:
                            matched_excel_value = excel_value # 记录原始的Excel值
                            break # 找到匹配项，跳出内部循环

                    if matched_excel_value is not None:
                        # 构造目标文件夹的完整路径
                        target_folder_path = os.path.join(destination_path, folder_name)

                        # 检查目标路径父文件夹是否存在，不存在则创建
                        target_parent_dir = os.path.dirname(target_folder_path)
                        if not os.path.exists(target_parent_dir):
                            try:
                                os.makedirs(target_parent_dir)
                                log_message(f"创建目标路径父文件夹: {target_parent_dir}")
                            except OSError as e:
                                log_message(f"错误: 无法创建目标路径父文件夹 {target_parent_dir}: {e}，无法移动文件夹 {folder_name}")
                                print(f"错误: 无法创建目标路径父文件夹 {target_parent_dir}，无法移动文件夹 {folder_name}。请查看日志文件 {LOG_FILE}。")
                                skipped_count += 1 # 无法创建目标文件夹，计入跳过
                                continue # 跳过当前文件夹

                        # 检查目标文件夹是否已存在，避免重复移动或覆盖
                        if os.path.exists(target_folder_path):
                             log_message(f"目标文件夹 {target_folder_path} 已存在，跳过移动文件夹 {folder_name}")
                             print(f"目标文件夹 {target_folder_path} 已存在，跳过移动文件夹 {folder_name}。详细信息请查看日志文件 {LOG_FILE}。")
                             skipped_count += 1 # 目标文件夹已存在，计入跳过
                             continue # 跳过当前文件夹

                        # 移动文件夹
                        try:
                            shutil.move(source_entry_path, target_folder_path)
                            print(f"成功移动文件夹 '{folder_name}' 到 '{target_folder_path}' (匹配到Excel值: '{matched_excel_value}')")
                            moved_count += 1 # 成功移动，计数增加
                            log_message(f"成功移动文件夹 '{folder_name}' 到 '{target_folder_path}' (匹配到Excel值: '{matched_excel_value}')")
                        except Exception as e:
                            log_message(f"移动文件夹 '{folder_name}' 到 '{target_folder_path}' 失败: {e}")
                            print(f"移动文件夹 '{folder_name}' 到 '{target_folder_path}' 失败: {e}。详细信息请查看日志文件 {LOG_FILE}。")
                            skipped_count += 1 # 移动失败，计入跳过

                    else:
                        # 文件夹名符合格式，但在Excel列表中未找到匹配项
                        log_message(f"文件夹 '{folder_name}' (编号: '{id_in_folder}', 姓名: '{name_in_folder}') 未在Excel指定列表 ({excel_path}, Sheet '{sheet_name}', 列 '{name_col}') 中找到匹配项，跳过")
                        not_matched_count += 1 # 未匹配，计数增加

                else:
                    # 文件夹名包含'-'但分割部分数量不为2
                    log_message(f"文件夹 '{folder_name}' 格式不正确（应为'编号-姓名'，实际分割部分数量不为2），跳过")
                    skipped_count += 1 # 格式不正确，计入跳过
            else:
                # 文件夹名不包含'-'
                log_message(f"文件夹 '{folder_name}' 不包含'-'，跳过")
                skipped_count += 1 # 不包含'-'，计入跳过
        # 如果是文件，则跳过
        # else:
        #     log_message(f"'{entry_name}' 是文件，跳过") # 如果需要记录跳过的文件，可以取消注释

    log_message(f"--- 任务处理完成：成功移动 {moved_count} 个文件夹，跳过 {skipped_count} 个文件夹（格式不正确、目标已存在或移动失败），未在Excel中匹配到 {not_matched_count} 个文件夹。---")
    print(f"任务处理完成：成功移动 {moved_count} 个文件夹，详细信息请查看日志文件 {LOG_FILE}。")

print("所有任务处理完成。")

# 关闭日志文件
if log_file_handle:
    log_file_handle.close() 