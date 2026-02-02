# -*- coding: utf-8 -*-
import pandas as pd
import pyodbc
import traceback
import time  # 用于计时
import numpy as np  # 引入numpy用于插入空值 NaN
import os # 引入os模块用于路径操作
import tkinter as tk # 引入tkinter用于GUI
from tkinter import filedialog # 引入filedialog用于文件对话框

# --- !!! 请根据你的环境修改下面的 pyodbc 连接字符串 !!! ---
PYODBC_CONN_STRING = "DRIVER={ODBC Driver 17 for SQL Server};SERVER=192.168.0.117;DATABASE=FQD;UID=test;PWD=Forcome123;"  # <-- 确保这里是你的连接信息

# --- SQL 查询语句 (已包含MF003作为第一列, 并重命名为'工序序号') ---
sql_query = """
WITH TopLevelItems AS (
    SELECT DISTINCT BOM.CB001 AS ProductID
    FROM BOMCB AS BOM LEFT JOIN BOMCB AS IsComponentCheck ON BOM.CB001 = IsComponentCheck.CB005
    WHERE IsComponentCheck.CB005 IS NULL AND EXISTS (SELECT 1 FROM BOMCB WHERE CB001 = BOM.CB001 AND CB005 IS NOT NULL)
), BomHierarchy AS (
    SELECT BOM.CB001 AS TopLevelProductID, BOM.CB001 AS ParentID, BOM.CB005 AS ComponentID, 1 AS BomLevel, BOM.CB008 AS ComponentUsageQty,
           CAST(RIGHT(REPLICATE('0', 50) + CAST(BOM.CB001 AS VARCHAR(50)), 50) + '/L1/' + RIGHT(REPLICATE('0', 50) + CAST(BOM.CB005 AS VARCHAR(50)), 50) AS VARCHAR(MAX)) AS SortPath
    FROM BOMCB AS BOM WHERE BOM.CB001 IN (SELECT ProductID FROM TopLevelItems) AND BOM.CB005 IS NOT NULL
    UNION ALL
    SELECT BH.TopLevelProductID, BOM_Rec.CB001 AS ParentID, BOM_Rec.CB005 AS ComponentID, BH.BomLevel + 1 AS BomLevel, BOM_Rec.CB008 AS ComponentUsageQty,
           CAST(BH.SortPath + '/L' + CAST(BH.BomLevel + 1 AS VARCHAR(2)) + '/' + RIGHT(REPLICATE('0', 50) + CAST(BOM_Rec.CB005 AS VARCHAR(50)), 50) AS VARCHAR(MAX)) AS SortPath
    FROM BOMCB AS BOM_Rec INNER JOIN BomHierarchy AS BH ON BOM_Rec.CB001 = BH.ComponentID
), CombinedResults AS (
    SELECT TLI.ProductID AS TopLevelProductID, TopLevelInfo.MB002 AS TopLevelProductName, TLI.ProductID AS ParentID, TopLevelInfo.MB002 AS ParentName,
           TopLevelInfo.MB025 AS Parent_MB025_Status, -- MB025 for TopLevel/Parent
           NULL AS ComponentID, NULL AS ComponentName, NULL AS Component_MB025_Status, -- No specific component here
           TopLevelInfo.MB003 AS ComponentSpec, NULL AS ComponentUsageQty, 0 AS BomLevel,
           Routing.MF003 AS ProcessingSequence, Routing.MF004 AS OperationCode, OpMaster.MW002 AS OperationName, Routing.MF007 AS WorkshopName,
           Routing.MF009 AS StandardManHours, Routing.MF024 AS StandardMachineHours, Routing.UDF02 AS EquipmentCode_UDF, Routing.UDF03 AS EquipmentName_UDF,
           CAST(RIGHT(REPLICATE('0', 50) + CAST(TLI.ProductID AS VARCHAR(50)), 50) + '/L0' AS VARCHAR(MAX)) AS SortPathForOrdering
    FROM TopLevelItems AS TLI INNER JOIN INVMB AS TopLevelInfo ON TLI.ProductID = TopLevelInfo.MB001
    LEFT JOIN BOMMF AS Routing ON TLI.ProductID = Routing.MF001 LEFT JOIN CMSMW AS OpMaster ON Routing.MF004 = OpMaster.MW001
    UNION ALL
    SELECT BH.TopLevelProductID, TopLevelInfo.MB002 AS TopLevelProductName, BH.ParentID, ParentInfo.MB002 AS ParentName,
           ParentInfo.MB025 AS Parent_MB025_Status, -- MB025 for Parent
           BH.ComponentID, CompInfo.MB002 AS ComponentName, CompInfo.MB025 AS Component_MB025_Status, -- MB025 for Component
           CompInfo.MB003 AS ComponentSpec,
           BH.ComponentUsageQty, BH.BomLevel, Routing.MF003 AS ProcessingSequence, Routing.MF004 AS OperationCode, OpMaster.MW002 AS OperationName, Routing.MF007 AS WorkshopName,
           Routing.MF009 AS StandardManHours, Routing.MF024 AS StandardMachineHours, Routing.UDF02 AS EquipmentCode_UDF, Routing.UDF03 AS EquipmentName_UDF, BH.SortPath AS SortPathForOrdering
    FROM BomHierarchy AS BH LEFT JOIN INVMB AS TopLevelInfo ON BH.TopLevelProductID = TopLevelInfo.MB001 LEFT JOIN INVMB AS ParentInfo ON BH.ParentID = ParentInfo.MB001
    LEFT JOIN INVMB AS CompInfo ON BH.ComponentID = CompInfo.MB001 LEFT JOIN BOMMF AS Routing ON BH.ComponentID = Routing.MF001 LEFT JOIN CMSMW AS OpMaster ON Routing.MF004 = OpMaster.MW001
)
-- 4. 最终输出并排序
-- *** 修改点 1: 在 SELECT 列表的最前面增加了 ProcessingSequence AS '工序序号' ***
SELECT ProcessingSequence AS '工序序号', -- <-- 【新增列】根据您的要求，将MF003(ProcessingSequence)作为第一列并命名
       TopLevelProductID AS '顶层成品品号', TopLevelProductName AS '顶层成品品名', ParentID AS '父项品号', ParentName AS '父项品名',
       Parent_MB025_Status AS '父项MB025',
       ComponentID AS '元件品号', ComponentName AS '元件品名',
       Component_MB025_Status AS '元件MB025',
       ComponentSpec AS '元件规格', ComponentUsageQty AS '组成用量', BomLevel AS 'BOM阶次',
       ProcessingSequence AS '加工顺序', OperationCode AS '工艺代码', OperationName AS '工艺名称', WorkshopName AS '车间名称',
       StandardManHours AS '标准人时', StandardMachineHours AS '标准机时', EquipmentCode_UDF AS '设备编号_UDF', EquipmentName_UDF AS '设备名称_UDF'
FROM CombinedResults
ORDER BY
    TopLevelProductID,
    SortPathForOrdering,
    ProcessingSequence DESC;
"""

# --- 执行查询并将结果保存到 Excel ---
conn = None
try:
    start_time = time.time()
    print("正在连接数据库...")
    # --- 数据库连接 (安全, 仅读取) ---
    conn = pyodbc.connect(PYODBC_CONN_STRING)
    print("数据库连接成功。")

    print("正在执行【读取】数据库信息的操作...")
    # --- 执行 SQL 查询 (安全, 仅读取) ---
    df_all_details = pd.read_sql(sql_query, conn)
    query_end_time = time.time()
    print(f"数据库信息【读取】完成。查询耗时: {query_end_time - start_time:.2f} 秒。")

    if not df_all_details.empty:
        # --- 数据筛选逻辑 (初步筛选，在内存中操作) ---
        print("开始在【内存中】进行初步数据筛选...")
        df_filtered = df_all_details.copy()
        filter_end_time = time.time()
        print(f"【内存中】初步筛选完成。耗时: {filter_end_time - query_end_time:.2f} 秒。")
        # --- 初步筛选逻辑结束 ---

        if not df_filtered.empty:
            # --- Part 1: 调整DataFrame结构和表头 (在内存中操作) ---
            print("\n开始在【内存中】调整结构和表头 (Part 1)...")
            adjustment_start_time_p1 = time.time()
            rename_dict_p1 = {
                '顶层成品品号': '产品编号', '顶层成品品名': '产品名称',
                '父项品号': '料号', '父项品名': '品名'
            }
            df_filtered.rename(columns=rename_dict_p1, inplace=True)
            df_filtered.insert(2, '工单单号', np.nan)
            df_filtered.insert(3, '工单单别', np.nan)
            adjustment_end_time_p1 = time.time()
            print(f"【内存中】结构和表头调整完成 (Part 1)。耗时: {adjustment_end_time_p1 - adjustment_start_time_p1:.2f} 秒。")
            # --- Part 1 调整结束 ---

            # --- Part 2: 根据BOM阶次调整 '料号' 和 '品名' (在内存中操作) ---
            print("\n开始在【内存中】调整 '料号' 和 '品名' (Part 2)...")
            adjustment_start_time_p2 = time.time()
            non_zero_bom_mask = df_filtered['BOM阶次'] != 0
            df_filtered.loc[non_zero_bom_mask, '料号'] = df_filtered.loc[non_zero_bom_mask, '元件品号']
            df_filtered.loc[non_zero_bom_mask, '品名'] = df_filtered.loc[non_zero_bom_mask, '元件品名']
            adjustment_end_time_p2 = time.time()
            print(f"【内存中】'料号' 和 '品名' 调整完成 (Part 2)。耗时: {adjustment_end_time_p2 - adjustment_start_time_p2:.2f} 秒。")
            # --- Part 2 调整结束 ---

            # --- Part 3: 数据准备、计算和最终列选择/重命名 (在内存中操作) ---
            print("\n开始在【内存中】进行数据准备和最终列格式化 (Part 3)...")
            adjustment_start_time_p3 = time.time()

            df_filtered.rename(columns={'元件规格': '物料描述'}, inplace=True)
            print(" - 列 '元件规格' 重命名为 '物料描述'")
            df_filtered['标准工时'] = df_filtered['标准人时'].fillna(0) + df_filtered['标准机时'].fillna(0)
            print(" - 计算 '标准工时'")
            if '工艺名称' in df_filtered.columns and '设备名称_UDF' in df_filtered.columns:
                df_filtered['设备名称_UDF'] = df_filtered['工艺名称']
                print(" - 数据替换: 已将原始 '工艺名称' 数据复制到 '设备名称_UDF' 列")
            else:
                print(" - 警告: '工艺名称' 或 '设备名称_UDF' 列不存在于 df_filtered，无法执行数据替换。")

            # *** 修改点 2: 将 '工序序号' 添加到列清单的最前面 ***
            intermediate_columns_source_names = [
                '工序序号',       # <-- 【新增列】确保 '工序序号' 作为第一列被选中
                '产品编号', '产品名称', '工单单号', '工单单别',
                '料号', '品名', '物料描述', 'BOM阶次',
                '设备编号_UDF', '设备名称_UDF', '车间名称', '标准工时',
                '组成用量',
                '父项MB025', '元件MB025'
            ]

            actual_intermediate_cols = []
            for col_name in intermediate_columns_source_names:
                if col_name in df_filtered.columns:
                    actual_intermediate_cols.append(col_name)
                else:
                    if col_name not in ['父项MB025', '元件MB025', '工序序号']:
                        print(f"严重警告：核心数据列 '{col_name}' 在df_filtered中不存在，后续处理可能失败或结果不准确。")
                    else:
                        print(f"警告：辅助列 '{col_name}' 在df_filtered中不存在，相关逻辑可能无法正常工作。")

            try:
                df_intermediate = df_filtered[actual_intermediate_cols].copy()
                print(f" - 已选择中间过程列，准备进行最终筛选: {actual_intermediate_cols}")
            except KeyError as e:
                print(f"错误：尝试选择列创建df_intermediate时出错，列 '{e}' 不存在于 df_filtered 中。")
                print("df_filtered 当前可用列名:", df_filtered.columns.tolist())
                raise

            final_rename_map = {
                '设备编号_UDF': '设备编号', '设备名称_UDF': '工艺名称', '车间名称': '车间',
            }
            df_intermediate.rename(columns=final_rename_map, inplace=True)
            print(f" - 已将df_intermediate中的列重命名，当前列: {list(df_intermediate.columns)}")
            adjustment_end_time_p3 = time.time()
            print(f"【内存中】数据准备和最终列格式化完成 (Part 3)。耗时: {adjustment_end_time_p3 - adjustment_start_time_p3:.2f} 秒。")
            # --- Part 3 调整结束 ---

            # --- Part 4: 最终行筛选 (根据BOM阶次=0 或 元件为自制品'M') ---
            print("\n开始在【内存中】进行最终行筛选 (Part 4)...")
            filtering_start_time_p4 = time.time()
            condition_level_0 = df_intermediate['BOM阶次'] == 0
            if '元件MB025' in df_intermediate.columns:
                condition_component_is_self_made = (df_intermediate['BOM阶次'] != 0) & (df_intermediate['元件MB025'] == 'M')
                print(" - 筛选条件: (BOM阶次为0) 或 (BOM阶次不为0 且 元件MB025为'M')")
            else:
                condition_component_is_self_made = pd.Series([False] * len(df_intermediate), index=df_intermediate.index)
                print(" - 警告: '元件MB025' 列不在df_intermediate中，筛选将仅保留BOM阶次为0的行。")
            final_filter_mask = condition_level_0 | condition_component_is_self_made
            df_final = df_intermediate[final_filter_mask].copy()
            rows_before = len(df_intermediate)
            rows_after = len(df_final)
            filtering_end_time_p4 = time.time()
            print(f" - 从 {rows_before} 行筛选至 {rows_after} 行")
            print(f"【内存中】最终行筛选完成 (Part 4)。耗时: {filtering_end_time_p4 - filtering_start_time_p4:.2f} 秒。")
            # --- Part 4 调整结束 ---

            # --- 在写入Excel前，移除辅助列 '父项MB025', '元件MB025' ---
            columns_to_drop_from_final_output = []
            if '父项MB025' in df_final.columns:
                columns_to_drop_from_final_output.append('父项MB025')
            if '元件MB025' in df_final.columns:
                columns_to_drop_from_final_output.append('元件MB025')
            if columns_to_drop_from_final_output:
                df_final.drop(columns=columns_to_drop_from_final_output, inplace=True)
                print(f"\n - 已从最终输出DataFrame中移除辅助列: {columns_to_drop_from_final_output}")

            if not df_final.empty:
                # --- 使用文件对话框选择输出路径和文件名 ---
                root = tk.Tk()
                root.withdraw()
                default_filename = "工艺资料表(计划排产模版).xlsx"
                output_filepath = filedialog.asksaveasfilename(
                    defaultextension=".xlsx",
                    filetypes=[("Excel 文件", "*.xlsx"), ("所有文件", "*.*")],
                    initialfile=default_filename,
                    title="选择保存位置和文件名"
                )
                if output_filepath:
                    output_directory = os.path.dirname(output_filepath)
                    if output_directory and not os.path.exists(output_directory):
                        try:
                            os.makedirs(output_directory)
                            print(f" - 目录 '{output_directory}' 不存在，已创建。")
                        except OSError as e:
                            print(f" - 错误: 创建目录 '{output_directory}' 失败: {e}")
                            raise
                    elif not output_directory:
                         print(f" - 将文件保存在当前工作目录。")
                    print(f"\n准备将最终结果【写入】到本地 Excel 文件: {output_filepath} ...")
                    df_final.to_excel(output_filepath, index=False, engine='openpyxl')
                    print(f"\n【成功】最终结果已保存到文件: {output_filepath}")
                    print("\n最终文件表头:")
                    print(df_final.columns.tolist())
                else:
                    print("\n用户取消了文件保存操作，未生成 Excel 文件。")
            else:
                print("\n【内存中】经过最终筛选后，没有符合条件的记录，无法生成 Excel 文件。")
        else:
            print("\n【内存中】初步筛选后没有符合条件的记录，无法进行后续调整。")
    else:
        print("\n【数据库读取】未查询到任何顶层成品或其 BOM 数据。")

# --- 错误处理和数据库连接关闭 (保持不变) ---
except pyodbc.Error as db_err:
    sqlstate = db_err.args[0]
    message = str(db_err.args[1])
    print(f"\n【数据库错误】 SQLSTATE: {sqlstate}\n消息: {message}")
    traceback.print_exc()
except pd.errors.EmptyDataError:
    print("\n【错误】读取数据库后得到空结果，无法处理。")
    traceback.print_exc()
except KeyError as key_err:
    print(f"\n【程序错误】处理数据时发生列名错误: {key_err}。请检查列名是否存在。")
    if 'df_filtered' in locals():
        print("df_filtered 当前列名:", df_filtered.columns.tolist())
    if 'df_intermediate' in locals():
        print("df_intermediate 当前列名:", df_intermediate.columns.tolist())
    traceback.print_exc()
except PermissionError as perm_err:
    print(f"\n【文件写入错误】无法保存 Excel 文件: {perm_err}。请检查文件是否被占用或文件夹权限。")
    traceback.print_exc()
except Exception as e:
    print(f"\n【未知错误】执行过程中发生错误: {e}")
    traceback.print_exc()
finally:
    if conn:
        try:
            conn.close()
            print("\n数据库连接已【安全】关闭。")
        except Exception as close_err:
            print(f"关闭数据库连接时出错: {close_err}")