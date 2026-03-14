# -*- coding: utf-8 -*-
"""
LB Data Export Excel - Grasshopper GHPython Component
=====================================================
將 Ladybug 逐時數據依 metadata type 分群輸出至 Excel
檔名自動加上日期時間戳記

Inputs:
    file_path (str): 輸出資料夾或檔案路徑
    header (list): Ladybug header 物件
    values (DataTree): 逐時數據 (每 branch 對應一個 header)
    run (bool): 設為 True 執行

Outputs:
    a (str): 狀態訊息
    b (str): 儲存的檔案路徑
"""

import os
import datetime
import System
import clr
from System import Array
from System.Collections.Generic import Dictionary, List

clr.AddReference("Microsoft.Office.Interop.Excel")
import Microsoft.Office.Interop.Excel as Excel

from Grasshopper import DataTree
from Grasshopper.Kernel.Data import GH_Path


def get_detailed_header_text(h):
    try:
        d_type = getattr(h, 'data_type', 'Unknown')
        unit = getattr(h, 'unit', '')
        line1 = "{} ({})".format(d_type, unit)

        meta = getattr(h, 'metadata', {})
        sys_val = meta.get('System')
        if not sys_val:
            sys_val = meta.get('name', 'Unknown')
        line2 = "System: " + str(sys_val)

        type_val = meta.get('type', 'Unknown')
        line3 = "Type: " + str(type_val)

        return line1 + "\n" + line2 + "\n" + line3
    except:
        return str(h)


def write_grouped_excel(path, headers_input, data_tree):
    xlApp = None
    wb = None
    file_full_path = None
    try:
        # --- 0. 處理 Header 資料結構 ---
        headers_flat = []
        if hasattr(headers_input, "BranchCount"):
            for i in range(headers_input.BranchCount):
                branch = headers_input.Branch(i)
                for item in branch:
                    headers_flat.append(item)
        else:
            headers_flat = list(headers_input)

        # --- 1. 路徑處理 (加上日期時間戳記) ---
        time_stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M")
        if os.path.isdir(path):
            file_full_path = os.path.join(path, "result_grouped_{}.xlsx".format(time_stamp))
        else:
            root, ext = os.path.splitext(path)
            file_full_path = "{}_{}.xlsx".format(root, time_stamp)

        # --- 2. 數據預處理 ---
        organized_data = {}
        branch_count = data_tree.BranchCount

        formatted_dates = []
        if len(headers_flat) > 0 and hasattr(headers_flat[0], 'analysis_period'):
            raw_dates = headers_flat[0].analysis_period.datetimes
            formatted_dates = ["{:02d}/{:02d} {:02d}:00".format(d.month, d.day, d.hour) for d in raw_dates]
        else:
            return "Error: Header invalid.", None

        num_rows = len(formatted_dates)

        for i in range(branch_count):
            if i < len(headers_flat):
                h = headers_flat[i]
                try:
                    group_key = h.metadata.get('type', 'Misc_Data')
                except:
                    group_key = 'Misc_Data'

                if group_key not in organized_data:
                    organized_data[group_key] = []

                organized_data[group_key].append((h, data_tree.Branch(i)))

        # --- 3. 啟動 Excel ---
        xlApp = Excel.ApplicationClass()
        xlApp.Visible = True
        xlApp.DisplayAlerts = False
        wb = xlApp.Workbooks.Add()

        # --- 4. 針對每個 Type 建立 Sheet ---
        sheet_index = 1

        for type_name, column_list in organized_data.items():
            if sheet_index <= wb.Sheets.Count:
                ws = wb.Sheets[sheet_index]
            else:
                ws = wb.Sheets.Add(After=wb.Sheets[wb.Sheets.Count])
            sheet_index += 1

            safe_name = str(type_name)[:30]
            safe_name = safe_name.replace(":", "").replace("/", "").replace("\\", "").replace("?", "").replace("*", "").replace("[", "").replace("]", "")
            try:
                ws.Name = safe_name
            except:
                ws.Name = "Sheet_" + str(sheet_index)

            num_cols = len(column_list)
            total_cols = 1 + num_cols
            data_matrix = Array.CreateInstance(object, num_rows + 1, total_cols)

            data_matrix[0, 0] = "Date/Time"
            for col_idx in range(num_cols):
                h_obj = column_list[col_idx][0]
                data_matrix[0, col_idx + 1] = get_detailed_header_text(h_obj)

            for r in range(num_rows):
                data_matrix[r + 1, 0] = formatted_dates[r]
                for col_idx in range(num_cols):
                    data_vals = column_list[col_idx][1]
                    if r < len(data_vals):
                        data_matrix[r + 1, col_idx + 1] = data_vals[r]
                    else:
                        data_matrix[r + 1, col_idx + 1] = 0

            ws.Columns[1].NumberFormat = "@"

            start_cell = ws.Cells[1, 1]
            end_cell = ws.Cells[num_rows + 1, total_cols]
            write_range = ws.Range[start_cell, end_cell]
            write_range.Value2 = data_matrix

            header_row = ws.Range[ws.Cells[1, 1], ws.Cells[1, total_cols]]
            header_row.WrapText = True
            ws.Columns.AutoFit()

            for c in range(2, total_cols + 2):
                try:
                    if ws.Columns[c].ColumnWidth > 50:
                        ws.Columns[c].ColumnWidth = 50
                except:
                    pass

        # --- 5. 存檔 ---
        wb.SaveAs(file_full_path)

        return "Success: " + file_full_path, file_full_path

    except Exception as e:
        return "Error: " + str(e), file_full_path


# --- 主程式執行區 ---
a = "Waiting..."
b = None

if 'run' in globals() and run:
    if file_path and header and values:
        msg, saved_path = write_grouped_excel(file_path, header, values)
        a = msg
        b = saved_path
    else:
        a = "Check inputs"
elif file_path and header and values:
    msg, saved_path = write_grouped_excel(file_path, header, values)
    a = msg
    b = saved_path
