using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace Colugo.Models
{
    public class LBDataExportModel
    {
        /// <summary>
        /// 比照 Python: get_detailed_header_text(h)
        /// </summary>
        public string GetDetailedHeaderText(dynamic h)
        {
            try
            {
                string dType = "Unknown";
                string unit = "";
                try { dType = h.data_type.ToString(); } catch { }
                try { unit = h.unit.ToString(); } catch { }
                string line1 = string.Format("{0} ({1})", dType, unit);

                string sysVal = "Unknown";
                try
                {
                    var meta = h.metadata;
                    try { sysVal = meta["System"]?.ToString(); } catch { }
                    if (string.IsNullOrEmpty(sysVal))
                        try { sysVal = meta["name"]?.ToString() ?? "Unknown"; } catch { }
                }
                catch { }
                string line2 = "System: " + sysVal;

                string typeVal = "Unknown";
                try { typeVal = h.metadata["type"]?.ToString() ?? "Unknown"; } catch { }
                string line3 = "Type: " + typeVal;

                return line1 + "\n" + line2 + "\n" + line3;
            }
            catch
            {
                return h.ToString();
            }
        }

        /// <summary>
        /// 比照 Python: write_grouped_excel(path, headers_input, data_tree)
        /// 唯一差異：檔名自動加上日期時間戳記
        /// </summary>
        public (string message, string filePath) WriteGroupedExcel(
            string path,
            List<object> headersFlat,
            List<List<double>> branches)
        {
            Excel.Application xlApp = null;
            Excel.Workbook wb = null;
            string fileFullPath = null;

            try
            {
                // --- 1. 路徑處理 (加上日期戳記) ---
                string timeStamp = DateTime.Now.ToString("yyyyMMdd_HHmm");
                if (Directory.Exists(path))
                {
                    fileFullPath = Path.Combine(path, string.Format("result_grouped_{0}.xlsx", timeStamp));
                }
                else
                {
                    string root = Path.ChangeExtension(path, null);
                    fileFullPath = root + "_" + timeStamp + ".xlsx";
                }

                // --- 2. 數據預處理 ---
                // formatted_dates: 從第一個 header 的 analysis_period.datetimes 取得
                var formattedDates = new List<string>();
                if (headersFlat.Count > 0)
                {
                    try
                    {
                        dynamic firstHeader = headersFlat[0];
                        var rawDates = firstHeader.analysis_period.datetimes;
                        foreach (var d in rawDates)
                            formattedDates.Add(string.Format("{0:00}/{1:00} {2:00}:00",
                                (int)d.month, (int)d.day, (int)d.hour));
                    }
                    catch
                    {
                        return ("Error: Header invalid.", null);
                    }
                }
                else
                {
                    return ("Error: Header invalid.", null);
                }

                int numRows = formattedDates.Count;

                // organized_data: 按 metadata["type"] 分群
                var organizedData = new Dictionary<string, List<(object header, List<double> values)>>();
                int branchCount = branches.Count;

                for (int i = 0; i < branchCount; i++)
                {
                    if (i >= headersFlat.Count) break;

                    dynamic h = headersFlat[i];
                    string groupKey = "Misc_Data";
                    try { groupKey = h.metadata["type"]?.ToString() ?? "Misc_Data"; } catch { }

                    if (!organizedData.ContainsKey(groupKey))
                        organizedData[groupKey] = new List<(object, List<double>)>();

                    organizedData[groupKey].Add((headersFlat[i], branches[i]));
                }

                // --- 3. 啟動 Excel ---
                xlApp = new Excel.Application();
                xlApp.Visible = true;
                xlApp.DisplayAlerts = false;
                wb = xlApp.Workbooks.Add();

                // --- 4. 針對每個 Type 建立 Sheet ---
                int sheetIndex = 1;

                foreach (var kvp in organizedData)
                {
                    string typeName = kvp.Key;
                    var columnList = kvp.Value;

                    Excel.Worksheet ws;
                    if (sheetIndex <= wb.Sheets.Count)
                        ws = (Excel.Worksheet)wb.Sheets[sheetIndex];
                    else
                        ws = (Excel.Worksheet)wb.Sheets.Add(After: wb.Sheets[wb.Sheets.Count]);
                    sheetIndex++;

                    // 工作表命名
                    string safeName = typeName.Length > 30 ? typeName.Substring(0, 30) : typeName;
                    foreach (char c in new[] { ':', '/', '\\', '?', '*', '[', ']' })
                        safeName = safeName.Replace(c.ToString(), "");
                    try { ws.Name = safeName; }
                    catch { ws.Name = "Sheet_" + sheetIndex; }

                    int numCols = columnList.Count;
                    int totalCols = 1 + numCols;

                    // 建立二維陣列 (比照 Python 的 Array.CreateInstance)
                    object[,] dataMatrix = new object[numRows + 1, totalCols];

                    // 表頭
                    dataMatrix[0, 0] = "Date/Time";
                    for (int colIdx = 0; colIdx < numCols; colIdx++)
                        dataMatrix[0, colIdx + 1] = GetDetailedHeaderText(columnList[colIdx].header);

                    // 數據
                    for (int r = 0; r < numRows; r++)
                    {
                        dataMatrix[r + 1, 0] = formattedDates[r];
                        for (int colIdx = 0; colIdx < numCols; colIdx++)
                        {
                            var dataVals = columnList[colIdx].values;
                            dataMatrix[r + 1, colIdx + 1] = (r < dataVals.Count) ? (object)dataVals[r] : 0;
                        }
                    }

                    // 寫入 Excel
                    ((Excel.Range)ws.Columns[1]).NumberFormat = "@";

                    Excel.Range startCell = (Excel.Range)ws.Cells[1, 1];
                    Excel.Range endCell = (Excel.Range)ws.Cells[numRows + 1, totalCols];
                    Excel.Range writeRange = ws.Range[startCell, endCell];
                    writeRange.Value2 = dataMatrix;

                    // 格式化
                    Excel.Range headerRow = ws.Range[ws.Cells[1, 1], ws.Cells[1, totalCols]];
                    headerRow.WrapText = true;
                    ws.Columns.AutoFit();

                    // 欄寬上限 50 (比照 Python: for c in range(2, total_cols + 2))
                    for (int c = 2; c <= totalCols + 1; c++)
                    {
                        try
                        {
                            var col = (Excel.Range)ws.Columns[c];
                            if ((double)col.ColumnWidth > 50)
                                col.ColumnWidth = 50;
                        }
                        catch { }
                    }
                }

                // --- 5. 存檔 ---
                wb.SaveAs(fileFullPath);
                return ("Success: " + fileFullPath, fileFullPath);
            }
            catch (Exception ex)
            {
                return ("Error: " + ex.Message, fileFullPath);
            }
            finally
            {
                if (wb != null) Marshal.ReleaseComObject(wb);
                if (xlApp != null) Marshal.ReleaseComObject(xlApp);
            }
        }
    }
}
