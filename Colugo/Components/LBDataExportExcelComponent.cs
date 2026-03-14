using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using Grasshopper.Kernel;
using Grasshopper.Kernel.Data;
using Grasshopper.Kernel.Types;
using Excel = Microsoft.Office.Interop.Excel;

namespace Colugo.Components
{
    public class LBDataExportExcelComponent : GH_Component
    {
        public LBDataExportExcelComponent()
            : base("LB Data Export Excel", "LBExportExcel",
                "Export Ladybug hourly data to a grouped Excel file (one sheet per data type), with auto date-stamped filename.",
                "Colugo", "IO")
        {
        }

        protected override void RegisterInputParams(GH_InputParamManager pManager)
        {
            pManager.AddTextParameter("File Path", "P", "Output folder or file path", GH_ParamAccess.item);
            pManager.AddGenericParameter("Header", "H", "Ladybug header objects (list)", GH_ParamAccess.list);
            pManager.AddNumberParameter("Values", "V", "Hourly data values (tree: one branch per header)", GH_ParamAccess.tree);
            pManager.AddBooleanParameter("Run", "R", "Set to True to execute", GH_ParamAccess.item, false);
        }

        protected override void RegisterOutputParams(GH_OutputParamManager pManager)
        {
            pManager.AddTextParameter("Message", "M", "Status message", GH_ParamAccess.item);
            pManager.AddTextParameter("File Path", "F", "Saved file path", GH_ParamAccess.item);
        }

        protected override void SolveInstance(IGH_DataAccess DA)
        {
            string filePath = null;
            var headerGoos = new List<IGH_Goo>();
            GH_Structure<GH_Number> valueTree = null;
            bool run = false;

            if (!DA.GetData(0, ref filePath)) return;
            if (!DA.GetDataList(1, headerGoos)) return;
            if (!DA.GetDataTree(2, out valueTree)) return;
            DA.GetData(3, ref run);

            if (!run)
            {
                DA.SetData(0, "Waiting...");
                return;
            }

            if (string.IsNullOrEmpty(filePath) || headerGoos.Count == 0)
            {
                DA.SetData(0, "Check inputs");
                return;
            }

            // --- 0. 展平 Header ---
            var headersFlat = new List<object>();
            foreach (var goo in headerGoos)
            {
                if (goo is GH_ObjectWrapper wrapper)
                    headersFlat.Add(wrapper.Value);
                else
                    headersFlat.Add(goo);
            }

            // --- 轉換 DataTree → List<List<double>> ---
            var branches = new List<List<double>>();
            for (int i = 0; i < valueTree.PathCount; i++)
            {
                var branch = valueTree.Branches[i];
                var vals = new List<double>();
                foreach (var item in branch)
                    vals.Add(item.Value);
                branches.Add(vals);
            }

            // --- 執行寫入 ---
            string msg;
            string savedPath;
            WriteGroupedExcel(filePath, headersFlat, branches, out msg, out savedPath);

            DA.SetData(0, msg);
            DA.SetData(1, savedPath);
        }

        // ===================================================================
        // 比照 Python: write_grouped_excel(path, headers_input, data_tree)
        // 唯一差異：檔名自動加上日期時間戳記
        // ===================================================================
        private void WriteGroupedExcel(
            string path,
            List<object> headersFlat,
            List<List<double>> branches,
            out string message,
            out string fileFullPath)
        {
            Excel.Application xlApp = null;
            Excel.Workbook wb = null;
            fileFullPath = null;

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

                // --- 2. 數據預處理：取得時間軸 ---
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
                        message = "Error: Header invalid.";
                        return;
                    }
                }
                else
                {
                    message = "Error: Header invalid.";
                    return;
                }

                int numRows = formattedDates.Count;

                // --- 數據預處理：按 metadata["type"] 分群 ---
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
                xlApp.Visible = false;
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

                    // 建立二維陣列
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
                    ws.Range[startCell, endCell].Value2 = dataMatrix;

                    // 格式化
                    Excel.Range headerRow = ws.Range[ws.Cells[1, 1], ws.Cells[1, totalCols]];
                    headerRow.WrapText = true;
                    ws.Columns.AutoFit();

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

                // --- 5. 存檔並關閉 ---
                wb.SaveAs(fileFullPath);
                wb.Close(false);
                wb = null;
                message = "Success: " + fileFullPath;
            }
            catch (Exception ex)
            {
                message = "Error: " + ex.Message;
                try { if (wb != null) { wb.Close(false); } } catch { }
            }
            finally
            {
                if (wb != null) { Marshal.ReleaseComObject(wb); wb = null; }
                if (xlApp != null) { xlApp.Quit(); Marshal.ReleaseComObject(xlApp); xlApp = null; }
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        // ===================================================================
        // 比照 Python: get_detailed_header_text(h)
        // ===================================================================
        private string GetDetailedHeaderText(dynamic h)
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

        public override Guid ComponentGuid => new Guid("c19286f9-3d7d-4a55-a2ba-866ea674ddd6");
        protected override System.Drawing.Bitmap Icon => null;
        public override GH_Exposure Exposure => GH_Exposure.primary;
    }
}
