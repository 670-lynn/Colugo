using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using Grasshopper.Kernel;
using Grasshopper.Kernel.Data;
using Grasshopper.Kernel.Types;
using Excel = Microsoft.Office.Interop.Excel;

namespace Colugo.Components
{
    public class GridHarmonyEvaluatorComponent : GH_Component
    {
        public GridHarmonyEvaluatorComponent()
            : base("Grid Harmony Evaluator", "GridHarmony",
                "Evaluate LEED v4.1 EA Credit: Grid Harmonization Case 3.\n" +
                "Reads Honeybee hourly energy simulation data from Excel,\n" +
                "performs peak shaving analysis, 2-hour load shifting check,\n" +
                "and optional BESS simulation.",
                "Colugo", "LEED")
        {
        }

        // =================================================================
        // 1. 輸入參數
        // =================================================================
        protected override void RegisterInputParams(GH_InputParamManager pManager)
        {
            pManager.AddTextParameter("Excel Path", "Excel", "result_grouped.xlsx file path", GH_ParamAccess.item);           // 0
            pManager.AddIntegerParameter("Peak Hours", "Peak", "Grid peak hour indices (e.g. 13-17)", GH_ParamAccess.list);    // 1
            pManager.AddNumberParameter("Battery (kWh)", "BESS", "Battery capacity in kWh (0 = no battery)", GH_ParamAccess.item, 0.0); // 2
            pManager.AddBooleanParameter("Run", "Run", "Set True to execute", GH_ParamAccess.item, false);                    // 3

            pManager[1].Optional = true;
        }

        // =================================================================
        // 2. 輸出參數
        // =================================================================
        protected override void RegisterOutputParams(GH_OutputParamManager pManager)
        {
            pManager.AddIntegerParameter("Points", "Pts", "Estimated LEED points (0-2)", GH_ParamAccess.item);                // 0
            pManager.AddTextParameter("Compliance", "Report", "Compliance report text", GH_ParamAccess.item);                  // 1
            pManager.AddNumberParameter("Peak Shaving %", "PeakRed", "Peak reduction percentage", GH_ParamAccess.item);        // 2
            pManager.AddNumberParameter("Total Demand", "Demand", "Hourly total demand (W)", GH_ParamAccess.list);             // 3
            pManager.AddNumberParameter("PV Generation", "PV", "Hourly PV production (W)", GH_ParamAccess.list);               // 4
            pManager.AddNumberParameter("Net Purchased", "NetPurch", "Hourly net purchased (W)", GH_ParamAccess.list);         // 5
            pManager.AddNumberParameter("Surplus", "Surplus", "Hourly surplus (W)", GH_ParamAccess.list);                      // 6
            pManager.AddNumberParameter("Modified Load", "ModLoad", "Hourly load after BESS (W)", GH_ParamAccess.list);        // 7
            pManager.AddTextParameter("Battery Info", "BattInfo", "Battery simulation summary", GH_ParamAccess.item);          // 8
            pManager.AddTextParameter("Methodology", "Method", "Calculation methodology and LEED review notes", GH_ParamAccess.item); // 9
        }

        // =================================================================
        // 3. 主邏輯 — 先模擬、後評分
        // =================================================================
        protected override void SolveInstance(IGH_DataAccess DA)
        {
            string excelPath = "";
            var peakHours = new List<int>();
            double batteryKwh = 0;
            bool run = false;

            DA.GetData(0, ref excelPath);
            DA.GetDataList(1, peakHours);
            DA.GetData(2, ref batteryKwh);
            DA.GetData(3, ref run);

            if (!run)
            {
                DA.SetData(0, 0);
                DA.SetData(1, "Set Run = True to evaluate.");
                Message = "Idle";
                return;
            }

            if (string.IsNullOrWhiteSpace(excelPath) || !File.Exists(excelPath))
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "Excel file not found: " + excelPath);
                return;
            }

            if (peakHours.Count == 0)
                peakHours = new List<int> { 13, 14, 15, 16, 17 };

            // --- 讀取 Excel ---
            ExcelData data;
            try
            {
                data = ReadExcel(excelPath);
            }
            catch (Exception ex)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "Excel read error: " + ex.Message);
                return;
            }

            if (data.TotalDemand.Count == 0)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "No demand data found.");
                return;
            }

            int hours = data.TotalDemand.Count;

            // =============================================================
            // STEP 1: 先執行 BESS 模擬，得到 finalEvaluatedLoad
            // =============================================================
            var finalEvaluatedLoad = new List<double>();
            string battInfo = "No BESS simulation (capacity = 0).";
            double totalSavedKwh = 0;
            double annualDemandKwh = 0;
            double directUseKwh = 0;
            double ssr = 0;

            if (batteryKwh > 0)
            {
                var calcSurplus = new List<double>();
                var calcPurchased = new List<double>();
                double annualDemandWh = 0, annualDirectUseWh = 0;
                int count = Math.Min(data.TotalDemand.Count, data.Produced.Count);

                for (int i = 0; i < count; i++)
                {
                    double d = data.TotalDemand[i];
                    double p = data.Produced[i];
                    annualDemandWh += d;
                    annualDirectUseWh += Math.Min(d, p);
                    double balance = p - d;
                    if (balance > 0)
                    {
                        calcSurplus.Add(balance);
                        calcPurchased.Add(0);
                    }
                    else
                    {
                        calcSurplus.Add(0);
                        calcPurchased.Add(-balance);
                    }
                }

                // 電池逐時充放電模擬
                double capacityWh = batteryKwh * 1000;
                double currentCharge = 0;
                double totalSavedWh = 0;

                for (int i = 0; i < count; i++)
                {
                    double s = calcSurplus[i];
                    double pur = calcPurchased[i];

                    if (s > 0)
                    {
                        double room = capacityWh - currentCharge;
                        currentCharge += Math.Min(s, room);
                    }
                    else if (pur > 0 && currentCharge > 0)
                    {
                        double discharge = Math.Min(pur, currentCharge);
                        currentCharge -= discharge;
                        totalSavedWh += discharge;
                        pur -= discharge;
                    }
                    finalEvaluatedLoad.Add(pur);
                }

                totalSavedKwh = totalSavedWh / 1000.0;
                annualDemandKwh = annualDemandWh / 1000.0;
                directUseKwh = annualDirectUseWh / 1000.0;
                ssr = annualDemandKwh > 0 ? (directUseKwh + totalSavedKwh) / annualDemandKwh * 100.0 : 0;

                battInfo = string.Format(
                    "BESS Capacity: {0:F1} kWh\n" +
                    "Annual Battery Saved: {1:F1} kWh\n" +
                    "Annual Total Demand: {2:F1} kWh\n" +
                    "Direct PV Use: {3:F1} kWh\n" +
                    "Self-Sufficiency Rate (SSR): {4:F1}%",
                    batteryKwh, totalSavedKwh, annualDemandKwh, directUseKwh, ssr);
            }
            else
            {
                // 無電池：使用 NetPurchased 或 TotalDemand
                finalEvaluatedLoad = data.NetPurchased.Count > 0
                    ? new List<double>(data.NetPurchased)
                    : new List<double>(data.TotalDemand);
            }

            // =============================================================
            // STEP 2: 用優化後的曲線執行 LEED Case 3 判定
            // =============================================================

            // --- A. 尖峰削減判定 ---
            double basePeak = data.TotalDemand.Max();
            double proposedPeak = finalEvaluatedLoad.Max();
            double peakReduction = basePeak > 0 ? (basePeak - proposedPeak) / basePeak * 100.0 : 0;
            bool peakPass = peakReduction >= 10.0;

            // --- B. 2 小時負荷轉移判定 (必須在 Peak Hours 內) ---
            double shiftThreshold = basePeak * 0.10;
            int maxConsecutiveInPeak = 0;
            int currentConsecutive = 0;

            for (int i = 0; i < hours; i++)
            {
                int hourOfDay = i % 24;
                if (peakHours.Contains(hourOfDay))
                {
                    // 放電量 = 原始需量 - 優化後負載
                    double hourlyDischarge = data.TotalDemand[i] - (i < finalEvaluatedLoad.Count ? finalEvaluatedLoad[i] : data.TotalDemand[i]);
                    if (hourlyDischarge >= shiftThreshold)
                    {
                        currentConsecutive++;
                        if (currentConsecutive > maxConsecutiveInPeak)
                            maxConsecutiveInPeak = currentConsecutive;
                    }
                    else
                    {
                        currentConsecutive = 0;
                    }
                }
                else
                {
                    // 非尖峰時段：重置計數 (轉移必須在 Peak Hours 內)
                    currentConsecutive = 0;
                }
            }

            bool shiftPass = maxConsecutiveInPeak >= 2;

            // --- 計分 ---
            int points = 0;
            if (peakPass || shiftPass) points = 1;
            if (peakPass && shiftPass) points = 2;

            // =============================================================
            // STEP 3: 合規報告
            // =============================================================
            string bessNote = batteryKwh > 0
                ? string.Format("   (Evaluated with {0:F1} kWh BESS simulation)", batteryKwh)
                : "   (No BESS - evaluated using raw Net Purchased data)";

            string report = string.Format(
                "=== LEED v4.1 Grid Harmonization - Case 3 ===\n" +
                "{0}\n\n" +
                "A. Peak Shaving Analysis (10% Reduction Required):\n" +
                "   Base Peak (Total Demand):    {1:F1} W ({2:F3} kW)\n" +
                "   Proposed Peak (Optimized):   {3:F1} W ({4:F3} kW)\n" +
                "   Peak Reduction:              {5:F1}%\n" +
                "   Result:                      {6}\n\n" +
                "B. 2-Hour Load Shifting (within Peak Hours):\n" +
                "   Peak Hours:                  {7}\n" +
                "   Shifting Threshold:          {8:F1} W (10% of {1:F0} W base peak)\n" +
                "   Max Consecutive Hours:       {9}\n" +
                "   Result:                      {10}\n\n" +
                "=== Scoring ===\n" +
                "   Peak Shaving:  {6,-8}  (+1 pt if PASS)\n" +
                "   Load Shifting: {10,-8}  (+1 pt if PASS)\n" +
                "   Estimated Points: {11}\n" +
                "   Overall: {12}\n\n" +
                "=== System Prerequisites ===\n" +
                "   - Building must have permanent load management hardware\n" +
                "     (BESS, smart inverter, or BMS with DR capability)\n" +
                "   - System must be automated (no manual override for compliance)\n" +
                "   - Peak hours must align with regional utility TOU schedule",
                bessNote,
                basePeak, basePeak / 1000.0,
                proposedPeak, proposedPeak / 1000.0,
                peakReduction,
                peakPass ? "PASS" : "FAIL",
                string.Join(", ", peakHours.Select(h => h.ToString("00") + ":00")),
                shiftThreshold, maxConsecutiveInPeak,
                shiftPass ? "PASS" : "FAIL",
                points,
                (peakPass && shiftPass) ? "FULLY COMPLIANT" : (points > 0 ? "PARTIALLY COMPLIANT" : "NON-COMPLIANT"));

            // =============================================================
            // STEP 4: 計算方式說明 (Methodology)
            // =============================================================
            string methodology =
                "=== Colugo Grid Harmony Evaluator - Methodology ===\n\n" +

                "1. DATA SOURCE\n" +
                "   Input: Honeybee/EnergyPlus annual hourly simulation (8760 hrs)\n" +
                "   Excel sheets matched by keyword:\n" +
                "     - 'Electricity Dem' -> Total Demand (W)\n" +
                "     - 'Produced'        -> PV Generation (W)\n" +
                "     - 'Purchased'       -> Purchased Electricity (W)\n" +
                "     - 'Net.*Purchased'  -> Net Purchased (W)\n" +
                "     - 'Surplus'         -> Surplus fed back to grid (W)\n\n" +

                "2. EVALUATION SEQUENCE: Simulate First, Score Second\n" +
                "   If BESS capacity > 0:\n" +
                "     a. Calculate hourly surplus = PV - Demand (clamp >= 0)\n" +
                "     b. Calculate hourly deficit = Demand - PV (clamp >= 0)\n" +
                "     c. Simulate battery hour-by-hour:\n" +
                "        - Surplus hour: charge = min(surplus, remaining capacity)\n" +
                "        - Deficit hour: discharge = min(deficit, current charge)\n" +
                "        - Modified Load = deficit - discharge\n" +
                "     d. Use Modified Load as 'finalEvaluatedLoad' for scoring\n" +
                "   If BESS = 0:\n" +
                "     Use Net Purchased from Excel as 'finalEvaluatedLoad'\n\n" +

                "3. LEED CASE 3 SCORING (applied to finalEvaluatedLoad)\n\n" +

                "   A. Peak Shaving (1 point):\n" +
                "      Formula:\n" +
                "        Reduction% = (max(TotalDemand) - max(finalEvaluatedLoad))\n" +
                "                     / max(TotalDemand) x 100%\n" +
                "      Threshold: Reduction% >= 10%\n\n" +

                "   B. 2-Hour Load Shifting (1 point):\n" +
                "      For each hour i:\n" +
                "        hourlyDischarge = TotalDemand[i] - finalEvaluatedLoad[i]\n" +
                "      Condition: hourlyDischarge >= 10% of base peak\n" +
                "                 AND hour must be within user-defined Peak Hours\n" +
                "      Threshold: >= 2 consecutive qualifying hours in Peak Hours\n" +
                "      Note: consecutive count resets if hour exits Peak Hours\n\n" +

                "   C. Combined Score:\n" +
                "      0 pts = neither criterion met\n" +
                "      1 pt  = either peak shaving OR load shifting passed\n" +
                "      2 pts = both criteria met (optimal design)\n\n" +

                "4. BESS MODEL ASSUMPTIONS\n" +
                "   - Ideal battery: 100% round-trip efficiency\n" +
                "   - No max charge/discharge power limit\n" +
                "   - No self-discharge or degradation\n" +
                "   - Hourly time resolution\n" +
                "   - No cross-day storage optimization\n\n" +

                "5. LEED SUBMISSION NOTES\n" +
                "   - Results are based on EnergyPlus simulation, not metered data\n" +
                "   - Actual BESS sizing should account for round-trip losses (~85-90%)\n" +
                "   - Peak hours must match regional utility TOU schedule\n" +
                "   - Hardware prerequisites: permanent BESS/BMS with automated DR\n" +
                "   - Reviewer may request manufacturer spec sheets for BESS capacity\n\n" +

                "Reference: LEED v4.1 BD+C, EA Credit: Grid Harmonization, Case 3\n" +
                "Tool: Colugo v" + typeof(GridHarmonyEvaluatorComponent).Assembly.GetName().Version.ToString();

            Message = points > 0 ? points + " pt" + (points > 1 ? "s" : "") : "0 pts";

            // --- 設定輸出 ---
            DA.SetData(0, points);
            DA.SetData(1, report);
            DA.SetData(2, peakReduction);
            DA.SetDataList(3, data.TotalDemand);
            DA.SetDataList(4, data.Produced);
            DA.SetDataList(5, data.NetPurchased);
            DA.SetDataList(6, data.Surplus);
            DA.SetDataList(7, finalEvaluatedLoad);
            DA.SetData(8, battInfo);
            DA.SetData(9, methodology);
        }

        // =================================================================
        // 4. Excel 批次讀取
        // =================================================================
        private class ExcelData
        {
            public List<string> Labels = new List<string>();
            public List<double> TotalDemand = new List<double>();
            public List<double> Produced = new List<double>();
            public List<double> Purchased = new List<double>();
            public List<double> NetPurchased = new List<double>();
            public List<double> Surplus = new List<double>();
        }

        private ExcelData ReadExcel(string path)
        {
            var data = new ExcelData();
            Excel.Application xlApp = null;
            Excel.Workbook wb = null;

            try
            {
                xlApp = new Excel.Application();
                xlApp.Visible = false;
                xlApp.DisplayAlerts = false;
                wb = xlApp.Workbooks.Open(path, ReadOnly: true);

                var rules = new[]
                {
                    new { Key = "Demand",       Match = (Func<string, bool>)(n => n.IndexOf("Electricity Dem", StringComparison.OrdinalIgnoreCase) >= 0) },
                    new { Key = "Produced",     Match = (Func<string, bool>)(n => n.IndexOf("Produced", StringComparison.OrdinalIgnoreCase) >= 0) },
                    new { Key = "Purchased",    Match = (Func<string, bool>)(n => n.IndexOf("Purchased", StringComparison.OrdinalIgnoreCase) >= 0 && n.IndexOf("Net", StringComparison.OrdinalIgnoreCase) < 0) },
                    new { Key = "NetPurchased", Match = (Func<string, bool>)(n => n.IndexOf("Net", StringComparison.OrdinalIgnoreCase) >= 0 && n.IndexOf("Purchased", StringComparison.OrdinalIgnoreCase) >= 0) },
                    new { Key = "Surplus",      Match = (Func<string, bool>)(n => n.IndexOf("Surplus", StringComparison.OrdinalIgnoreCase) >= 0) }
                };

                foreach (Excel.Worksheet ws in wb.Worksheets)
                {
                    string sheetName = ws.Name;
                    string matchedKey = null;
                    foreach (var rule in rules)
                    {
                        if (rule.Match(sheetName)) { matchedKey = rule.Key; break; }
                    }
                    if (matchedKey == null) continue;

                    Excel.Range usedRange = ws.UsedRange;
                    object[,] rawData = usedRange.Value2 as object[,];
                    if (rawData == null) continue;

                    int rowCount = rawData.GetLength(0);
                    int colCount = rawData.GetLength(1);
                    if (colCount < 2) continue;

                    int startRow = 2;
                    for (int r = 1; r <= Math.Min(rowCount, 50); r++)
                    {
                        var cellVal = rawData[r, 1];
                        if (cellVal != null && cellVal.ToString().Contains("/"))
                        {
                            startRow = r;
                            break;
                        }
                    }

                    var labels = new List<string>();
                    var values = new List<double>();

                    for (int r = startRow; r <= rowCount; r++)
                    {
                        string label = rawData[r, 1]?.ToString() ?? "";
                        double val = 0;
                        var cellValue = rawData[r, 2];
                        if (cellValue != null)
                        {
                            if (cellValue is double dv)
                                val = dv;
                            else
                                double.TryParse(cellValue.ToString(), NumberStyles.Float, CultureInfo.InvariantCulture, out val);
                        }
                        labels.Add(label);
                        values.Add(val);
                    }

                    if (data.Labels.Count == 0 && labels.Count > 0)
                        data.Labels = labels;

                    switch (matchedKey)
                    {
                        case "Demand": data.TotalDemand = values; break;
                        case "Produced": data.Produced = values; break;
                        case "Purchased": data.Purchased = values; break;
                        case "NetPurchased": data.NetPurchased = values; break;
                        case "Surplus": data.Surplus = values; break;
                    }
                }
            }
            finally
            {
                if (wb != null) { wb.Close(false); Marshal.ReleaseComObject(wb); }
                if (xlApp != null) { xlApp.Quit(); Marshal.ReleaseComObject(xlApp); }
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }

            return data;
        }

        public override Guid ComponentGuid => new Guid("ce64da0a-3a46-4655-9c74-49e5e9e55fcf");
        protected override System.Drawing.Bitmap Icon => null;
        public override GH_Exposure Exposure => GH_Exposure.primary;
    }
}
