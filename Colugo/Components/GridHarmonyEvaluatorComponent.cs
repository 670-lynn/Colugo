using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using Grasshopper.Kernel;
using Excel = Microsoft.Office.Interop.Excel;

namespace Colugo.Components
{
    public class GridHarmonyEvaluatorComponent : GH_Component
    {
        public GridHarmonyEvaluatorComponent()
            : base("EA Credit_Grid Harmonization", "GridHarmony",
                "LEED v4.1 EA Credit: Grid Harmonization.\n" +
                "Baseline = NetPurchased (grid-facing demand after PV).\n" +
                "Evaluates BESS contribution to peak shaving and load shifting.",
                "Colugo", "LEED")
        {
        }

        // =================================================================
        // 1. 輸入參數
        // =================================================================
        protected override void RegisterInputParams(GH_InputParamManager pManager)
        {
            pManager.AddTextParameter("Excel Path", "Excel", "Honeybee result_grouped.xlsx file path", GH_ParamAccess.item);                  // 0
            pManager.AddIntegerParameter("Peak Hours", "Peak", "Grid peak hour indices (e.g. 13-17)", GH_ParamAccess.list);                    // 1
            pManager.AddNumberParameter("Battery (kWh)", "BESS", "Battery capacity in kWh (0 = no battery)", GH_ParamAccess.item, 0.0);       // 2
            pManager.AddNumberParameter("Efficiency", "Eff", "Battery round-trip efficiency (0-1, default 0.9 = 90%)", GH_ParamAccess.item, 0.9); // 3
            pManager.AddNumberParameter("Grid Peak (kW)", "GridPk", "Regional grid peak load (kW) for alignment check", GH_ParamAccess.item, 0.0); // 4
            pManager.AddBooleanParameter("Run", "Run", "Set True to execute", GH_ParamAccess.item, false);                                    // 5

            pManager[1].Optional = true;
            pManager[4].Optional = true;
        }

        // =================================================================
        // 2. 輸出參數
        // =================================================================
        protected override void RegisterOutputParams(GH_OutputParamManager pManager)
        {
            pManager.AddIntegerParameter("Points", "Pts", "Estimated LEED points (0-2)", GH_ParamAccess.item);                                     // 0
            pManager.AddTextParameter("Compliance", "Report", "Compliance report text", GH_ParamAccess.item);                                       // 1
            pManager.AddNumberParameter("Peak Shaving %", "PeakRed", "Peak reduction percentage", GH_ParamAccess.item);                             // 2
            pManager.AddNumberParameter("Net Purchased", "NetPurch", "Hourly grid demand without BESS: max(0, Demand-PV) (W)", GH_ParamAccess.list); // 3
            pManager.AddNumberParameter("BESS Discharge", "Discharge", "Hourly BESS discharge (W)", GH_ParamAccess.list);                           // 4
            pManager.AddNumberParameter("Grid Load", "GridLoad", "Hourly grid demand with BESS: NetPurch - Discharge (W)", GH_ParamAccess.list);    // 5
            pManager.AddTextParameter("Battery Info", "BattInfo", "Battery simulation summary", GH_ParamAccess.item);                               // 6
            pManager.AddTextParameter("Methodology", "Method", "Calculation methodology with actual values", GH_ParamAccess.item);                  // 7
            pManager.AddTextParameter("Recommended", "Rec", "Recommended minimum BESS capacity for 1pt and 2pts", GH_ParamAccess.item);        // 8
        }

        // =================================================================
        // 3. 主邏輯
        //    基準線 = NetPurchased (電網實際供電量，已扣除 PV)
        //    評估 BESS 能在此基準上再削減多少
        // =================================================================
        protected override void SolveInstance(IGH_DataAccess DA)
        {
            string excelPath = "";
            var peakHours = new List<int>();
            double batteryKwh = 0;
            double efficiency = 0.9;
            double gridPeakKw = 0;
            bool run = false;

            DA.GetData(0, ref excelPath);
            DA.GetDataList(1, peakHours);
            DA.GetData(2, ref batteryKwh);
            DA.GetData(3, ref efficiency);
            DA.GetData(4, ref gridPeakKw);
            DA.GetData(5, ref run);

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

            efficiency = Math.Max(0.01, Math.Min(1.0, efficiency));

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
            // STEP 1: 計算 NetPurchased = max(0, Demand - PV)
            //   這是電網實際看到的負載（基準線）
            // =============================================================
            var netPurchased = new double[hours];
            var surplus = new double[hours];

            if (data.Produced.Count > 0)
            {
                int count = Math.Min(hours, data.Produced.Count);
                for (int i = 0; i < count; i++)
                {
                    double d = data.TotalDemand[i];
                    double p = data.Produced[i];
                    netPurchased[i] = Math.Max(0, d - p);
                    surplus[i] = Math.Max(0, p - d);
                }
            }
            else
            {
                // 沒有 PV 資料 → 電網供全部
                for (int i = 0; i < hours; i++)
                    netPurchased[i] = data.TotalDemand[i];
            }

            double basePeakW = netPurchased.Max();
            double basePeakKw = basePeakW / 1000.0;

            // =============================================================
            // STEP 2: BESS 模擬
            //   從 PV 餘電充電（乘效率），於缺電時放電
            //   GridLoad[i] = NetPurchased[i] - BESS_Discharge[i]
            // =============================================================
            var bessDischarge = new double[hours];
            var gridLoad = new double[hours];
            Array.Copy(netPurchased, gridLoad, hours);

            string battInfo = "No BESS simulation (capacity = 0 kWh).";

            if (batteryKwh > 0)
            {
                if (data.Produced.Count == 0)
                {
                    AddRuntimeMessage(GH_RuntimeMessageLevel.Warning,
                        "BESS capacity specified but no PV production data found. Cannot simulate battery charge source.");
                }
                else
                {
                    double capacityWh = batteryKwh * 1000;
                    double currentCharge = 0;
                    double totalDischargedWh = 0;

                    for (int i = 0; i < hours; i++)
                    {
                        if (surplus[i] > 0)
                        {
                            // PV 餘電 → 充電（乘效率）
                            double effectiveSurplus = surplus[i] * efficiency;
                            double room = capacityWh - currentCharge;
                            currentCharge += Math.Min(effectiveSurplus, room);
                        }
                        else if (netPurchased[i] > 0 && currentCharge > 0)
                        {
                            // 需購電 → 放電抵消
                            double discharge = Math.Min(netPurchased[i], currentCharge);
                            currentCharge -= discharge;
                            totalDischargedWh += discharge;
                            bessDischarge[i] = discharge;
                        }

                        gridLoad[i] = netPurchased[i] - bessDischarge[i];
                    }

                    double totalDischargedKwh = totalDischargedWh / 1000.0;
                    double gridLoadPeakKw = gridLoad.Max() / 1000.0;

                    battInfo = string.Format(
                        "BESS Capacity: {0:F1} kWh\n" +
                        "Round-trip Efficiency: {1:P0}\n" +
                        "Annual BESS Discharge: {2:F1} kWh\n" +
                        "Grid Peak (without BESS): {3:F3} kW\n" +
                        "Grid Peak (with BESS): {4:F3} kW",
                        batteryKwh, efficiency,
                        totalDischargedKwh,
                        basePeakKw, gridLoadPeakKw);
                }
            }

            // =============================================================
            // STEP 3: LEED Case 3 評分
            //   基準 = max(NetPurchased)  電網看到的尖峰
            //   提案 = max(GridLoad)      BESS 介入後電網看到的尖峰
            // =============================================================

            // --- A. 尖峰削減 (Peak Shaving) ---
            double proposedPeakW = gridLoad.Max();
            double proposedPeakKw = proposedPeakW / 1000.0;
            double peakReduction = basePeakW > 0 ? (basePeakW - proposedPeakW) / basePeakW * 100.0 : 0;
            bool peakPass = peakReduction >= 10.0;

            // --- B. 2 小時負荷轉移 (Load Shifting) ---
            double shiftThresholdW = basePeakW * 0.10;
            double shiftThresholdKw = shiftThresholdW / 1000.0;
            int maxConsecutiveInPeak = 0;
            int currentConsecutive = 0;

            for (int i = 0; i < hours; i++)
            {
                int hourOfDay = i % 24;
                if (peakHours.Contains(hourOfDay))
                {
                    if (bessDischarge[i] >= shiftThresholdW)
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
                    currentConsecutive = 0;
                }
            }

            bool shiftPass = maxConsecutiveInPeak >= 2;

            // --- 計分 ---
            int points = 0;
            if (peakPass || shiftPass) points = 1;
            if (peakPass && shiftPass) points = 2;

            // --- C. 電網對接驗證 (Grid Alignment) ---
            double totalDischargeInPeak = 0;
            double totalDischargeAll = 0;
            int dischargeHoursInPeak = 0;
            int dischargeHoursTotal = 0;

            for (int i = 0; i < hours; i++)
            {
                if (bessDischarge[i] > 0)
                {
                    dischargeHoursTotal++;
                    totalDischargeAll += bessDischarge[i];
                    int hourOfDay = i % 24;
                    if (peakHours.Contains(hourOfDay))
                    {
                        dischargeHoursInPeak++;
                        totalDischargeInPeak += bessDischarge[i];
                    }
                }
            }

            double gridAlignmentPct = totalDischargeAll > 0
                ? totalDischargeInPeak / totalDischargeAll * 100.0
                : 0;

            string gridAlignNote = gridPeakKw > 0
                ? string.Format(
                    "   Grid Peak Reference:         {0:F1} kW\n" +
                    "   Building Peak / Grid Peak:   {1:F2}%",
                    gridPeakKw, basePeakKw / gridPeakKw * 100.0)
                : "   Grid Peak Reference:         Not provided";

            // =============================================================
            // STEP 4: 合規報告
            // =============================================================
            string bessNote = batteryKwh > 0
                ? string.Format("   (Evaluated with {0:F1} kWh BESS, efficiency {1:P0})", batteryKwh, efficiency)
                : "   (No BESS - no active load management)";

            string report = string.Format(
                "=== LEED v4.1 Grid Harmonization - Case 3 ===\n" +
                "{0}\n" +
                "   Baseline: NetPurchased = max(0, Demand - PV)\n" +
                "   (Grid-facing demand, PV already factored into baseline)\n" +
                "   BESS contribution evaluated on top of this baseline.\n\n" +

                "A. Peak Shaving Analysis (10% Reduction Required):\n" +
                "   Grid Peak (without BESS):  {1:F3} kW\n" +
                "   Grid Peak (with BESS):     {2:F3} kW\n" +
                "   Peak Reduction:            {3:F1}%\n" +
                "   Result:                    {4}\n\n" +

                "B. 2-Hour Load Shifting (within Peak Hours):\n" +
                "   Peak Hours:                {5}\n" +
                "   Shifting Threshold:        {6:F3} kW (10% of {1:F3} kW grid peak)\n" +
                "   Max Consecutive Hours:     {7}\n" +
                "   Result:                    {8}\n\n" +

                "C. Grid Alignment Check:\n" +
                "   BESS Discharge in Peak Hours: {9} hrs ({10:F1}% of total discharge energy)\n" +
                "   Total BESS Discharge Hours:   {11}\n" +
                "{12}\n\n" +

                "=== Scoring (Credit Cap: 2 pts max) ===\n" +
                "   LEED v4.1 Grid Harmonization provides 4 pathways:\n" +
                "     1. Peak Load Optimization (reduce on-peak >= 10%)       +1 pt\n" +
                "     2. Flexible Operating Scenarios (shift 10% for 2 hrs)   +1 pt\n" +
                "     3. On-site thermal/electricity storage (reduce >= 10%)  +1 pt\n" +
                "     4. Grid resilience technologies (utility program)       +1 pt\n" +
                "   Projects select any combination; credit cap = 2 pts.\n\n" +

                "   This tool evaluates pathways 1 & 2 via BESS simulation:\n" +
                "   Peak Shaving (pathway 1):  {4,-8}  (+1 pt if PASS)\n" +
                "   Load Shifting (pathway 2): {8,-8}  (+1 pt if PASS)\n" +
                "   Estimated Points: {13} / 2 (max)\n" +
                "   Overall: {14}\n\n" +

                "=== System Prerequisites ===\n" +
                "   - Building must have permanent load management hardware\n" +
                "     (BESS, smart inverter, or BMS with DR capability)\n" +
                "   - System must be automated (no manual override for compliance)\n" +
                "   - Peak hours must align with regional utility TOU schedule",
                bessNote,                                                              // {0}
                basePeakKw,                                                            // {1}
                proposedPeakKw,                                                        // {2}
                peakReduction,                                                         // {3}
                peakPass ? "PASS" : "FAIL",                                            // {4}
                string.Join(", ", peakHours.Select(h => h.ToString("00") + ":00")),    // {5}
                shiftThresholdKw,                                                      // {6}
                maxConsecutiveInPeak,                                                  // {7}
                shiftPass ? "PASS" : "FAIL",                                           // {8}
                dischargeHoursInPeak,                                                  // {9}
                gridAlignmentPct,                                                      // {10}
                dischargeHoursTotal,                                                   // {11}
                gridAlignNote,                                                         // {12}
                points,                                                                // {13}
                (peakPass && shiftPass) ? "FULLY COMPLIANT" : (points > 0 ? "PARTIALLY COMPLIANT" : "NON-COMPLIANT")); // {14}

            // =============================================================
            // STEP 5: 計算方式說明 — 帶入實際數值
            // =============================================================
            double annualDischargeKwh = bessDischarge.Sum() / 1000.0;

            string methodology = string.Format(
                "=== 計算方式說明 (含代入數值) ===\n\n" +

                "【LEED v4.1 EA Credit: Grid Harmonization 學分架構】\n" +
                "  本學分提供 4 種得分途徑（任選，上限 2 分）：\n" +
                "  途徑 1: Peak Load Optimization — 削減尖峰負載 >= 10%         (+1 pt)\n" +
                "  途徑 2: Flexible Operating Scenarios — 轉移 10% 負載達 2 小時 (+1 pt)\n" +
                "  途徑 3: On-site Storage — 現場儲能削減尖峰 >= 10%             (+1 pt)\n" +
                "  途徑 4: Grid Resilience Technologies — 電力公司韌性計畫       (+1 pt)\n" +
                "  任選 1 項 = 1 分，任選 2 項 = 2 分 (滿分)。\n\n" +

                "  本工具透過 BESS 模擬評估途徑 1 (Peak Shaving) 及途徑 2 (Load Shifting)。\n" +
                "  途徑 3 與途徑 1 判定條件相同 (皆為削減 >= 10%)，由 BESS 達成時同時滿足。\n" +
                "  途徑 4 需人工確認電力公司是否提供韌性計畫，本工具不予評估。\n\n" +

                "基準線 = NetPurchased = max(0, Demand - PV)，即電網實際供電量。\n" +
                "PV 已反映在基準線中，不額外計分；僅評估 BESS 在此基準上的貢獻。\n\n" +

                "【輸入資料】\n" +
                "  電池容量   = {0:F1} kWh\n" +
                "  充放電效率 = {1:P0}\n" +
                "  尖峰時段   = {2}\n" +
                "  全年需求資料筆數 = {3} 筆\n" +
                "  全年 PV 資料筆數 = {4} 筆\n\n" +

                "【基準線計算】\n" +
                "  NetPurchased[i] = max(0, Demand[i] - PV[i])\n" +
                "  電網尖峰 (無 BESS) = max(NetPurchased[i]) = {5:F3} kW\n\n" +

                "【BESS 模擬】\n" +
                "  充電：餘電 = max(0, PV - Demand)，實際入電 = 餘電 x {1} (效率損耗)\n" +
                "  放電：缺電時，放電量 = min(NetPurchased[i], 目前電量)\n" +
                "  全年 BESS 總放電量 = {6:F1} kWh\n" +
                "  電網最終負載 GridLoad[i] = NetPurchased[i] - BESS_Discharge[i]\n\n" +

                "【評分 A：尖峰削減 (Peak Shaving)】\n" +
                "  電網尖峰 (無 BESS) = {5:F3} kW\n" +
                "  電網尖峰 (有 BESS) = max(GridLoad[i]) = {7:F3} kW\n" +
                "  削減率             = ({5:F3} - {7:F3}) / {5:F3} x 100%\n" +
                "                     = {8:F1}%\n" +
                "  通過門檻           >= 10%\n" +
                "  結果               = {9}\n\n" +

                "【評分 B：負荷轉移 (Load Shifting)】\n" +
                "  轉移門檻 = 電網尖峰 x 10% = {5:F3} x 0.10 = {10:F3} kW\n" +
                "  尖峰時段內最大連續放電時數 = {11} 小時\n" +
                "  (需連續 >= 2 小時且每小時放電 >= {10:F3} kW)\n" +
                "  通過門檻 >= 2 小時\n" +
                "  結果     = {12}\n\n" +

                "【評分 C：電網對接驗證 (Grid Alignment)】\n" +
                "  尖峰時段內 BESS 放電時數 = {13} 小時 (占全部放電 {14:F1}%)\n" +
                "  全部 BESS 放電時數       = {15} 小時\n" +
                "{16}\n\n" +

                "【總分計算】\n" +
                "  尖峰削減 {9} + 負荷轉移 {12}\n" +
                "  = {17} 分\n\n" +

                "參考依據：LEED v4.1 BD+C, EA Credit: Grid Harmonization, Case 3",
                batteryKwh,                                                            // {0}
                efficiency,                                                            // {1}
                string.Join(", ", peakHours.Select(h => h.ToString("00") + ":00")),    // {2}
                hours,                                                                 // {3}
                data.Produced.Count,                                                   // {4}
                basePeakKw,                                                            // {5}
                annualDischargeKwh,                                                    // {6}
                proposedPeakKw,                                                        // {7}
                peakReduction,                                                         // {8}
                peakPass ? "PASS" : "FAIL",                                            // {9}
                shiftThresholdKw,                                                      // {10}
                maxConsecutiveInPeak,                                                  // {11}
                shiftPass ? "PASS" : "FAIL",                                           // {12}
                dischargeHoursInPeak,                                                  // {13}
                gridAlignmentPct,                                                      // {14}
                dischargeHoursTotal,                                                   // {15}
                gridPeakKw > 0
                    ? string.Format("  建築尖峰 / 電網尖峰 = {0:F3} kW / {1:F1} kW = {2:F2}%",
                        basePeakKw, gridPeakKw, basePeakKw / gridPeakKw * 100.0)
                    : "  (未提供 Grid Peak kW)",                                       // {16}
                points                                                                 // {17}
            );

            // =============================================================
            // STEP 6: 最佳容量推薦 — 二分搜尋最低容量
            // =============================================================
            string recommendation;
            if (data.Produced.Count == 0)
            {
                recommendation = "無 PV 資料，無法計算推薦容量。";
            }
            else
            {
                double recShift = FindMinCapacity(netPurchased, surplus, efficiency, basePeakW,
                    peakHours, hours, checkPeakShaving: false, checkLoadShifting: true);
                double recPeak = FindMinCapacity(netPurchased, surplus, efficiency, basePeakW,
                    peakHours, hours, checkPeakShaving: true, checkLoadShifting: false);
                double recBoth = FindMinCapacity(netPurchased, surplus, efficiency, basePeakW,
                    peakHours, hours, checkPeakShaving: true, checkLoadShifting: true);

                recommendation = string.Format(
                    "=== 最佳 BESS 容量推薦 (效率 {0:P0}) ===\n\n" +
                    "  達到 1 分 (僅 Load Shifting):  {1}\n" +
                    "  達到 1 分 (僅 Peak Shaving):   {2}\n" +
                    "  達到 2 分 (兩者皆通過):        {3}\n\n" +
                    "  電網尖峰 (基準) = {4:F3} kW\n" +
                    "  10% 削減目標    = {5:F3} kW\n" +
                    "  轉移門檻        = {6:F3} kW (連續 2 小時)",
                    efficiency,
                    recShift > 0 ? string.Format("{0:F1} kWh", recShift) : "無法達成",
                    recPeak > 0 ? string.Format("{0:F1} kWh", recPeak) : "無法達成",
                    recBoth > 0 ? string.Format("{0:F1} kWh", recBoth) : "無法達成",
                    basePeakKw,
                    basePeakKw * 0.90,
                    shiftThresholdKw);
            }

            Message = points > 0 ? points + " pt" + (points > 1 ? "s" : "") : "0 pts";

            // --- 設定輸出 ---
            DA.SetData(0, points);
            DA.SetData(1, report);
            DA.SetData(2, peakReduction);
            DA.SetDataList(3, netPurchased.ToList());
            DA.SetDataList(4, bessDischarge.ToList());
            DA.SetDataList(5, gridLoad.ToList());
            DA.SetData(6, battInfo);
            DA.SetData(7, methodology);
            DA.SetData(8, recommendation);
        }

        // =================================================================
        // 二分搜尋：找到滿足條件的最低電池容量 (kWh)
        // =================================================================
        private double FindMinCapacity(double[] netPurchased, double[] surplus,
            double efficiency, double basePeakW, List<int> peakHours, int hours,
            bool checkPeakShaving, bool checkLoadShifting)
        {
            // 先用上限測試是否可行
            double maxCap = 500.0; // kWh 上限
            if (!TestCapacity(netPurchased, surplus, efficiency, basePeakW,
                peakHours, hours, maxCap, checkPeakShaving, checkLoadShifting))
                return -1; // 即使 500 kWh 也無法達成

            double lo = 0.0, hi = maxCap;
            // 二分搜尋，精確到 0.5 kWh
            while (hi - lo > 0.5)
            {
                double mid = (lo + hi) / 2.0;
                if (TestCapacity(netPurchased, surplus, efficiency, basePeakW,
                    peakHours, hours, mid, checkPeakShaving, checkLoadShifting))
                    hi = mid;
                else
                    lo = mid;
            }
            return Math.Ceiling(hi * 2) / 2.0; // 無條件進位到 0.5 kWh
        }

        private bool TestCapacity(double[] netPurchased, double[] surplus,
            double efficiency, double basePeakW, List<int> peakHours, int hours,
            double capacityKwh, bool checkPeakShaving, bool checkLoadShifting)
        {
            double capacityWh = capacityKwh * 1000;
            double currentCharge = 0;
            double gridLoadMax = 0;
            double shiftThresholdW = basePeakW * 0.10;
            int maxConsec = 0, consec = 0;

            for (int i = 0; i < hours; i++)
            {
                double discharge = 0;

                if (surplus[i] > 0)
                {
                    double eff = surplus[i] * efficiency;
                    double room = capacityWh - currentCharge;
                    currentCharge += Math.Min(eff, room);
                }
                else if (netPurchased[i] > 0 && currentCharge > 0)
                {
                    discharge = Math.Min(netPurchased[i], currentCharge);
                    currentCharge -= discharge;
                }

                double gl = netPurchased[i] - discharge;
                if (gl > gridLoadMax) gridLoadMax = gl;

                // Load shifting 計算
                int h = i % 24;
                if (peakHours.Contains(h))
                {
                    if (discharge >= shiftThresholdW)
                    {
                        consec++;
                        if (consec > maxConsec) maxConsec = consec;
                    }
                    else
                        consec = 0;
                }
                else
                    consec = 0;
            }

            bool peakOk = !checkPeakShaving || ((basePeakW - gridLoadMax) / basePeakW * 100.0 >= 10.0);
            bool shiftOk = !checkLoadShifting || (maxConsec >= 2);
            return peakOk && shiftOk;
        }

        // =================================================================
        // 4. Excel 讀取 (需要 TotalDemand 和 Produced)
        // =================================================================
        private class ExcelData
        {
            public List<string> Labels = new List<string>();
            public List<double> TotalDemand = new List<double>();
            public List<double> Produced = new List<double>();
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
                    new { Key = "Demand",   Match = (Func<string, bool>)(n => n.IndexOf("Electricity Dem", StringComparison.OrdinalIgnoreCase) >= 0) },
                    new { Key = "Produced", Match = (Func<string, bool>)(n => n.IndexOf("Produced", StringComparison.OrdinalIgnoreCase) >= 0) }
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
        protected override System.Drawing.Bitmap Icon
        {
            get
            {
                var assembly = System.Reflection.Assembly.GetExecutingAssembly();
                using (var stream = assembly.GetManifestResourceStream("Colugo.Resources.GridHarmony_Icon.png"))
                {
                    if (stream == null) return null;
                    return new System.Drawing.Bitmap(stream);
                }
            }
        }
        public override GH_Exposure Exposure => GH_Exposure.primary;
    }
}
