using System;
using System.Collections.Generic;
using System.Data.SQLite;
using System.IO;
using System.Linq;
using System.Text;
using Grasshopper.Kernel;

namespace Colugo.Components
{
    public class RenewableEnergyComponent : GH_Component
    {
        public RenewableEnergyComponent()
            : base("EA_Renewable Energy", "RE",
                "LEED v4.1 EA Credit: Renewable Energy evaluator.\n" +
                "Reads Honeybee SQL for annual energy & PV production,\n" +
                "calculates GHG offset % and points (max 5 + EP).",
                "Colugo", "LEED EA")
        { }

        // ═══════════════════════════════════════════════════════════════
        //  Inputs
        // ═══════════════════════════════════════════════════════════════
        protected override void RegisterInputParams(GH_InputParamManager pManager)
        {
            pManager.AddTextParameter("SQL Path", "_sql",
                "Path to Honeybee EnergyPlus output SQL file", GH_ParamAccess.item);                     // 0
            pManager.AddNumberParameter("CO2e Elec", "_co2Elec",
                "Electricity emission coefficient (tCO2e/kWh), e.g. 0.000502",
                GH_ParamAccess.item, 0.000502);                                                          // 1
            pManager.AddNumberParameter("CO2e Gas", "_co2Gas",
                "Natural gas emission coefficient (tCO2e/therm), e.g. 0.005302",
                GH_ParamAccess.item, 0.005302);                                                          // 2
            pManager.AddNumberParameter("System Efficiency", "_sysEff",
                "On-site RE system efficiency (0-1), accounts for transmission/conversion losses",
                GH_ParamAccess.item, 1.0);                                                               // 3
            pManager.AddNumberParameter("New Off-site kWh", "_newOffKwh",
                "Annual new off-site renewable energy purchase (kWh/yr)",
                GH_ParamAccess.item, 0.0);                                                               // 4
            pManager.AddNumberParameter("New Off-site Years", "_newOffYrs",
                "New off-site contract length (years, max 15)",
                GH_ParamAccess.item, 15.0);                                                              // 5
            pManager.AddNumberParameter("Existing Off-site kWh", "_exOffKwh",
                "Annual existing off-site renewable energy purchase (kWh/yr)",
                GH_ParamAccess.item, 0.0);                                                               // 6
            pManager.AddNumberParameter("Existing Off-site Years", "_exOffYrs",
                "Existing off-site contract length (years, max 15)",
                GH_ParamAccess.item, 15.0);                                                              // 7
            pManager.AddNumberParameter("Green-e EAC kWh", "_greeneKwh",
                "Annual Green-e certified EACs/carbon offsets purchase (kWh/yr)",
                GH_ParamAccess.item, 0.0);                                                               // 8
            pManager.AddNumberParameter("Green-e EAC Years", "_greeneYrs",
                "Green-e EACs contract length (years, max 15)",
                GH_ParamAccess.item, 15.0);                                                              // 9
            pManager.AddNumberParameter("Other EAC kWh", "_otherKwh",
                "Annual other EACs/carbon offsets purchase (kWh/yr)",
                GH_ParamAccess.item, 0.0);                                                               // 10
            pManager.AddNumberParameter("Other EAC Years", "_otherYrs",
                "Other EACs contract length (years, max 15)",
                GH_ParamAccess.item, 15.0);                                                              // 11
            pManager.AddBooleanParameter("Run", "_run",
                "Set True to execute", GH_ParamAccess.item, false);                                      // 12
        }

        // ═══════════════════════════════════════════════════════════════
        //  Outputs
        // ═══════════════════════════════════════════════════════════════
        protected override void RegisterOutputParams(GH_OutputParamManager pManager)
        {
            pManager.AddIntegerParameter("Total Points", "Pts",
                "Final Renewable Energy score (0-5, +EP)", GH_ParamAccess.item);                         // 0
            pManager.AddTextParameter("Points Breakdown", "Break",
                "Score breakdown by procurement category", GH_ParamAccess.item);                         // 1
            pManager.AddTextParameter("Calculation Log", "CalcLog",
                "Detailed calculation process (Chinese + English)", GH_ParamAccess.item);                // 2
            pManager.AddTextParameter("Methodology", "Method",
                "Methodology description (Chinese + English)", GH_ParamAccess.item);                     // 3
            pManager.AddNumberParameter("CO2eT", "CO2eT",
                "Total annual GHG emissions (tCO2e)", GH_ParamAccess.item);                              // 4
            pManager.AddNumberParameter("CO2eR On-site", "CO2eR_on",
                "On-site renewable GHG offset (tCO2e)", GH_ParamAccess.item);                            // 5
            pManager.AddBooleanParameter("EP Eligible", "EP",
                "True if 100% on-site offset (Exemplary Performance)", GH_ParamAccess.item);             // 6
            pManager.AddTextParameter("LEED Narrative", "Narr",
                "Auto-generated LEED narrative report", GH_ParamAccess.item);                            // 7
            pManager.AddTextParameter("Recommended PV", "RecPV",
                "Recommended PV production for each point level", GH_ParamAccess.item);                  // 8
        }

        // ═══════════════════════════════════════════════════════════════
        //  Table 1 thresholds
        // ═══════════════════════════════════════════════════════════════
        // Each array: index 0=1pt, 1=2pt, 2=3pt, 3=4pt, 4=5pt
        private static readonly double[] T_OnSite =      { 0.02, 0.06, 0.15, 0.35, 0.60 };
        private static readonly double[] T_NewOff =      { 0.20, 0.40, 0.60, 0.80, 1.00 };
        private static readonly double[] T_ExistOff =    { 0.60, 0.80, 1.00 };             // max 3 pts
        private static readonly double[] T_GreeneEAC =   { 1.00, 2.00, 3.00 };             // max 3 pts
        private static readonly double[] T_OtherEAC =    { 1.50 };                          // max 1 pt

        private static int LookupPoints(double pct, double[] thresholds)
        {
            int pts = 0;
            for (int i = 0; i < thresholds.Length; i++)
            {
                if (pct >= thresholds[i]) pts = i + 1;
                else break;
            }
            return pts;
        }

        // ═══════════════════════════════════════════════════════════════
        //  SolveInstance
        // ═══════════════════════════════════════════════════════════════
        protected override void SolveInstance(IGH_DataAccess DA)
        {
            string sqlPath = "";
            double co2Elec = 0.000502, co2Gas = 0.005302, sysEff = 1.0;
            double newOffKwh = 0, newOffYrs = 15;
            double exOffKwh = 0, exOffYrs = 15;
            double greeneKwh = 0, greeneYrs = 15;
            double otherKwh = 0, otherYrs = 15;
            bool run = false;

            DA.GetData(0, ref sqlPath);
            DA.GetData(1, ref co2Elec);
            DA.GetData(2, ref co2Gas);
            DA.GetData(3, ref sysEff);
            DA.GetData(4, ref newOffKwh);
            DA.GetData(5, ref newOffYrs);
            DA.GetData(6, ref exOffKwh);
            DA.GetData(7, ref exOffYrs);
            DA.GetData(8, ref greeneKwh);
            DA.GetData(9, ref greeneYrs);
            DA.GetData(10, ref otherKwh);
            DA.GetData(11, ref otherYrs);
            DA.GetData(12, ref run);

            if (!run)
            {
                DA.SetData(0, 0);
                DA.SetData(1, "Set Run = True to execute.");
                Message = "Idle";
                return;
            }

            if (string.IsNullOrWhiteSpace(sqlPath) || !File.Exists(sqlPath))
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "SQL file not found: " + sqlPath);
                return;
            }

            sysEff = Math.Max(0.01, Math.Min(1.0, sysEff));
            newOffYrs = Math.Max(1, Math.Min(15, newOffYrs));
            exOffYrs = Math.Max(1, Math.Min(15, exOffYrs));
            greeneYrs = Math.Max(1, Math.Min(15, greeneYrs));
            otherYrs = Math.Max(1, Math.Min(15, otherYrs));

            // ═══════════════════════════════════════════════════════════
            //  STEP 1: Read SQL — annual totals
            // ═══════════════════════════════════════════════════════════
            double annualElecKwh = 0;   // total electricity consumption
            double annualGasThm = 0;    // total natural gas (therms)
            double annualPvKwh = 0;     // on-site PV production

            try
            {
                string connStr = "Data Source=" + sqlPath + ";Version=3;Read Only=True;";

                // --- Electricity consumption ---
                // Try Rate variable (W) first, then meter (J)
                double? elecSum = ReadAnnualSum(connStr,
                    "Facility Total Electricity Demand Rate", false);
                if (elecSum.HasValue)
                {
                    // Rate in W, hourly → sum of W over 8760 hours / 1000 = kWh
                    annualElecKwh = elecSum.Value / 1000.0;
                }
                else
                {
                    elecSum = ReadAnnualSum(connStr, "Electricity:Facility", true);
                    if (elecSum.HasValue)
                        annualElecKwh = elecSum.Value / 3600000.0; // J → kWh
                }

                if (annualElecKwh <= 0)
                {
                    AddRuntimeMessage(GH_RuntimeMessageLevel.Error,
                        "No electricity consumption data found in SQL.");
                    return;
                }

                // --- Natural gas ---
                double? gasSum = ReadAnnualSum(connStr, "NaturalGas:Facility", true);
                if (gasSum.HasValue && gasSum.Value > 0)
                    annualGasThm = gasSum.Value / 105505585.0; // J → therms (1 therm = 105,505,585 J)

                // --- On-site PV production ---
                // Try Energy variable (J) first, then Rate (W)
                double? pvSum = ReadAnnualSum(connStr,
                    "Facility Total Produced Electricity Energy", false);
                if (pvSum.HasValue)
                {
                    annualPvKwh = pvSum.Value / 3600000.0; // J → kWh
                }
                else
                {
                    pvSum = ReadAnnualSum(connStr,
                        "Facility Total Produced Electricity Rate", false);
                    if (pvSum.HasValue)
                        annualPvKwh = pvSum.Value / 1000.0; // W sum → kWh
                }

                annualPvKwh *= sysEff; // apply system efficiency
            }
            catch (Exception ex)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "SQL read error: " + ex.Message);
                return;
            }

            // ═══════════════════════════════════════════════════════════
            //  STEP 2: Equation 1 — Total annual GHG emissions
            // ═══════════════════════════════════════════════════════════
            double co2eT = annualElecKwh * co2Elec + annualGasThm * co2Gas;

            if (co2eT <= 0)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error,
                    "Total GHG emissions = 0. Check emission coefficients.");
                return;
            }

            // ═══════════════════════════════════════════════════════════
            //  STEP 3: Equation 2 & 5 — GHG offset by each category
            //  Apply contract proration (Eq.5) for off-site purchases
            // ═══════════════════════════════════════════════════════════
            double co2eR_onsite = annualPvKwh * co2Elec;

            double proNewOff = newOffKwh * (newOffYrs / 15.0);
            double co2eR_newOff = proNewOff * co2Elec;

            double proExOff = exOffKwh * (exOffYrs / 15.0);
            double co2eR_exOff = proExOff * co2Elec;

            double proGreene = greeneKwh * (greeneYrs / 15.0);
            double co2eR_greene = proGreene * co2Elec;

            double proOther = otherKwh * (otherYrs / 15.0);
            double co2eR_other = proOther * co2Elec;

            // ═══════════════════════════════════════════════════════════
            //  STEP 4: Equation 3 — % offset (physical renewables)
            // ═══════════════════════════════════════════════════════════
            double pctOnsite = co2eR_onsite / co2eT;
            double pctNewOff = co2eR_newOff / co2eT;
            double pctExOff = co2eR_exOff / co2eT;

            // ═══════════════════════════════════════════════════════════
            //  STEP 5: Equation 4 — % offset (EACs, denominator deducted)
            // ═══════════════════════════════════════════════════════════
            double eacDenom = co2eT - co2eR_onsite - co2eR_newOff - co2eR_exOff;
            double pctGreene = eacDenom > 0 ? co2eR_greene / eacDenom : 0;
            double pctOther = eacDenom > 0 ? co2eR_other / eacDenom : 0;

            // ═══════════════════════════════════════════════════════════
            //  STEP 6: Table 1 — Points per category
            // ═══════════════════════════════════════════════════════════
            int ptsOnsite = LookupPoints(pctOnsite, T_OnSite);
            int ptsNewOff = LookupPoints(pctNewOff, T_NewOff);
            int ptsExOff = LookupPoints(pctExOff, T_ExistOff);
            int ptsGreene = LookupPoints(pctGreene, T_GreeneEAC);
            int ptsOther = LookupPoints(pctOther, T_OtherEAC);

            int totalPoints = Math.Min(5, ptsOnsite + ptsNewOff + ptsExOff + ptsGreene + ptsOther);
            bool isEP = pctOnsite >= 1.0;

            // ═══════════════════════════════════════════════════════════
            //  STEP 7: Points breakdown
            // ═══════════════════════════════════════════════════════════
            var brk = new StringBuilder();
            brk.AppendLine("=== LEED v4.1 Renewable Energy Score ===");
            brk.AppendLine(string.Format("Total Points: {0} / 5{1}", totalPoints, isEP ? " + EP" : ""));
            brk.AppendLine();
            brk.AppendLine(string.Format("  On-site Renewables:       {0,7:P1} → {1} pt(s)", pctOnsite, ptsOnsite));
            brk.AppendLine(string.Format("  New Off-site Renewables:  {0,7:P1} → {1} pt(s)", pctNewOff, ptsNewOff));
            brk.AppendLine(string.Format("  Existing Off-site:        {0,7:P1} → {1} pt(s)  (cap 3)", pctExOff, ptsExOff));
            brk.AppendLine(string.Format("  Green-e EACs/Offsets:     {0,7:P1} → {1} pt(s)  (cap 3)", pctGreene, ptsGreene));
            brk.AppendLine(string.Format("  Other EACs/Offsets:       {0,7:P1} → {1} pt(s)  (cap 1)", pctOther, ptsOther));
            brk.AppendLine();
            brk.AppendLine(string.Format("  Sum = {0}+{1}+{2}+{3}+{4} = {5}, capped at min(5, {5}) = {6}",
                ptsOnsite, ptsNewOff, ptsExOff, ptsGreene, ptsOther,
                ptsOnsite + ptsNewOff + ptsExOff + ptsGreene + ptsOther, totalPoints));
            if (isEP)
                brk.AppendLine("  ★ Exemplary Performance: 100% on-site offset achieved!");

            // ═══════════════════════════════════════════════════════════
            //  STEP 8: Calculation Log (中英文)
            // ═══════════════════════════════════════════════════════════
            string calcLog = BuildCalcLog(
                annualElecKwh, annualGasThm, annualPvKwh, sysEff,
                co2Elec, co2Gas, co2eT,
                co2eR_onsite, co2eR_newOff, co2eR_exOff, co2eR_greene, co2eR_other,
                pctOnsite, pctNewOff, pctExOff, pctGreene, pctOther,
                eacDenom,
                newOffKwh, newOffYrs, exOffKwh, exOffYrs,
                greeneKwh, greeneYrs, otherKwh, otherYrs,
                ptsOnsite, ptsNewOff, ptsExOff, ptsGreene, ptsOther,
                totalPoints, isEP);

            // ═══════════════════════════════════════════════════════════
            //  STEP 9: Methodology (中英文)
            // ═══════════════════════════════════════════════════════════
            string methodology = BuildMethodology(co2Elec, co2Gas);

            // ═══════════════════════════════════════════════════════════
            //  STEP 10: Narrative
            // ═══════════════════════════════════════════════════════════
            string narrative = BuildNarrative(
                annualElecKwh, annualGasThm, annualPvKwh, co2eT,
                co2eR_onsite, pctOnsite, ptsOnsite,
                co2eR_newOff, pctNewOff, ptsNewOff, newOffKwh,
                co2eR_exOff, pctExOff, ptsExOff,
                co2eR_greene, pctGreene, ptsGreene,
                co2eR_other, pctOther, ptsOther,
                totalPoints, isEP);

            // ═══════════════════════════════════════════════════════════
            //  STEP 11: Recommended PV (中英文)
            // ═══════════════════════════════════════════════════════════
            string recPV = BuildRecommendation(co2eT, co2Elec, sysEff, annualPvKwh);

            // ═══════════════════════════════════════════════════════════
            //  Set outputs
            // ═══════════════════════════════════════════════════════════
            Message = totalPoints + " pt" + (totalPoints != 1 ? "s" : "") + (isEP ? "+EP" : "");

            DA.SetData(0, totalPoints);
            DA.SetData(1, brk.ToString());
            DA.SetData(2, calcLog);
            DA.SetData(3, methodology);
            DA.SetData(4, co2eT);
            DA.SetData(5, co2eR_onsite);
            DA.SetData(6, isEP);
            DA.SetData(7, narrative);
            DA.SetData(8, recPV);
        }

        // ═══════════════════════════════════════════════════════════════
        //  SQL helper — read annual sum of a variable
        // ═══════════════════════════════════════════════════════════════
        private double? ReadAnnualSum(string connStr, string varName, bool isMeter)
        {
            using (var conn = new SQLiteConnection(connStr))
            {
                conn.Open();
                string where = isMeter
                    ? "rdd.Name = @name AND rdd.IsMeter = 1"
                    : (varName.Contains("%")
                        ? "rdd.Name LIKE @name AND rdd.IsMeter = 0"
                        : "rdd.Name = @name AND rdd.IsMeter = 0");

                string sql = string.Format(@"
                    SELECT SUM(rd.Value)
                    FROM ReportData rd
                    JOIN ReportDataDictionary rdd
                        ON rd.ReportDataDictionaryIndex = rdd.ReportDataDictionaryIndex
                    JOIN Time t ON rd.TimeIndex = t.TimeIndex
                    WHERE {0}
                      AND rdd.ReportingFrequency = 'Hourly'
                      AND (t.WarmupFlag = 0 OR t.WarmupFlag IS NULL)
                      AND t.DayType NOT LIKE '%DesignDay%'", where);

                using (var cmd = new SQLiteCommand(sql, conn))
                {
                    cmd.Parameters.AddWithValue("@name", varName);
                    var result = cmd.ExecuteScalar();
                    if (result == null || result is DBNull) return null;
                    double val = Convert.ToDouble(result);
                    return val > 0 ? val : (double?)null;
                }
            }
        }

        // ═══════════════════════════════════════════════════════════════
        //  Calculation Log (中英文)
        // ═══════════════════════════════════════════════════════════════
        private string BuildCalcLog(
            double elecKwh, double gasThm, double pvKwh, double sysEff,
            double co2Elec, double co2Gas, double co2eT,
            double co2eR_on, double co2eR_newOff, double co2eR_exOff,
            double co2eR_greene, double co2eR_other,
            double pctOn, double pctNewOff, double pctExOff, double pctGreene, double pctOther,
            double eacDenom,
            double newOffKwh, double newOffYrs, double exOffKwh, double exOffYrs,
            double greeneKwh, double greeneYrs, double otherKwh, double otherYrs,
            int ptsOn, int ptsNewOff, int ptsExOff, int ptsGreene, int ptsOther,
            int totalPts, bool isEP)
        {
            var sb = new StringBuilder();

            // ─── 中文 ───
            sb.AppendLine("=== 計算過程 ===");
            sb.AppendLine();
            sb.AppendLine("【Equation 1：建築年度總碳排】");
            sb.AppendLine(string.Format("  年用電量 = {0:N0} kWh", elecKwh));
            if (gasThm > 0)
                sb.AppendLine(string.Format("  年用氣量 = {0:N1} therms", gasThm));
            sb.AppendLine(string.Format("  電力碳排係數 = {0} tCO2e/kWh", co2Elec));
            if (gasThm > 0)
                sb.AppendLine(string.Format("  天然氣碳排係數 = {0} tCO2e/therm", co2Gas));
            sb.AppendLine(string.Format("  CO2eT = {0:N0} × {1} + {2:N1} × {3}",
                elecKwh, co2Elec, gasThm, co2Gas));
            sb.AppendLine(string.Format("        = {0:F4} tCO2e", co2eT));
            sb.AppendLine();

            sb.AppendLine("【Equation 2 & 5：各類再生能源碳排抵消】");
            sb.AppendLine(string.Format("  On-site PV 年產電 = {0:N0} kWh（系統效率 {1:P0}）", pvKwh, sysEff));
            sb.AppendLine(string.Format("    CO2eR_onsite = {0:N0} × {1} = {2:F4} tCO2e", pvKwh, co2Elec, co2eR_on));
            if (newOffKwh > 0)
            {
                sb.AppendLine(string.Format("  New Off-site = {0:N0} kWh × ({1:F0}/15) = {2:N0} kWh（折算後）",
                    newOffKwh, newOffYrs, newOffKwh * newOffYrs / 15.0));
                sb.AppendLine(string.Format("    CO2eR_newOff = {0:F4} tCO2e", co2eR_newOff));
            }
            if (exOffKwh > 0)
            {
                sb.AppendLine(string.Format("  Existing Off-site = {0:N0} kWh × ({1:F0}/15) = {2:N0} kWh",
                    exOffKwh, exOffYrs, exOffKwh * exOffYrs / 15.0));
                sb.AppendLine(string.Format("    CO2eR_exOff = {0:F4} tCO2e", co2eR_exOff));
            }
            if (greeneKwh > 0)
            {
                sb.AppendLine(string.Format("  Green-e EACs = {0:N0} kWh × ({1:F0}/15) = {2:N0} kWh",
                    greeneKwh, greeneYrs, greeneKwh * greeneYrs / 15.0));
                sb.AppendLine(string.Format("    CO2eR_greene = {0:F4} tCO2e", co2eR_greene));
            }
            if (otherKwh > 0)
            {
                sb.AppendLine(string.Format("  Other EACs = {0:N0} kWh × ({1:F0}/15) = {2:N0} kWh",
                    otherKwh, otherYrs, otherKwh * otherYrs / 15.0));
                sb.AppendLine(string.Format("    CO2eR_other = {0:F4} tCO2e", co2eR_other));
            }
            sb.AppendLine();

            sb.AppendLine("【Equation 3：物理再生能源抵消百分比】");
            sb.AppendLine(string.Format("  On-site:      {0:F4} / {1:F4} = {2:P2}", co2eR_on, co2eT, pctOn));
            sb.AppendLine(string.Format("  New Off-site: {0:F4} / {1:F4} = {2:P2}", co2eR_newOff, co2eT, pctNewOff));
            sb.AppendLine(string.Format("  Exist Off:    {0:F4} / {1:F4} = {2:P2}", co2eR_exOff, co2eT, pctExOff));
            sb.AppendLine();

            sb.AppendLine("【Equation 4：EACs 抵消百分比（分母已扣除物理再生能源）】");
            sb.AppendLine(string.Format("  分母 = CO2eT - On - New - Exist = {0:F4} - {1:F4} - {2:F4} - {3:F4} = {4:F4}",
                co2eT, co2eR_on, co2eR_newOff, co2eR_exOff, eacDenom));
            sb.AppendLine(string.Format("  Green-e: {0:F4} / {1:F4} = {2:P2}", co2eR_greene, eacDenom, pctGreene));
            sb.AppendLine(string.Format("  Other:   {0:F4} / {1:F4} = {2:P2}", co2eR_other, eacDenom, pctOther));
            sb.AppendLine();

            sb.AppendLine("【Table 1 查表得分】");
            sb.AppendLine(string.Format("  On-site {0:P1} → {1} 分", pctOn, ptsOn));
            sb.AppendLine(string.Format("  New Off-site {0:P1} → {1} 分", pctNewOff, ptsNewOff));
            sb.AppendLine(string.Format("  Existing Off-site {0:P1} → {1} 分（上限 3）", pctExOff, ptsExOff));
            sb.AppendLine(string.Format("  Green-e EACs {0:P1} → {1} 分（上限 3）", pctGreene, ptsGreene));
            sb.AppendLine(string.Format("  Other EACs {0:P1} → {1} 分（上限 1）", pctOther, ptsOther));
            sb.AppendLine(string.Format("  總分 = min(5, {0}+{1}+{2}+{3}+{4}) = {5}",
                ptsOn, ptsNewOff, ptsExOff, ptsGreene, ptsOther, totalPts));
            if (isEP) sb.AppendLine("  ★ Exemplary Performance：On-site 100% 抵消！");
            sb.AppendLine();

            // ─── English ───
            sb.AppendLine("=== Calculation Log ===");
            sb.AppendLine();
            sb.AppendLine("[Equation 1: Total Annual GHG Emissions]");
            sb.AppendLine(string.Format("  Annual electricity = {0:N0} kWh", elecKwh));
            if (gasThm > 0)
                sb.AppendLine(string.Format("  Annual natural gas = {0:N1} therms", gasThm));
            sb.AppendLine(string.Format("  CO2eT = {0:F4} tCO2e", co2eT));
            sb.AppendLine();

            sb.AppendLine("[Equation 2 & 5: GHG Offset by Category]");
            sb.AppendLine(string.Format("  On-site PV = {0:N0} kWh (eff {1:P0}) -> CO2eR = {2:F4} tCO2e",
                pvKwh, sysEff, co2eR_on));
            if (newOffKwh > 0)
                sb.AppendLine(string.Format("  New Off-site = {0:N0} kWh x ({1:F0}/15) -> CO2eR = {2:F4} tCO2e",
                    newOffKwh, newOffYrs, co2eR_newOff));
            if (exOffKwh > 0)
                sb.AppendLine(string.Format("  Existing Off-site = {0:N0} kWh x ({1:F0}/15) -> CO2eR = {2:F4} tCO2e",
                    exOffKwh, exOffYrs, co2eR_exOff));
            if (greeneKwh > 0)
                sb.AppendLine(string.Format("  Green-e EACs = {0:N0} kWh x ({1:F0}/15) -> CO2eR = {2:F4} tCO2e",
                    greeneKwh, greeneYrs, co2eR_greene));
            if (otherKwh > 0)
                sb.AppendLine(string.Format("  Other EACs = {0:N0} kWh x ({1:F0}/15) -> CO2eR = {2:F4} tCO2e",
                    otherKwh, otherYrs, co2eR_other));
            sb.AppendLine();

            sb.AppendLine("[Equation 3: Physical Renewable Offset %]");
            sb.AppendLine(string.Format("  On-site: {0:P2}, New Off: {1:P2}, Exist Off: {2:P2}",
                pctOn, pctNewOff, pctExOff));
            sb.AppendLine();

            sb.AppendLine("[Equation 4: EACs Offset % (denominator deducted)]");
            sb.AppendLine(string.Format("  Denominator = {0:F4} tCO2e", eacDenom));
            sb.AppendLine(string.Format("  Green-e: {0:P2}, Other: {1:P2}", pctGreene, pctOther));
            sb.AppendLine();

            sb.AppendLine("[Table 1 Scoring]");
            sb.AppendLine(string.Format("  On-site {0:P1} -> {1} pt, New Off {2:P1} -> {3} pt, Exist Off {4:P1} -> {5} pt",
                pctOn, ptsOn, pctNewOff, ptsNewOff, pctExOff, ptsExOff));
            sb.AppendLine(string.Format("  Green-e {0:P1} -> {1} pt (cap 3), Other {2:P1} -> {3} pt (cap 1)",
                pctGreene, ptsGreene, pctOther, ptsOther));
            sb.AppendLine(string.Format("  Total = min(5, {0}) = {1}{2}",
                ptsOn + ptsNewOff + ptsExOff + ptsGreene + ptsOther, totalPts,
                isEP ? " + Exemplary Performance" : ""));

            return sb.ToString();
        }

        // ═══════════════════════════════════════════════════════════════
        //  Methodology (中英文)
        // ═══════════════════════════════════════════════════════════════
        private string BuildMethodology(double co2Elec, double co2Gas)
        {
            var sb = new StringBuilder();

            sb.AppendLine("=== 方法論說明 ===");
            sb.AppendLine();
            sb.AppendLine("【分析依據】");
            sb.AppendLine("  本分析依據 LEED v4.1 BD+C「EA Credit: Renewable Energy」進行。");
            sb.AppendLine("  該學分旨在透過增加現場自給與電網端碳緩解項目，降低建築溫室氣體排放。");
            sb.AppendLine("  學分上限 5 分，若 On-site 達 100% 抵消可獲 Exemplary Performance。");
            sb.AppendLine();
            sb.AppendLine("【計算方法】");
            sb.AppendLine("  Equation 1: CO2eT = Σ(EnergySource_i × CO2eCoeff_i)");
            sb.AppendLine("    計算建築年度總碳排（電力 + 天然氣）。");
            sb.AppendLine("  Equation 2: CO2eR = Σ(RenewableSource_i × CO2eCoeff_i)");
            sb.AppendLine("    分別計算各類再生能源的碳排抵消量。");
            sb.AppendLine("  Equation 3: %offset = CO2eR / CO2eT");
            sb.AppendLine("    物理再生能源的抵消百分比，對照 Table 1 查得分。");
            sb.AppendLine("  Equation 4: %EAC = CO2eR_EAC / (CO2eT - CO2eR_on - CO2eR_new - CO2eR_exist)");
            sb.AppendLine("    EACs 的分母須扣除物理再生能源的貢獻，避免重複計算。");
            sb.AppendLine("  Equation 5: Equivalent = Annual × (合約年限 / 15)");
            sb.AppendLine("    不足 15 年的合約按比例折算。");
            sb.AppendLine();
            sb.AppendLine("【資料來源】");
            sb.AppendLine("  建築用電與現場產電：Honeybee EnergyPlus SQL 模擬輸出。");
            sb.AppendLine(string.Format("  碳排係數：電力 {0} tCO2e/kWh，天然氣 {1} tCO2e/therm。", co2Elec, co2Gas));
            sb.AppendLine("  離岸再生能源與 EACs：使用者手動輸入（合約/採購行為）。");
            sb.AppendLine();

            sb.AppendLine("=== Methodology ===");
            sb.AppendLine();
            sb.AppendLine("[Analysis Basis]");
            sb.AppendLine("  This analysis follows LEED v4.1 BD+C 'EA Credit: Renewable Energy'.");
            sb.AppendLine("  The credit aims to reduce GHG emissions through on-site renewables,");
            sb.AppendLine("  off-site procurement, EACs, and carbon offsets. Max 5 pts + EP.");
            sb.AppendLine();
            sb.AppendLine("[Calculation Method]");
            sb.AppendLine("  Eq.1: CO2eT = sum(EnergySource_i x CO2eCoeff_i)");
            sb.AppendLine("  Eq.2: CO2eR = sum(RenewableSource_i x CO2eCoeff_i) per category");
            sb.AppendLine("  Eq.3: %offset = CO2eR / CO2eT (for physical renewables)");
            sb.AppendLine("  Eq.4: %EAC = CO2eR_EAC / (CO2eT - on - new - exist) (denominator deducted)");
            sb.AppendLine("  Eq.5: Equivalent = Annual x (ContractYears / 15) (proration)");
            sb.AppendLine();
            sb.AppendLine("[Data Sources]");
            sb.AppendLine("  Building energy & PV: Honeybee EnergyPlus SQL output.");
            sb.AppendLine(string.Format("  Emission coefficients: electricity {0} tCO2e/kWh, gas {1} tCO2e/therm.", co2Elec, co2Gas));
            sb.AppendLine("  Off-site renewables & EACs: user-provided (contract/procurement data).");

            return sb.ToString();
        }

        // ═══════════════════════════════════════════════════════════════
        //  LEED Narrative
        // ═══════════════════════════════════════════════════════════════
        private string BuildNarrative(
            double elecKwh, double gasThm, double pvKwh, double co2eT,
            double co2eR_on, double pctOn, int ptsOn,
            double co2eR_newOff, double pctNewOff, int ptsNewOff, double newOffKwh,
            double co2eR_exOff, double pctExOff, int ptsExOff,
            double co2eR_greene, double pctGreene, int ptsGreene,
            double co2eR_other, double pctOther, int ptsOther,
            int totalPts, bool isEP)
        {
            var sb = new StringBuilder();
            sb.AppendLine("LEED v4.1 EA Credit: Renewable Energy - Narrative Report");
            sb.AppendLine("========================================================");
            sb.AppendLine();

            sb.AppendLine("1. Annual Building Greenhouse Gas Emissions");
            sb.AppendLine("-------------------------------------------");
            sb.AppendLine(string.Format(
                "Based on the EnergyPlus whole-building simulation, the building's annual electricity " +
                "consumption is {0:N0} kWh{1}. Using the emission coefficients from EA Prerequisite " +
                "Minimum Energy Performance, the total annual greenhouse gas emissions (CO2eT) are " +
                "estimated at {2:F2} metric tons of CO2e.",
                elecKwh,
                gasThm > 0 ? string.Format(" and natural gas consumption is {0:N0} therms", gasThm) : "",
                co2eT));
            sb.AppendLine();

            sb.AppendLine("2. Renewable Energy Procurement Strategies");
            sb.AppendLine("------------------------------------------");

            if (pvKwh > 0)
                sb.AppendLine(string.Format(
                    "On-site Renewables: A photovoltaic system produces {0:N0} kWh/year of usable energy, " +
                    "offsetting {1:F2} tCO2e ({2:P1} of total emissions). This achieves {3} point(s).",
                    pvKwh, co2eR_on, pctOn, ptsOn));

            if (newOffKwh > 0)
                sb.AppendLine(string.Format(
                    "New Off-site Renewables: {0:N0} kWh/year procured, offsetting {1:P1} of emissions " +
                    "for {2} point(s).", newOffKwh, pctNewOff, ptsNewOff));

            if (co2eR_exOff > 0)
                sb.AppendLine(string.Format(
                    "Existing Off-site Renewables: Procurement offsets {0:P1} of emissions for {1} point(s).",
                    pctExOff, ptsExOff));

            if (co2eR_greene > 0)
                sb.AppendLine(string.Format(
                    "Green-e Certified EACs/Carbon Offsets: Purchase offsets {0:P1} of remaining emissions " +
                    "for {1} point(s).", pctGreene, ptsGreene));

            if (co2eR_other > 0)
                sb.AppendLine(string.Format(
                    "Other EACs/Carbon Offsets: Purchase offsets {0:P1} of remaining emissions for {1} point(s).",
                    pctOther, ptsOther));

            sb.AppendLine();
            sb.AppendLine("3. Results Summary");
            sb.AppendLine("------------------");
            sb.AppendLine(string.Format(
                "The project achieves {0} point(s) under EA Credit: Renewable Energy.{1}",
                totalPts, isEP ? " Additionally, the project qualifies for Exemplary Performance " +
                "(100% of emissions offset through on-site renewable energy)." : ""));

            return sb.ToString();
        }

        // ═══════════════════════════════════════════════════════════════
        //  Recommended PV (中英文)
        // ═══════════════════════════════════════════════════════════════
        private string BuildRecommendation(double co2eT, double co2Elec, double sysEff, double currentPvKwh)
        {
            var sb = new StringBuilder();
            double currentPct = co2Elec > 0 ? (currentPvKwh * co2Elec) / co2eT : 0;

            sb.AppendLine("=== 建議 PV 產電量（On-site Renewables）===");
            sb.AppendLine();
            sb.AppendLine(string.Format("  建築年度碳排 CO2eT = {0:F2} tCO2e", co2eT));
            sb.AppendLine(string.Format("  目前 PV 產電 = {0:N0} kWh/yr（抵消 {1:P1}）", currentPvKwh, currentPct));
            sb.AppendLine(string.Format("  系統效率 = {0:P0}", sysEff));
            sb.AppendLine();

            for (int pts = 1; pts <= 5; pts++)
            {
                double threshold = T_OnSite[pts - 1];
                double neededCO2eR = co2eT * threshold;
                double neededKwh = co2Elec > 0 ? neededCO2eR / co2Elec / sysEff : 0;
                string status = currentPvKwh >= neededKwh ? "✓ 已達成" : string.Format("需增加 {0:N0} kWh", neededKwh - currentPvKwh);
                sb.AppendLine(string.Format("  {0} 分（≥{1:P0}）：需要 {2:N0} kWh/yr → {3}",
                    pts, threshold, neededKwh, status));
            }

            double epKwh = co2Elec > 0 ? co2eT / co2Elec / sysEff : 0;
            string epStatus = currentPvKwh >= epKwh ? "✓ 已達成" : string.Format("需增加 {0:N0} kWh", epKwh - currentPvKwh);
            sb.AppendLine(string.Format("  EP（100%）：需要 {0:N0} kWh/yr → {1}", epKwh, epStatus));

            sb.AppendLine();
            sb.AppendLine("=== Recommended PV Production (On-site) ===");
            sb.AppendLine();
            sb.AppendLine(string.Format("  CO2eT = {0:F2} tCO2e, current PV = {1:N0} kWh/yr ({2:P1})",
                co2eT, currentPvKwh, currentPct));
            sb.AppendLine();

            for (int pts = 1; pts <= 5; pts++)
            {
                double threshold = T_OnSite[pts - 1];
                double neededKwh = co2Elec > 0 ? co2eT * threshold / co2Elec / sysEff : 0;
                string status = currentPvKwh >= neededKwh ? "achieved" : string.Format("need +{0:N0} kWh", neededKwh - currentPvKwh);
                sb.AppendLine(string.Format("  {0} pt (>={1:P0}): {2:N0} kWh/yr -> {3}",
                    pts, threshold, neededKwh, status));
            }

            double epKwhEn = co2Elec > 0 ? co2eT / co2Elec / sysEff : 0;
            string epStatusEn = currentPvKwh >= epKwhEn ? "achieved" : string.Format("need +{0:N0} kWh", epKwhEn - currentPvKwh);
            sb.AppendLine(string.Format("  EP (100%): {0:N0} kWh/yr -> {1}", epKwhEn, epStatusEn));

            return sb.ToString();
        }

        // ═══════════════════════════════════════════════════════════════
        //  Component identity
        // ═══════════════════════════════════════════════════════════════
        public override Guid ComponentGuid => new Guid("b3d72e18-9f5a-4c83-a6b1-7e4d2f8c0a95");

        protected override System.Drawing.Bitmap Icon
        {
            get
            {
                var asm = System.Reflection.Assembly.GetExecutingAssembly();
                using (var stream = asm.GetManifestResourceStream("Colugo.Resources.GridHarmony_Icon.png"))
                {
                    if (stream == null) return null;
                    return new System.Drawing.Bitmap(stream);
                }
            }
        }

        public override GH_Exposure Exposure => GH_Exposure.primary;
    }
}
