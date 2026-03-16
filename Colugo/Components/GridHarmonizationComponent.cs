using System;
using System.Collections.Generic;
using System.Data.SQLite;
using System.IO;
using System.Linq;
using System.Text;
using Grasshopper.Kernel;
using Grasshopper.Kernel.Data;
using Grasshopper.Kernel.Types;

namespace Colugo.Components
{
    public class GridHarmonizationComponent : GH_Component
    {
        public GridHarmonizationComponent()
            : base("EA_Grid Harmonization", "GridHarm",
                "LEED v4.1 EA Credit: Grid Harmonization evaluator.\n" +
                "Reads Honeybee EnergyPlus SQL, evaluates Case 1/2/3 scoring\n" +
                "with battery storage and load management support.",
                "Colugo", "LEED EA")
        { }

        // ═══════════════════════════════════════════════════════════════
        //  Inputs
        // ═══════════════════════════════════════════════════════════════
        protected override void RegisterInputParams(GH_InputParamManager pManager)
        {
            pManager.AddTextParameter("SQL Path", "_sql",
                "Path to Honeybee EnergyPlus output SQL file", GH_ParamAccess.item);                    // 0
            pManager.AddIntegerParameter("Peak Start", "_peakStart",
                "On-peak start hour (0-23)", GH_ParamAccess.item, 13);                                  // 1
            pManager.AddIntegerParameter("Peak End", "_peakEnd",
                "On-peak end hour (0-23)", GH_ParamAccess.item, 17);                                    // 2
            pManager.AddIntegerParameter("Peak Days", "_peakDays",
                "Peak days (0=Mon..6=Sun), default Mon-Fri", GH_ParamAccess.list);                      // 3
            pManager.AddNumberParameter("Cooling COP", "_copC",
                "Cooling system COP for Ideal Loads conversion", GH_ParamAccess.item, 3.0);             // 4
            pManager.AddNumberParameter("Heating COP", "_copH",
                "Heating system COP for Ideal Loads conversion", GH_ParamAccess.item, 1.0);             // 5
            pManager.AddNumberParameter("Shed Schedule", "_shed",
                "Load shed kW per hour (8760 values, or 24 for daily pattern, or 1 for uniform on-peak)",
                GH_ParamAccess.list);                                                                    // 6
            pManager.AddNumberParameter("Proposed Schedule", "_proposed",
                "Proposed demand kW (8760 values) for Flexible Operating evaluation",
                GH_ParamAccess.list);                                                                    // 7
            pManager.AddNumberParameter("Battery kWh", "_battKwh",
                "Battery energy capacity in kWh (0 = no battery)", GH_ParamAccess.item, 0.0);           // 8
            pManager.AddNumberParameter("Charge Rate kW", "_chgKw",
                "Max battery charge rate in kW", GH_ParamAccess.item, 50.0);                            // 9
            pManager.AddNumberParameter("Discharge Rate kW", "_disKw",
                "Max battery discharge rate in kW", GH_ParamAccess.item, 50.0);                         // 10
            pManager.AddNumberParameter("Battery Efficiency", "_eff",
                "Round-trip efficiency (0-1)", GH_ParamAccess.item, 0.90);                              // 11
            pManager.AddNumberParameter("SOC Min", "_socMin",
                "Minimum state of charge (0-1)", GH_ParamAccess.item, 0.1);                             // 12
            pManager.AddNumberParameter("SOC Max", "_socMax",
                "Maximum state of charge (0-1)", GH_ParamAccess.item, 0.9);                             // 13
            pManager.AddBooleanParameter("Case 1", "_case1",
                "Enrolled in DR program (2 pts)", GH_ParamAccess.item, false);                          // 14
            pManager.AddBooleanParameter("Case 2", "_case2",
                "DR capable building (1 pt)", GH_ParamAccess.item, false);                              // 15
            pManager.AddBooleanParameter("Grid Resilience", "_resil",
                "Utility has grid resilience program (Case 3, 1 pt)", GH_ParamAccess.item, false);      // 16
            pManager.AddBooleanParameter("Run", "_run",
                "Set True to execute", GH_ParamAccess.item, false);                                     // 17

            pManager[3].Optional = true;   // peak days
            pManager[6].Optional = true;   // shed schedule
            pManager[7].Optional = true;   // proposed schedule
        }

        // ═══════════════════════════════════════════════════════════════
        //  Outputs
        // ═══════════════════════════════════════════════════════════════
        protected override void RegisterOutputParams(GH_OutputParamManager pManager)
        {
            pManager.AddIntegerParameter("Total Points", "Pts",
                "Final Grid Harmonization score (0-2)", GH_ParamAccess.item);                            // 0
            pManager.AddTextParameter("Points Breakdown", "Break",
                "Score breakdown by Case and strategy", GH_ParamAccess.item);                            // 1
            pManager.AddTextParameter("Calculation Log", "CalcLog",
                "Detailed calculation process (Chinese + English)", GH_ParamAccess.item);                // 2
            pManager.AddTextParameter("Methodology", "Method",
                "Methodology description (Chinese + English)", GH_ParamAccess.item);                     // 3
            pManager.AddNumberParameter("Baseline Load", "Base",
                "8760 hourly baseline electrical load (kW)", GH_ParamAccess.list);                       // 4
            pManager.AddNumberParameter("Net Load", "Net",
                "8760 hourly net load after PV deduction (kW)", GH_ParamAccess.list);                    // 5
            pManager.AddNumberParameter("Modified Load", "Mod",
                "8760 hourly load after all strategies (kW)", GH_ParamAccess.list);                      // 6
            pManager.AddNumberParameter("Battery SOC", "SOC",
                "8760 hourly battery state of charge (0-1)", GH_ParamAccess.list);                       // 7
            pManager.AddNumberParameter("Peak Reduction %", "PeakRed",
                "On-peak demand reduction percentage", GH_ParamAccess.item);                             // 8
            pManager.AddNumberParameter("Annual Chart", "Annual",
                "Annual load curves: {0}=baseline, {1}=modified (8760 kW)", GH_ParamAccess.tree);       // 9
            pManager.AddNumberParameter("Typical Day", "TypDay",
                "Typical day profiles: {0}=WD base, {1}=WE base, {2}=WD mod, {3}=WE mod (24 kW)",
                GH_ParamAccess.tree);                                                                    // 10
            pManager.AddTextParameter("LEED Narrative", "Narr",
                "Auto-generated LEED Case 3 narrative report", GH_ParamAccess.item);                     // 11
            pManager.AddBooleanParameter("Secondary Peak Warning", "Warn",
                "True if strategies create a new higher peak", GH_ParamAccess.item);                     // 12
            pManager.AddTextParameter("Recommended BESS", "RecBESS",
                "Recommended minimum battery capacity for 1pt and 2pts", GH_ParamAccess.item);           // 13
        }

        // ═══════════════════════════════════════════════════════════════
        //  Internal data structures
        // ═══════════════════════════════════════════════════════════════
        private class HourRecord
        {
            public int Month, Day, Hour;
            public string DayType;
        }

        // Map EnergyPlus DayType string → 0=Mon..6=Sun
        private static int DayTypeToIndex(string dayType)
        {
            if (dayType == null) return -1;
            string dt = dayType.Trim();
            if (dt.StartsWith("Monday", StringComparison.OrdinalIgnoreCase)) return 0;
            if (dt.StartsWith("Tuesday", StringComparison.OrdinalIgnoreCase)) return 1;
            if (dt.StartsWith("Wednesday", StringComparison.OrdinalIgnoreCase)) return 2;
            if (dt.StartsWith("Thursday", StringComparison.OrdinalIgnoreCase)) return 3;
            if (dt.StartsWith("Friday", StringComparison.OrdinalIgnoreCase)) return 4;
            if (dt.StartsWith("Saturday", StringComparison.OrdinalIgnoreCase)) return 5;
            if (dt.StartsWith("Sunday", StringComparison.OrdinalIgnoreCase)) return 6;
            if (dt.StartsWith("Holiday", StringComparison.OrdinalIgnoreCase)) return 7;
            return -1;
        }

        // ═══════════════════════════════════════════════════════════════
        //  SolveInstance — main logic
        // ═══════════════════════════════════════════════════════════════
        protected override void SolveInstance(IGH_DataAccess DA)
        {
            // --- Read inputs ---
            string sqlPath = "";
            int peakStart = 13, peakEnd = 17;
            var peakDays = new List<int>();
            double copC = 3.0, copH = 1.0;
            var shedInput = new List<double>();
            var proposedInput = new List<double>();
            double battKwh = 0, chgKw = 50, disKw = 50, eff = 0.9, socMin = 0.1, socMax = 0.9;
            bool case1 = false, case2 = false, resilience = false, run = false;

            DA.GetData(0, ref sqlPath);
            DA.GetData(1, ref peakStart);
            DA.GetData(2, ref peakEnd);
            DA.GetDataList(3, peakDays);
            DA.GetData(4, ref copC);
            DA.GetData(5, ref copH);
            DA.GetDataList(6, shedInput);
            DA.GetDataList(7, proposedInput);
            DA.GetData(8, ref battKwh);
            DA.GetData(9, ref chgKw);
            DA.GetData(10, ref disKw);
            DA.GetData(11, ref eff);
            DA.GetData(12, ref socMin);
            DA.GetData(13, ref socMax);
            DA.GetData(14, ref case1);
            DA.GetData(15, ref case2);
            DA.GetData(16, ref resilience);
            DA.GetData(17, ref run);

            if (!run)
            {
                DA.SetData(0, 0);
                DA.SetData(1, "Set Run = True to execute.");
                Message = "Idle";
                return;
            }

            // Validate
            if (string.IsNullOrWhiteSpace(sqlPath) || !File.Exists(sqlPath))
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "SQL file not found: " + sqlPath);
                return;
            }

            if (peakDays.Count == 0)
                peakDays = new List<int> { 0, 1, 2, 3, 4 }; // Mon-Fri

            copC = Math.Max(0.1, copC);
            copH = Math.Max(0.1, copH);
            eff = Math.Max(0.01, Math.Min(1.0, eff));
            socMin = Math.Max(0.0, Math.Min(0.99, socMin));
            socMax = Math.Max(socMin + 0.01, Math.Min(1.0, socMax));

            // ═══════════════════════════════════════════════════════════
            //  STEP 1: Read SQL data
            //  EnergyPlus variables come in two unit types:
            //    "Rate" variables → W (watts)   → ÷1000 = kW
            //    "Energy" variables / meters → J → ÷3,600,000 = kW (hourly avg)
            // ═══════════════════════════════════════════════════════════
            var log = new StringBuilder();
            double[] facilityElec, produced, coolingIdeal, heatingIdeal;
            HourRecord[] timeInfo;
            bool elecIsRate = false;   // track unit type for facility electricity
            bool prodIsRate = false;   // track unit type for produced electricity

            try
            {
                string connStr = "Data Source=" + sqlPath + ";Version=3;Read Only=True;";
                timeInfo = ReadTimeInfo(connStr);
                int n = timeInfo.Length;

                if (n == 0)
                {
                    AddRuntimeMessage(GH_RuntimeMessageLevel.Error,
                        "No hourly simulation data found in SQL. Check that the file is a valid EnergyPlus output.");
                    return;
                }

                // --- Facility electricity demand ---
                // Try Rate variable first (W), then meter (J)
                facilityElec = ReadHourlySeries(connStr,
                    "Facility Total Electricity Demand Rate", false, false, n);
                if (facilityElec != null)
                {
                    elecIsRate = true;
                }
                else
                {
                    facilityElec = ReadHourlySeries(connStr,
                        "Electricity:Facility", true, false, n);
                }
                if (facilityElec == null)
                {
                    AddRuntimeMessage(GH_RuntimeMessageLevel.Error,
                        "No electricity demand data found. Looked for:\n" +
                        "  - Facility Total Electricity Demand Rate (variable, W)\n" +
                        "  - Electricity:Facility (meter, J)\n" +
                        "Ensure Honeybee simulation includes electricity output.");
                    return;
                }

                // --- Produced electricity (PV) ---
                // Try Rate variable first (W), then meter (J)
                produced = ReadHourlySeries(connStr,
                    "Facility Total Produced Electricity Rate", false, false, n);
                if (produced != null)
                {
                    prodIsRate = true;
                }
                else
                {
                    produced = ReadHourlySeries(connStr,
                        "ElectricityProduced:Facility", true, false, n);
                    if (produced == null)
                        produced = ReadHourlySeries(connStr,
                            "%Produced%Electricity%", false, true, n);
                }
                if (produced == null)
                {
                    produced = new double[n];
                    AddRuntimeMessage(GH_RuntimeMessageLevel.Warning,
                        "No on-site electricity production data found. Assuming no PV.");
                }

                // --- Cooling / Heating ideal loads (always in J) ---
                coolingIdeal = ReadHourlySeries(connStr,
                    "Zone Ideal Loads Supply Air Total Cooling Energy", false, true, n);
                if (coolingIdeal == null)
                    coolingIdeal = new double[n];

                heatingIdeal = ReadHourlySeries(connStr,
                    "Zone Ideal Loads Supply Air Total Heating Energy", false, true, n);
                if (heatingIdeal == null)
                    heatingIdeal = new double[n];
            }
            catch (Exception ex)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "SQL read error: " + ex.Message);
                return;
            }

            int hours = timeInfo.Length;

            // ═══════════════════════════════════════════════════════════
            //  STEP 2: Build baseline load (kW)
            //  baseline = Facility_Elec + Cooling_Ideal/COP_C + Heating_Ideal/COP_H
            //  Rate (W) → ÷1000 = kW;  Energy (J) → ÷3,600,000 = kW
            // ═══════════════════════════════════════════════════════════
            const double W_TO_KW = 1.0 / 1000.0;
            const double J_TO_KW = 1.0 / 3600000.0;
            double elecFactor = elecIsRate ? W_TO_KW : J_TO_KW;
            double prodFactor = prodIsRate ? W_TO_KW : J_TO_KW;

            var baselineLoad = new double[hours];
            var producedKw = new double[hours];
            for (int i = 0; i < hours; i++)
            {
                double elecKw = facilityElec[i] * elecFactor;
                double coolKw = coolingIdeal[i] * J_TO_KW / copC;   // Ideal Loads always in J
                double heatKw = heatingIdeal[i] * J_TO_KW / copH;   // Ideal Loads always in J
                baselineLoad[i] = elecKw + coolKw + heatKw;
                producedKw[i] = produced[i] * prodFactor;
            }

            // ═══════════════════════════════════════════════════════════
            //  STEP 3: Net load (baseline - PV production)
            // ═══════════════════════════════════════════════════════════
            var netLoad = new double[hours];
            for (int i = 0; i < hours; i++)
                netLoad[i] = Math.Max(0, baselineLoad[i] - producedKw[i]);

            // ═══════════════════════════════════════════════════════════
            //  STEP 4: Determine on-peak mask
            // ═══════════════════════════════════════════════════════════
            var isOnPeak = new bool[hours];
            for (int i = 0; i < hours; i++)
            {
                int hod = timeInfo[i].Hour;  // 0-23
                int dow = DayTypeToIndex(timeInfo[i].DayType);
                bool hourMatch = (peakStart <= peakEnd)
                    ? (hod >= peakStart && hod <= peakEnd)
                    : (hod >= peakStart || hod <= peakEnd); // wrap-around
                bool dayMatch = peakDays.Contains(dow);
                isOnPeak[i] = hourMatch && dayMatch;
            }

            // On-peak baseline peak
            double peakBaseline = 0;
            int peakBaselineHour = 0;
            for (int i = 0; i < hours; i++)
            {
                if (isOnPeak[i] && netLoad[i] > peakBaseline)
                {
                    peakBaseline = netLoad[i];
                    peakBaselineHour = i;
                }
            }

            if (peakBaseline <= 0)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error,
                    "On-peak baseline demand is zero. Check peak hour/day settings.");
                return;
            }

            double threshold10pct = peakBaseline * 0.10;

            // ═══════════════════════════════════════════════════════════
            //  STEP 5: Expand shed schedule
            // ═══════════════════════════════════════════════════════════
            var shedKw = new double[hours];
            bool hasShed = shedInput.Count > 0;
            if (hasShed)
            {
                if (shedInput.Count >= hours)
                {
                    for (int i = 0; i < hours; i++)
                        shedKw[i] = Math.Max(0, shedInput[i]);
                }
                else if (shedInput.Count == 24)
                {
                    for (int i = 0; i < hours; i++)
                        shedKw[i] = isOnPeak[i] ? Math.Max(0, shedInput[timeInfo[i].Hour]) : 0;
                }
                else if (shedInput.Count == 1)
                {
                    double val = Math.Max(0, shedInput[0]);
                    for (int i = 0; i < hours; i++)
                        shedKw[i] = isOnPeak[i] ? val : 0;
                }
                else
                {
                    AddRuntimeMessage(GH_RuntimeMessageLevel.Warning,
                        "Shed schedule must be 1, 24, or 8760 values. Ignoring.");
                    hasShed = false;
                }
            }

            // ═══════════════════════════════════════════════════════════
            //  STEP 6A: Peak Load Optimization scoring (shed only, no battery)
            // ═══════════════════════════════════════════════════════════
            double peakAfterShed = 0;
            for (int i = 0; i < hours; i++)
            {
                double val = netLoad[i] - shedKw[i];
                if (isOnPeak[i] && val > peakAfterShed)
                    peakAfterShed = val;
            }
            double shedReductionPct = (peakBaseline - peakAfterShed) / peakBaseline * 100.0;
            bool peakOptPass = hasShed && shedReductionPct >= 10.0;
            int peakOptPoint = peakOptPass ? 1 : 0;

            // ═══════════════════════════════════════════════════════════
            //  STEP 6B: On-site Storage scoring (battery simulation)
            // ═══════════════════════════════════════════════════════════
            var battDischarge = new double[hours];
            var battCharge = new double[hours];
            var batterySoc = new double[hours];
            double soc = (socMin + socMax) / 2.0; // start at midpoint
            bool hasBattery = battKwh > 0;

            if (hasBattery)
            {
                double capKwh = battKwh;
                for (int i = 0; i < hours; i++)
                {
                    if (isOnPeak[i])
                    {
                        // Discharge during on-peak
                        double available = (soc - socMin) * capKwh;
                        double discharge = Math.Min(disKw, Math.Min(available, netLoad[i]));
                        discharge = Math.Max(0, discharge);
                        soc -= discharge / capKwh;
                        battDischarge[i] = discharge;
                    }
                    else
                    {
                        // Charge during off-peak
                        double room = (socMax - soc) * capKwh;
                        double chargeEnergy = Math.Min(chgKw, room / eff);
                        chargeEnergy = Math.Max(0, chargeEnergy);
                        soc += chargeEnergy * eff / capKwh;
                        battCharge[i] = chargeEnergy;
                    }
                    batterySoc[i] = soc;
                }
            }

            // Peak after battery (only battery, no shed)
            double peakAfterBattery = 0;
            for (int i = 0; i < hours; i++)
            {
                double val = netLoad[i] - battDischarge[i] + battCharge[i];
                if (isOnPeak[i] && val > peakAfterBattery)
                    peakAfterBattery = val;
            }
            double battReductionPct = hasBattery
                ? (peakBaseline - peakAfterBattery) / peakBaseline * 100.0 : 0;
            bool storagePass = hasBattery && battReductionPct >= 10.0;
            int storagePoint = storagePass ? 1 : 0;

            // ═══════════════════════════════════════════════════════════
            //  STEP 6C: Flexible Operating scoring (proposed vs baseline)
            // ═══════════════════════════════════════════════════════════
            bool hasProposed = proposedInput.Count >= hours;
            bool flexiblePass = false;
            double flexShiftedKw = 0;
            double flexTimeOffset = 0;

            if (hasProposed)
            {
                // Find on-peak hours where proposed < baseline (load was reduced)
                var reductionHours = new List<int>();
                var increaseHours = new List<int>();
                double totalReduced = 0;
                double totalIncreased = 0;

                for (int i = 0; i < hours; i++)
                {
                    double diff = netLoad[i] - proposedInput[i];
                    if (isOnPeak[i] && diff > 0)
                    {
                        reductionHours.Add(i);
                        totalReduced += diff;
                    }
                    else if (!isOnPeak[i] && diff < 0)
                    {
                        increaseHours.Add(i);
                        totalIncreased += Math.Abs(diff);
                    }
                }

                // Check >=10% of peak was shifted
                flexShiftedKw = totalReduced > 0
                    ? reductionHours.Max(h => netLoad[h] - proposedInput[h]) : 0;
                bool magnitudePass = flexShiftedKw >= threshold10pct;

                // Check time offset >= 2 hours between reduction center and increase center
                if (magnitudePass && reductionHours.Count > 0 && increaseHours.Count > 0)
                {
                    // Use weighted center of mass (within each day to avoid cross-day issues)
                    double redCenter = WeightedCenter(reductionHours,
                        reductionHours.Select(h => netLoad[h] - proposedInput[h]).ToList());
                    double incCenter = WeightedCenter(increaseHours,
                        increaseHours.Select(h => proposedInput[h] - netLoad[h]).ToList());
                    flexTimeOffset = Math.Abs(incCenter - redCenter);
                    flexiblePass = flexTimeOffset >= 2.0;
                }
            }
            else if (proposedInput.Count > 0 && proposedInput.Count < hours)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Warning,
                    "Proposed schedule must have 8760 values for Flexible Operating evaluation. Ignoring.");
            }

            int flexiblePoint = flexiblePass ? 1 : 0;

            // ═══════════════════════════════════════════════════════════
            //  STEP 6D: Grid Resilience (user input)
            // ═══════════════════════════════════════════════════════════
            int resiliencePoint = resilience ? 1 : 0;

            // ═══════════════════════════════════════════════════════════
            //  STEP 7: Scoring
            // ═══════════════════════════════════════════════════════════
            int case1Points = case1 ? 2 : 0;
            int case2Points = (!case1 && case2) ? 1 : 0;
            int case3Points = Math.Min(2, peakOptPoint + storagePoint + flexiblePoint + resiliencePoint);
            int basePoints = Math.Max(case1Points, case2Points);
            int totalPoints = Math.Min(2, basePoints + case3Points);

            // ═══════════════════════════════════════════════════════════
            //  STEP 8: Build modified load (all strategies combined)
            // ═══════════════════════════════════════════════════════════
            var modifiedLoad = new double[hours];
            for (int i = 0; i < hours; i++)
            {
                modifiedLoad[i] = netLoad[i] - shedKw[i] - battDischarge[i] + battCharge[i];
                modifiedLoad[i] = Math.Max(0, modifiedLoad[i]);
            }

            // Overall peak reduction
            double peakModified = 0;
            for (int i = 0; i < hours; i++)
                if (isOnPeak[i] && modifiedLoad[i] > peakModified)
                    peakModified = modifiedLoad[i];

            double overallReductionPct = (peakBaseline - peakModified) / peakBaseline * 100.0;

            // Secondary peak warning
            double baselineMax = netLoad.Max();
            double modifiedMax = modifiedLoad.Max();
            bool secondaryPeakWarning = modifiedMax > baselineMax;

            // ═══════════════════════════════════════════════════════════
            //  STEP 9: Chart data
            // ═══════════════════════════════════════════════════════════
            var annualChart = new GH_Structure<GH_Number>();
            var pathBase = new GH_Path(0);
            var pathMod = new GH_Path(1);
            for (int i = 0; i < hours; i++)
            {
                annualChart.Append(new GH_Number(netLoad[i]), pathBase);
                annualChart.Append(new GH_Number(modifiedLoad[i]), pathMod);
            }

            // Typical day profiles
            var typicalDay = BuildTypicalDayChart(timeInfo, netLoad, modifiedLoad, hours);

            // ═══════════════════════════════════════════════════════════
            //  STEP 10: Points breakdown
            // ═══════════════════════════════════════════════════════════
            string peakTimeStr = string.Format("{0}:00-{1}:00", peakStart, peakEnd);
            string peakDayStr = string.Join(",",
                peakDays.Select(d => new[] { "Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun" }
                    [Math.Min(d, 6)]));

            var breakdown = new StringBuilder();
            breakdown.AppendLine("=== LEED v4.1 Grid Harmonization Score ===");
            breakdown.AppendLine(string.Format("Total Points: {0} / 2", totalPoints));
            breakdown.AppendLine();
            breakdown.AppendLine(string.Format("Case 1 (DR Program Enrolled):    {0} → {1} pt(s)",
                case1 ? "YES" : "NO", case1Points));
            breakdown.AppendLine(string.Format("Case 2 (DR Capable):             {0} → {1} pt(s)",
                case2 ? "YES" : "NO", case2Points));
            breakdown.AppendLine(string.Format("  (Case 1 & 2 are mutually exclusive; max of the two = {0} pt)",
                basePoints));
            breakdown.AppendLine();
            breakdown.AppendLine(string.Format("Case 3 Strategies (cap 2):       {0} pt(s)", case3Points));
            breakdown.AppendLine(string.Format("  Peak Load Optimization:        {0} ({1:F1}% reduction) → {2} pt",
                peakOptPass ? "PASS" : "FAIL", hasShed ? shedReductionPct : 0, peakOptPoint));
            breakdown.AppendLine(string.Format("  On-site Storage:               {0} ({1:F1}% reduction) → {2} pt",
                storagePass ? "PASS" : "FAIL", battReductionPct, storagePoint));
            breakdown.AppendLine(string.Format("  Flexible Operating:            {0} ({1:F1} kW shifted, {2:F1} hr offset) → {3} pt",
                flexiblePass ? "PASS" : "FAIL", flexShiftedKw, flexTimeOffset, flexiblePoint));
            breakdown.AppendLine(string.Format("  Grid Resilience:               {0} → {1} pt",
                resilience ? "YES" : "NO", resiliencePoint));
            breakdown.AppendLine();
            breakdown.AppendLine(string.Format("Final = min(2, max(Case1,Case2) + Case3) = min(2, {0} + {1}) = {2}",
                basePoints, case3Points, totalPoints));

            // ═══════════════════════════════════════════════════════════
            //  STEP 11: Calculation Log (中英文)
            // ═══════════════════════════════════════════════════════════
            string calcLog = BuildCalculationLog(
                hours, peakTimeStr, peakDayStr, copC, copH,
                peakBaseline, peakBaselineHour, timeInfo,
                hasShed, shedReductionPct, peakAfterShed, shedKw,
                hasBattery, battKwh, chgKw, disKw, eff, socMin, socMax,
                battReductionPct, peakAfterBattery, battDischarge, battCharge,
                hasProposed, flexShiftedKw, flexTimeOffset, flexiblePass,
                resilience,
                peakOptPoint, storagePoint, flexiblePoint, resiliencePoint,
                case3Points, totalPoints, basePoints,
                threshold10pct);

            // ═══════════════════════════════════════════════════════════
            //  STEP 12: Methodology (中英文)
            // ═══════════════════════════════════════════════════════════
            string methodology = BuildMethodology(
                peakTimeStr, peakDayStr, copC, copH, hasBattery, hasShed, hasProposed);

            // ═══════════════════════════════════════════════════════════
            //  STEP 13: LEED Narrative
            // ═══════════════════════════════════════════════════════════
            string narrative = BuildNarrative(
                hours, peakBaseline, peakModified, overallReductionPct,
                peakTimeStr, peakDayStr,
                hasShed, shedReductionPct, hasBattery, battKwh, battReductionPct,
                hasProposed, flexShiftedKw, flexTimeOffset, flexiblePass,
                resilience, totalPoints, secondaryPeakWarning);

            // ═══════════════════════════════════════════════════════════
            //  STEP 14: Recommended BESS capacity
            // ═══════════════════════════════════════════════════════════
            string recommendation = BuildRecommendation(
                netLoad, isOnPeak, hours,
                chgKw, disKw, eff, socMin, socMax,
                peakBaseline, threshold10pct,
                hasShed, peakOptPass, resilience);

            // ═══════════════════════════════════════════════════════════
            //  Set outputs
            // ═══════════════════════════════════════════════════════════
            Message = totalPoints + " pt" + (totalPoints != 1 ? "s" : "");

            DA.SetData(0, totalPoints);
            DA.SetData(1, breakdown.ToString());
            DA.SetData(2, calcLog);
            DA.SetData(3, methodology);
            DA.SetDataList(4, baselineLoad.ToList());
            DA.SetDataList(5, netLoad.ToList());
            DA.SetDataList(6, modifiedLoad.ToList());
            DA.SetDataList(7, batterySoc.ToList());
            DA.SetData(8, overallReductionPct);
            DA.SetDataTree(9, annualChart);
            DA.SetDataTree(10, typicalDay);
            DA.SetData(11, narrative);
            DA.SetData(12, secondaryPeakWarning);
            DA.SetData(13, recommendation);
        }

        // ═══════════════════════════════════════════════════════════════
        //  SQL helpers
        // ═══════════════════════════════════════════════════════════════

        private HourRecord[] ReadTimeInfo(string connStr)
        {
            var records = new List<HourRecord>();
            using (var conn = new SQLiteConnection(connStr))
            {
                conn.Open();
                string sql = @"
                    SELECT DISTINCT t.TimeIndex, t.Month, t.Day, t.Hour, t.DayType
                    FROM Time t
                    WHERE (t.WarmupFlag = 0 OR t.WarmupFlag IS NULL)
                      AND t.DayType NOT LIKE '%DesignDay%'
                      AND t.Interval = 60
                    ORDER BY t.TimeIndex";

                using (var cmd = new SQLiteCommand(sql, conn))
                using (var reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        records.Add(new HourRecord
                        {
                            Month = reader.GetInt32(1),
                            Day = reader.GetInt32(2),
                            Hour = reader.GetInt32(3),
                            DayType = reader.GetString(4)
                        });
                    }
                }
            }
            return records.ToArray();
        }

        /// <summary>
        /// Read hourly series from EnergyPlus SQL.
        /// If isMeter=true, queries IsMeter=1 with exact name match.
        /// If isMeter=false and sumAcrossKeys=true, uses LIKE match and sums across zones.
        /// Returns values in original units (typically Joules for energy).
        /// Returns null if no data found.
        /// </summary>
        private double[] ReadHourlySeries(string connStr, string varName,
            bool isMeter, bool sumAcrossKeys, int expectedCount)
        {
            var values = new List<double>();
            using (var conn = new SQLiteConnection(connStr))
            {
                conn.Open();
                string whereClause;
                if (isMeter)
                {
                    whereClause = "rdd.Name = @name AND rdd.IsMeter = 1";
                }
                else if (varName.Contains("%"))
                {
                    whereClause = "rdd.Name LIKE @name AND rdd.IsMeter = 0";
                }
                else
                {
                    whereClause = "rdd.Name = @name AND rdd.IsMeter = 0";
                }

                string aggr = sumAcrossKeys ? "SUM(rd.Value)" : "rd.Value";
                string groupBy = sumAcrossKeys ? "GROUP BY t.TimeIndex" : "";

                string sql = string.Format(@"
                    SELECT {0}
                    FROM ReportData rd
                    JOIN ReportDataDictionary rdd
                        ON rd.ReportDataDictionaryIndex = rdd.ReportDataDictionaryIndex
                    JOIN Time t ON rd.TimeIndex = t.TimeIndex
                    WHERE {1}
                      AND rdd.ReportingFrequency = 'Hourly'
                      AND (t.WarmupFlag = 0 OR t.WarmupFlag IS NULL)
                      AND t.DayType NOT LIKE '%DesignDay%'
                    {2}
                    ORDER BY t.TimeIndex",
                    aggr, whereClause, groupBy);

                using (var cmd = new SQLiteCommand(sql, conn))
                {
                    cmd.Parameters.AddWithValue("@name", varName);
                    using (var reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                            values.Add(reader.GetDouble(0));
                    }
                }
            }

            if (values.Count == 0) return null;

            // Pad or trim to expected count
            if (values.Count < expectedCount)
            {
                while (values.Count < expectedCount)
                    values.Add(0);
            }
            else if (values.Count > expectedCount)
            {
                values = values.Take(expectedCount).ToList();
            }

            return values.ToArray();
        }

        // ═══════════════════════════════════════════════════════════════
        //  Typical day chart builder
        // ═══════════════════════════════════════════════════════════════
        private GH_Structure<GH_Number> BuildTypicalDayChart(
            HourRecord[] timeInfo, double[] netLoad, double[] modifiedLoad, int hours)
        {
            // {0}=workday baseline, {1}=weekend baseline, {2}=workday modified, {3}=weekend modified
            var wdBase = new double[24];
            var weBase = new double[24];
            var wdMod = new double[24];
            var weMod = new double[24];
            var wdCount = new int[24];
            var weCount = new int[24];

            for (int i = 0; i < hours; i++)
            {
                int hod = timeInfo[i].Hour;
                if (hod < 0 || hod > 23) continue;
                int dow = DayTypeToIndex(timeInfo[i].DayType);
                bool isWeekday = dow >= 0 && dow <= 4;

                if (isWeekday)
                {
                    wdBase[hod] += netLoad[i];
                    wdMod[hod] += modifiedLoad[i];
                    wdCount[hod]++;
                }
                else
                {
                    weBase[hod] += netLoad[i];
                    weMod[hod] += modifiedLoad[i];
                    weCount[hod]++;
                }
            }

            // Average
            for (int h = 0; h < 24; h++)
            {
                if (wdCount[h] > 0) { wdBase[h] /= wdCount[h]; wdMod[h] /= wdCount[h]; }
                if (weCount[h] > 0) { weBase[h] /= weCount[h]; weMod[h] /= weCount[h]; }
            }

            var tree = new GH_Structure<GH_Number>();
            for (int h = 0; h < 24; h++)
            {
                tree.Append(new GH_Number(wdBase[h]), new GH_Path(0));
                tree.Append(new GH_Number(weBase[h]), new GH_Path(1));
                tree.Append(new GH_Number(wdMod[h]), new GH_Path(2));
                tree.Append(new GH_Number(weMod[h]), new GH_Path(3));
            }
            return tree;
        }

        // ═══════════════════════════════════════════════════════════════
        //  Weighted center of mass (hour index)
        // ═══════════════════════════════════════════════════════════════
        private double WeightedCenter(List<int> indices, List<double> weights)
        {
            double totalW = weights.Sum();
            if (totalW <= 0) return 0;

            // Use modular arithmetic to handle day boundaries
            // Convert to hour-of-day for center calculation
            double sumWH = 0;
            for (int i = 0; i < indices.Count; i++)
            {
                double hod = indices[i] % 24;
                sumWH += hod * weights[i];
            }
            return sumWH / totalW;
        }

        // ═══════════════════════════════════════════════════════════════
        //  Calculation Log (中英文)
        // ═══════════════════════════════════════════════════════════════
        private string BuildCalculationLog(
            int hours, string peakTimeStr, string peakDayStr, double copC, double copH,
            double peakBaseline, int peakBaselineHour, HourRecord[] timeInfo,
            bool hasShed, double shedReductionPct, double peakAfterShed, double[] shedKw,
            bool hasBattery, double battKwh, double chgKw, double disKw, double eff,
            double socMin, double socMax,
            double battReductionPct, double peakAfterBattery, double[] battDischarge, double[] battCharge,
            bool hasProposed, double flexShiftedKw, double flexTimeOffset, bool flexiblePass,
            bool resilience,
            int peakOptPoint, int storagePoint, int flexiblePoint, int resiliencePoint,
            int case3Points, int totalPoints, int basePoints,
            double threshold10pct)
        {
            var sb = new StringBuilder();
            var peakHr = timeInfo[peakBaselineHour];
            string peakTimeLabel = string.Format("{0}/{1} {2}:00", peakHr.Month, peakHr.Day, peakHr.Hour);

            double totalShedKwh = shedKw.Sum();
            double totalBattDischargeKwh = battDischarge.Sum();
            double totalBattChargeKwh = battCharge.Sum();

            // ─── 中文 ───
            sb.AppendLine("=== 計算過程 ===");
            sb.AppendLine();
            sb.AppendLine("【基準線建立】");
            sb.AppendLine(string.Format("  資料來源：EnergyPlus SQL ({0} 小時)", hours));
            sb.AppendLine(string.Format("  基準負載 = Facility Total Electricity Demand Rate + Cooling_Ideal/{0:F1} + Heating_Ideal/{1:F1}", copC, copH));
            sb.AppendLine(string.Format("  單位轉換：Rate(W) → ÷1,000 = kW；Energy(J) → ÷3,600,000 = kW"));
            sb.AppendLine(string.Format("  淨負載 = max(0, 基準負載 - 現場產電)"));
            sb.AppendLine(string.Format("  尖峰時段：{0}，{1}", peakTimeStr, peakDayStr));
            sb.AppendLine(string.Format("  基準 On-peak 尖峰需求：{0:F3} kW（發生於 {1}）", peakBaseline, peakTimeLabel));
            sb.AppendLine(string.Format("  10% 門檻值：{0:F3} kW", threshold10pct));
            sb.AppendLine();

            sb.AppendLine("【Peak Load Optimization（非電池卸載策略）】");
            if (hasShed)
            {
                sb.AppendLine(string.Format("  卸載排程總量：{0:F1} kWh（全年）", totalShedKwh));
                sb.AppendLine(string.Format("  卸載後 On-peak 尖峰：{0:F3} kW", peakAfterShed));
                sb.AppendLine(string.Format("  削減率：({0:F3} - {1:F3}) / {0:F3} × 100% = {2:F1}%",
                    peakBaseline, peakAfterShed, shedReductionPct));
                sb.AppendLine(string.Format("  判定：{0}（門檻 ≥ 10%）→ {1} 分",
                    shedReductionPct >= 10 ? "PASS" : "FAIL", peakOptPoint));
            }
            else
            {
                sb.AppendLine("  未提供卸載排程 → 0 分");
            }
            sb.AppendLine();

            sb.AppendLine("【On-site Storage（電池儲能）】");
            if (hasBattery)
            {
                sb.AppendLine(string.Format("  電池容量：{0:F1} kWh，充電功率：{1:F1} kW，放電功率：{2:F1} kW",
                    battKwh, chgKw, disKw));
                sb.AppendLine(string.Format("  往返效率：{0:P0}，SOC 範圍：{1:F0}%-{2:F0}%",
                    eff, socMin * 100, socMax * 100));
                sb.AppendLine(string.Format("  全年放電總量：{0:F1} kWh，充電總量：{1:F1} kWh",
                    totalBattDischargeKwh, totalBattChargeKwh));
                sb.AppendLine(string.Format("  電池介入後 On-peak 尖峰：{0:F3} kW", peakAfterBattery));
                sb.AppendLine(string.Format("  削減率：({0:F3} - {1:F3}) / {0:F3} × 100% = {2:F1}%",
                    peakBaseline, peakAfterBattery, battReductionPct));
                sb.AppendLine(string.Format("  判定：{0}（門檻 ≥ 10%）→ {1} 分",
                    battReductionPct >= 10 ? "PASS" : "FAIL", storagePoint));
            }
            else
            {
                sb.AppendLine("  未配置電池 → 0 分");
            }
            sb.AppendLine();

            sb.AppendLine("【Flexible Operating Scenarios（彈性營運情境）】");
            if (hasProposed)
            {
                sb.AppendLine(string.Format("  已提供 Proposed Schedule（{0} 筆）", hours));
                sb.AppendLine(string.Format("  尖峰時段最大轉移量：{0:F3} kW（門檻 {1:F3} kW）",
                    flexShiftedKw, threshold10pct));
                sb.AppendLine(string.Format("  負載移動時間偏移：{0:F1} 小時（門檻 ≥ 2 小時）", flexTimeOffset));
                sb.AppendLine(string.Format("  判定：{0} → {1} 分",
                    flexiblePass ? "PASS" : "FAIL", flexiblePoint));
            }
            else
            {
                sb.AppendLine("  未提供 Proposed Schedule → 0 分");
            }
            sb.AppendLine();

            sb.AppendLine("【Grid Resilience Technologies（電網韌性）】");
            sb.AppendLine(string.Format("  使用者輸入：{0} → {1} 分", resilience ? "YES" : "NO", resiliencePoint));
            sb.AppendLine();

            sb.AppendLine("【總分計算】");
            sb.AppendLine(string.Format("  Case 3 = min(2, {0}+{1}+{2}+{3}) = {4}",
                peakOptPoint, storagePoint, flexiblePoint, resiliencePoint, case3Points));
            sb.AppendLine(string.Format("  Case 1/2 基礎分 = {0}", basePoints));
            sb.AppendLine(string.Format("  總分 = min(2, {0} + {1}) = {2}", basePoints, case3Points, totalPoints));
            sb.AppendLine();

            // ─── English ───
            sb.AppendLine("=== Calculation Log ===");
            sb.AppendLine();
            sb.AppendLine("[Baseline Establishment]");
            sb.AppendLine(string.Format("  Data source: EnergyPlus SQL ({0} hours)", hours));
            sb.AppendLine(string.Format("  Baseline load = Facility Total Electricity Demand Rate + Cooling_Ideal/{0:F1} + Heating_Ideal/{1:F1}", copC, copH));
            sb.AppendLine("  Unit conversion: Rate(W) / 1,000 = kW; Energy(J) / 3,600,000 = kW");
            sb.AppendLine("  Net load = max(0, Baseline - On-site Production)");
            sb.AppendLine(string.Format("  On-peak period: {0}, {1}", peakTimeStr, peakDayStr));
            sb.AppendLine(string.Format("  Baseline on-peak demand: {0:F3} kW (at {1})", peakBaseline, peakTimeLabel));
            sb.AppendLine(string.Format("  10% threshold: {0:F3} kW", threshold10pct));
            sb.AppendLine();

            sb.AppendLine("[Peak Load Optimization (non-battery shed)]");
            if (hasShed)
            {
                sb.AppendLine(string.Format("  Total annual shed: {0:F1} kWh", totalShedKwh));
                sb.AppendLine(string.Format("  On-peak peak after shed: {0:F3} kW", peakAfterShed));
                sb.AppendLine(string.Format("  Reduction: ({0:F3} - {1:F3}) / {0:F3} x 100% = {2:F1}%",
                    peakBaseline, peakAfterShed, shedReductionPct));
                sb.AppendLine(string.Format("  Result: {0} (threshold >= 10%) -> {1} pt",
                    shedReductionPct >= 10 ? "PASS" : "FAIL", peakOptPoint));
            }
            else
            {
                sb.AppendLine("  No shed schedule provided -> 0 pt");
            }
            sb.AppendLine();

            sb.AppendLine("[On-site Storage (battery)]");
            if (hasBattery)
            {
                sb.AppendLine(string.Format("  Battery: {0:F1} kWh, charge {1:F1} kW, discharge {2:F1} kW",
                    battKwh, chgKw, disKw));
                sb.AppendLine(string.Format("  Efficiency: {0:P0}, SOC range: {1:F0}%-{2:F0}%",
                    eff, socMin * 100, socMax * 100));
                sb.AppendLine(string.Format("  Annual discharge: {0:F1} kWh, charge: {1:F1} kWh",
                    totalBattDischargeKwh, totalBattChargeKwh));
                sb.AppendLine(string.Format("  On-peak peak after battery: {0:F3} kW", peakAfterBattery));
                sb.AppendLine(string.Format("  Reduction: ({0:F3} - {1:F3}) / {0:F3} x 100% = {2:F1}%",
                    peakBaseline, peakAfterBattery, battReductionPct));
                sb.AppendLine(string.Format("  Result: {0} (threshold >= 10%) -> {1} pt",
                    battReductionPct >= 10 ? "PASS" : "FAIL", storagePoint));
            }
            else
            {
                sb.AppendLine("  No battery configured -> 0 pt");
            }
            sb.AppendLine();

            sb.AppendLine("[Flexible Operating Scenarios]");
            if (hasProposed)
            {
                sb.AppendLine(string.Format("  Proposed schedule provided ({0} values)", hours));
                sb.AppendLine(string.Format("  Max shifted amount: {0:F3} kW (threshold {1:F3} kW)",
                    flexShiftedKw, threshold10pct));
                sb.AppendLine(string.Format("  Time offset: {0:F1} hours (threshold >= 2 hours)", flexTimeOffset));
                sb.AppendLine(string.Format("  Result: {0} -> {1} pt",
                    flexiblePass ? "PASS" : "FAIL", flexiblePoint));
            }
            else
            {
                sb.AppendLine("  No proposed schedule provided -> 0 pt");
            }
            sb.AppendLine();

            sb.AppendLine("[Grid Resilience Technologies]");
            sb.AppendLine(string.Format("  User input: {0} -> {1} pt", resilience ? "YES" : "NO", resiliencePoint));
            sb.AppendLine();

            sb.AppendLine("[Scoring Summary]");
            sb.AppendLine(string.Format("  Case 3 = min(2, {0}+{1}+{2}+{3}) = {4}",
                peakOptPoint, storagePoint, flexiblePoint, resiliencePoint, case3Points));
            sb.AppendLine(string.Format("  Case 1/2 base = {0}", basePoints));
            sb.AppendLine(string.Format("  Total = min(2, {0} + {1}) = {2}", basePoints, case3Points, totalPoints));

            return sb.ToString();
        }

        // ═══════════════════════════════════════════════════════════════
        //  Methodology (中英文)
        // ═══════════════════════════════════════════════════════════════
        private string BuildMethodology(
            string peakTimeStr, string peakDayStr, double copC, double copH,
            bool hasBattery, bool hasShed, bool hasProposed)
        {
            var sb = new StringBuilder();

            // ─── 中文 ───
            sb.AppendLine("=== 方法論說明 ===");
            sb.AppendLine();
            sb.AppendLine("【分析依據】");
            sb.AppendLine("  本分析依據 LEED v4.1 BD+C「EA Credit: Grid Harmonization」進行。");
            sb.AppendLine("  該學分旨在透過需求響應（Demand Response）技術與負載管理策略，");
            sb.AppendLine("  提升電網可靠性並降低溫室氣體排放。");
            sb.AppendLine();
            sb.AppendLine("【計分架構】");
            sb.AppendLine("  學分上限 2 分。Case 1（DR 計劃參與，2 分）與 Case 2（DR 能力，1 分）互斥；");
            sb.AppendLine("  Case 3（負載彈性與管理策略，1-2 分）可獨立或與 Case 1/2 組合（AND/OR）。");
            sb.AppendLine("  Case 3 包含四項子策略，每項 1 分，上限 2 分：");
            sb.AppendLine("    1. Peak Load Optimization — 非電池策略削減 ≥10% 尖峰負載");
            sb.AppendLine("    2. On-site Storage — 儲能系統削減 ≥10% 尖峰負載");
            sb.AppendLine("    3. Flexible Operating Scenarios — 將 ≥10% 尖峰負載移動 ≥2 小時");
            sb.AppendLine("    4. Grid Resilience Technologies — 電力公司提供韌性計劃");
            sb.AppendLine("  禁止雙重計分（Double-dipping）：同一削減動作不能同時計入多項策略。");
            sb.AppendLine();
            sb.AppendLine("【基準線建立方法】");
            sb.AppendLine("  資料來源：Honeybee EnergyPlus 模擬輸出的 SQL 檔案。");
            sb.AppendLine("  讀取 Facility Total Electricity Demand Rate 作為建築總用電基礎。");
            sb.AppendLine("  若模型使用 Ideal Air Loads，則將 Zone Ideal Loads 冷/熱負載");
            sb.AppendLine(string.Format("  分別除以 COP（冷氣={0:F1}，暖氣={1:F1}）轉換為等效電力需求。", copC, copH));
            sb.AppendLine("  基準負載 = Facility Total Electricity Demand Rate + Cooling/COP + Heating/COP（kW）");
            sb.AppendLine("  淨負載 = max(0, 基準負載 - 現場產電)");
            sb.AppendLine("  根據 LEED 規範，現場產電不計入需求響應貢獻。");
            sb.AppendLine();
            sb.AppendLine("【尖峰時段定義】");
            sb.AppendLine(string.Format("  On-peak 時段：{0}，適用日：{1}", peakTimeStr, peakDayStr));
            sb.AppendLine("  依 EA Prerequisite Minimum Energy Performance 定義之尖峰時段，");
            sb.AppendLine("  需配合當地氣候區與費率結構調整。");
            sb.AppendLine();

            if (hasShed)
            {
                sb.AppendLine("【Peak Load Optimization 方法】");
                sb.AppendLine("  使用者提供卸載排程（Shed Schedule），指定各小時的卸載量（kW）。");
                sb.AppendLine("  代表在 DR 事件期間關閉非必要照明、調高空調溫度等行動計畫。");
                sb.AppendLine("  判定：卸載後 On-peak 尖峰 ≤ 基準尖峰 × 90% 則 PASS。");
                sb.AppendLine();
            }

            if (hasBattery)
            {
                sb.AppendLine("【On-site Storage 方法】");
                sb.AppendLine("  模擬電池儲能系統（BESS）的逐時充放電行為。");
                sb.AppendLine("  Off-peak 時段充電（受充電功率與 SOC 上限約束）；");
                sb.AppendLine("  On-peak 時段放電（受放電功率與 SOC 下限約束）。");
                sb.AppendLine("  判定：電池放電後 On-peak 尖峰 ≤ 基準尖峰 × 90% 則 PASS。");
                sb.AppendLine();
            }

            if (hasProposed)
            {
                sb.AppendLine("【Flexible Operating Scenarios 方法】");
                sb.AppendLine("  比對基準排程與使用者提供的 Proposed Schedule（8760 小時）。");
                sb.AppendLine("  檢測尖峰時段的負載減少量是否 ≥ 基準尖峰的 10%，");
                sb.AppendLine("  且被轉移的負載出現在距原時段 ≥2 小時的離峰時段。");
                sb.AppendLine("  此策略獨立於電池，反映營運行為的改變（如預冷、產線排程調整）。");
                sb.AppendLine();
            }

            sb.AppendLine("【邊界檢查】");
            sb.AppendLine("  Secondary Peak Warning：檢查所有策略實施後的負載曲線，");
            sb.AppendLine("  確認 max(Modified_Load) 不超過 max(Baseline_Load)。");
            sb.AppendLine("  若產生新的更高峰值，將發出警告。");
            sb.AppendLine();

            // ─── English ───
            sb.AppendLine("=== Methodology ===");
            sb.AppendLine();
            sb.AppendLine("[Analysis Basis]");
            sb.AppendLine("  This analysis follows LEED v4.1 BD+C 'EA Credit: Grid Harmonization'.");
            sb.AppendLine("  The credit aims to increase participation in demand response technologies");
            sb.AppendLine("  and programs that improve grid reliability and reduce GHG emissions.");
            sb.AppendLine();
            sb.AppendLine("[Scoring Framework]");
            sb.AppendLine("  Credit cap: 2 points. Case 1 (DR program, 2pts) and Case 2 (DR capable, 1pt)");
            sb.AppendLine("  are mutually exclusive. Case 3 (Load Flexibility, 1-2pts) can stand alone or");
            sb.AppendLine("  combine with Case 1/2 (AND/OR). Case 3 has four sub-strategies, each 1pt, cap 2:");
            sb.AppendLine("    1. Peak Load Optimization - non-battery strategies reduce on-peak >= 10%");
            sb.AppendLine("    2. On-site Storage - storage reduces on-peak >= 10%");
            sb.AppendLine("    3. Flexible Operating Scenarios - shift >= 10% of peak load by >= 2 hours");
            sb.AppendLine("    4. Grid Resilience Technologies - utility resilience program");
            sb.AppendLine("  No double-dipping: the same reduction action cannot count for multiple strategies.");
            sb.AppendLine();
            sb.AppendLine("[Baseline Establishment]");
            sb.AppendLine("  Data source: Honeybee EnergyPlus simulation SQL output.");
            sb.AppendLine("  Facility Total Electricity Demand Rate provides total building electricity.");
            sb.AppendLine("  For Ideal Air Loads models, Zone Ideal Loads cooling/heating energy");
            sb.AppendLine(string.Format("  is converted to electrical demand using COP (cooling={0:F1}, heating={1:F1}).", copC, copH));
            sb.AppendLine("  Baseline = Facility Total Electricity Demand Rate + Cooling/COP + Heating/COP (kW)");
            sb.AppendLine("  Net load = max(0, Baseline - On-site Production)");
            sb.AppendLine("  Per LEED, on-site electricity generation does not count toward DR contribution.");
            sb.AppendLine();
            sb.AppendLine("[On-peak Definition]");
            sb.AppendLine(string.Format("  On-peak hours: {0}, applicable days: {1}", peakTimeStr, peakDayStr));
            sb.AppendLine("  Aligned with EA Prerequisite Minimum Energy Performance on-peak definition,");
            sb.AppendLine("  which varies based on utility climate and pricing structures.");
            sb.AppendLine();

            if (hasShed)
            {
                sb.AppendLine("[Peak Load Optimization Method]");
                sb.AppendLine("  User-provided Shed Schedule specifies hourly load reduction (kW).");
                sb.AppendLine("  Represents action plan for DR events: HVAC setback, lighting reduction, etc.");
                sb.AppendLine("  Pass criteria: on-peak peak after shed <= baseline peak x 90%.");
                sb.AppendLine();
            }

            if (hasBattery)
            {
                sb.AppendLine("[On-site Storage Method]");
                sb.AppendLine("  Simulates Battery Energy Storage System (BESS) hourly charge/discharge.");
                sb.AppendLine("  Charges during off-peak (constrained by charge rate and SOC max);");
                sb.AppendLine("  Discharges during on-peak (constrained by discharge rate and SOC min).");
                sb.AppendLine("  Pass criteria: on-peak peak after BESS <= baseline peak x 90%.");
                sb.AppendLine();
            }

            if (hasProposed)
            {
                sb.AppendLine("[Flexible Operating Scenarios Method]");
                sb.AppendLine("  Compares baseline demand with user-provided Proposed Schedule (8760 hours).");
                sb.AppendLine("  Checks whether on-peak load reduction >= 10% of baseline peak,");
                sb.AppendLine("  and the shifted load appears in off-peak hours >= 2 hours away.");
                sb.AppendLine("  This strategy is independent of battery; it reflects operational changes");
                sb.AppendLine("  such as pre-cooling, production schedule shifts, etc.");
                sb.AppendLine();
            }

            sb.AppendLine("[Boundary Check]");
            sb.AppendLine("  Secondary Peak Warning: verifies that max(Modified_Load) does not exceed");
            sb.AppendLine("  max(Baseline_Load) after all strategies are applied. Warns if a new higher");
            sb.AppendLine("  peak is created by load shifting.");

            return sb.ToString();
        }

        // ═══════════════════════════════════════════════════════════════
        //  LEED Narrative
        // ═══════════════════════════════════════════════════════════════
        private string BuildNarrative(
            int hours, double peakBaseline, double peakModified, double overallReductionPct,
            string peakTimeStr, string peakDayStr,
            bool hasShed, double shedReductionPct,
            bool hasBattery, double battKwh, double battReductionPct,
            bool hasProposed, double flexShiftedKw, double flexTimeOffset, bool flexiblePass,
            bool resilience, int totalPoints, bool secondaryPeakWarning)
        {
            var sb = new StringBuilder();
            sb.AppendLine("LEED v4.1 EA Credit: Grid Harmonization - Narrative Report");
            sb.AppendLine("==========================================================");
            sb.AppendLine();

            sb.AppendLine("1. Building Annual Load Shape Analysis");
            sb.AppendLine("--------------------------------------");
            sb.AppendLine(string.Format(
                "The building's annual load shape was derived from an {0}-hour EnergyPlus simulation " +
                "conducted via Honeybee. The on-peak period is defined as {1} on {2}, " +
                "aligned with the EA Prerequisite Minimum Energy Performance.",
                hours, peakTimeStr, peakDayStr));
            sb.AppendLine(string.Format(
                "The baseline on-peak electrical demand is {0:F1} kW.", peakBaseline));
            sb.AppendLine();

            sb.AppendLine("2. Load Flexibility and Management Strategies Implemented");
            sb.AppendLine("---------------------------------------------------------");

            if (hasShed)
            {
                sb.AppendLine(string.Format(
                    "Peak Load Optimization: A load shedding strategy was developed to reduce " +
                    "non-essential loads during on-peak hours. The strategy achieves a {0:F1}% reduction " +
                    "in on-peak demand{1}.",
                    shedReductionPct,
                    shedReductionPct >= 10 ? ", meeting the 10% threshold" : ", which does not meet the 10% threshold"));
            }

            if (hasBattery)
            {
                sb.AppendLine(string.Format(
                    "On-site Electricity Storage: A {0:F1} kWh battery energy storage system (BESS) " +
                    "is installed to store energy during off-peak hours and discharge during on-peak periods. " +
                    "The BESS achieves a {1:F1}% reduction in on-peak demand{2}.",
                    battKwh, battReductionPct,
                    battReductionPct >= 10
                        ? ", meeting the 10% threshold required for this strategy"
                        : ", which does not meet the 10% threshold"));
            }

            if (hasProposed)
            {
                sb.AppendLine(string.Format(
                    "Flexible Operating Scenarios: Operational schedule modifications shift {0:F1} kW " +
                    "of peak demand with a time offset of {1:F1} hours{2}.",
                    flexShiftedKw, flexTimeOffset,
                    flexiblePass
                        ? ", satisfying both the 10% magnitude and 2-hour duration requirements"
                        : ", which does not fully satisfy the requirements"));
            }

            if (resilience)
            {
                sb.AppendLine(
                    "Grid Resilience Technologies: The project is served by a utility with a " +
                    "resilience program in place that leverages strategies such as islanding " +
                    "and part-load operation.");
            }

            sb.AppendLine();

            sb.AppendLine("3. Results Summary");
            sb.AppendLine("------------------");
            sb.AppendLine(string.Format(
                "After implementing all strategies, the building's on-peak electrical demand " +
                "is reduced from {0:F1} kW to {1:F1} kW, representing an overall reduction of {2:F1}%.",
                peakBaseline, peakModified, overallReductionPct));

            if (secondaryPeakWarning)
            {
                sb.AppendLine(
                    "NOTE: Load management strategies have created a secondary peak that exceeds " +
                    "the original baseline peak. Review the modified load profile for potential issues.");
            }

            sb.AppendLine();
            sb.AppendLine(string.Format(
                "The project achieves {0} point(s) under EA Credit: Grid Harmonization.", totalPoints));
            sb.AppendLine();
            sb.AppendLine("All installed technologies are included in the scope of work for the " +
                "commissioning authority, and the load flexibility and management strategies " +
                "are documented in the building systems manual.");

            return sb.ToString();
        }

        // ═══════════════════════════════════════════════════════════════
        //  BESS capacity recommendation (binary search)
        // ═══════════════════════════════════════════════════════════════

        /// <summary>
        /// Binary search for the minimum battery kWh that achieves ≥10% on-peak reduction.
        /// Returns -1 if even maxCap cannot achieve the target.
        /// </summary>
        private double FindMinBatteryKwh(
            double[] netLoad, bool[] isOnPeak, int hours,
            double chgKw, double disKw, double eff, double socMin, double socMax,
            double peakBaseline, double maxCap)
        {
            // First test if maxCap can achieve it
            if (!TestBatteryCapacity(netLoad, isOnPeak, hours,
                    chgKw, disKw, eff, socMin, socMax, peakBaseline, maxCap))
                return -1;

            double lo = 0, hi = maxCap;
            while (hi - lo > 0.5)
            {
                double mid = (lo + hi) / 2.0;
                if (TestBatteryCapacity(netLoad, isOnPeak, hours,
                        chgKw, disKw, eff, socMin, socMax, peakBaseline, mid))
                    hi = mid;
                else
                    lo = mid;
            }
            return Math.Ceiling(hi * 2) / 2.0; // round up to 0.5 kWh
        }

        private bool TestBatteryCapacity(
            double[] netLoad, bool[] isOnPeak, int hours,
            double chgKw, double disKw, double eff, double socMin, double socMax,
            double peakBaseline, double capKwh)
        {
            double soc = (socMin + socMax) / 2.0;
            double gridPeakAfter = 0;

            for (int i = 0; i < hours; i++)
            {
                double discharge = 0;
                double charge = 0;

                if (isOnPeak[i])
                {
                    double available = (soc - socMin) * capKwh;
                    discharge = Math.Min(disKw, Math.Min(available, netLoad[i]));
                    discharge = Math.Max(0, discharge);
                    soc -= discharge / capKwh;
                }
                else
                {
                    double room = (socMax - soc) * capKwh;
                    charge = Math.Min(chgKw, room / eff);
                    charge = Math.Max(0, charge);
                    soc += charge * eff / capKwh;
                }

                double gl = netLoad[i] - discharge + charge;
                if (isOnPeak[i] && gl > gridPeakAfter)
                    gridPeakAfter = gl;
            }

            double reduction = (peakBaseline - gridPeakAfter) / peakBaseline * 100.0;
            return reduction >= 10.0;
        }

        private string BuildRecommendation(
            double[] netLoad, bool[] isOnPeak, int hours,
            double chgKw, double disKw, double eff, double socMin, double socMax,
            double peakBaseline, double threshold10pct,
            bool hasShed, bool peakOptPass, bool resilience)
        {
            var sb = new StringBuilder();
            double maxCap = 2000.0; // kWh search ceiling

            double minFor1pt = FindMinBatteryKwh(netLoad, isOnPeak, hours,
                chgKw, disKw, eff, socMin, socMax, peakBaseline, maxCap);

            // ─── 中文 ───
            sb.AppendLine("=== 建議電池容量 ===");
            sb.AppendLine();
            sb.AppendLine(string.Format("  基準 On-peak 尖峰：{0:F1} kW", peakBaseline));
            sb.AppendLine(string.Format("  10% 削減目標：{0:F1} kW", peakBaseline * 0.9));
            sb.AppendLine(string.Format("  模擬參數：充電 {0:F0} kW，放電 {1:F0} kW，效率 {2:P0}，SOC {3:F0}%-{4:F0}%",
                chgKw, disKw, eff, socMin * 100, socMax * 100));
            sb.AppendLine();

            if (minFor1pt > 0)
            {
                sb.AppendLine(string.Format("  ▶ On-site Storage (1 分)：最低需要 {0:F1} kWh", minFor1pt));
            }
            else
            {
                sb.AppendLine("  ▶ On-site Storage (1 分)：即使 2000 kWh 也無法達成 10% 削減");
                sb.AppendLine("    （可能原因：放電功率不足，或 off-peak 時間不夠充電）");
            }

            sb.AppendLine();
            sb.AppendLine("  ▶ 拿滿 2 分的組合建議：");

            if (minFor1pt > 0)
            {
                sb.AppendLine(string.Format("    方案 A：電池 {0:F1} kWh + 卸載 ≥{1:F1} kW → Storage(1) + Peak Opt(1) = 2 分",
                    minFor1pt, threshold10pct));
                sb.AppendLine(string.Format("    方案 B：電池 {0:F1} kWh + 電力公司韌性計劃 → Storage(1) + Resilience(1) = 2 分",
                    minFor1pt));

                if (hasShed && peakOptPass)
                    sb.AppendLine(string.Format("    ★ 你目前的卸載策略已 PASS → 只需電池 {0:F1} kWh 即可拿滿 2 分", minFor1pt));
                if (resilience)
                    sb.AppendLine(string.Format("    ★ 韌性計劃已勾選 → 只需電池 {0:F1} kWh 即可拿滿 2 分", minFor1pt));
            }
            else
            {
                sb.AppendLine("    電池路線不可行。建議：卸載(1) + 韌性(1)，或調整充放電功率後重試。");
            }

            sb.AppendLine();

            // ─── English ───
            sb.AppendLine("=== Recommended BESS Capacity ===");
            sb.AppendLine();
            sb.AppendLine(string.Format("  Baseline on-peak demand: {0:F1} kW", peakBaseline));
            sb.AppendLine(string.Format("  10% reduction target: {0:F1} kW", peakBaseline * 0.9));
            sb.AppendLine(string.Format("  Simulation parameters: charge {0:F0} kW, discharge {1:F0} kW, eff {2:P0}, SOC {3:F0}%-{4:F0}%",
                chgKw, disKw, eff, socMin * 100, socMax * 100));
            sb.AppendLine();

            if (minFor1pt > 0)
            {
                sb.AppendLine(string.Format("  > On-site Storage (1 pt): minimum {0:F1} kWh required", minFor1pt));
            }
            else
            {
                sb.AppendLine("  > On-site Storage (1 pt): even 2000 kWh cannot achieve 10% reduction");
                sb.AppendLine("    (Possible cause: discharge rate too low, or insufficient off-peak charging time)");
            }

            sb.AppendLine();
            sb.AppendLine("  > Paths to 2 points:");

            if (minFor1pt > 0)
            {
                sb.AppendLine(string.Format("    Option A: Battery {0:F1} kWh + shed >= {1:F1} kW -> Storage(1) + Peak Opt(1) = 2 pts",
                    minFor1pt, threshold10pct));
                sb.AppendLine(string.Format("    Option B: Battery {0:F1} kWh + grid resilience -> Storage(1) + Resilience(1) = 2 pts",
                    minFor1pt));

                if (hasShed && peakOptPass)
                    sb.AppendLine(string.Format("    * Your shed strategy already PASS -> only need {0:F1} kWh battery for 2 pts", minFor1pt));
                if (resilience)
                    sb.AppendLine(string.Format("    * Resilience checked -> only need {0:F1} kWh battery for 2 pts", minFor1pt));
            }
            else
            {
                sb.AppendLine("    Battery path not feasible. Consider: Shed(1) + Resilience(1), or adjust charge/discharge rates.");
            }

            return sb.ToString();
        }

        // ═══════════════════════════════════════════════════════════════
        //  Component identity
        // ═══════════════════════════════════════════════════════════════
        public override Guid ComponentGuid => new Guid("f7a83b12-6d4e-4c91-b5a8-3e2f1d0c9b78");

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
