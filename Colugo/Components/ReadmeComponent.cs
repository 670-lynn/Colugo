using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using Grasshopper.Kernel;

namespace Colugo.Components
{
    public class ReadmeComponent : GH_Component
    {
        public ReadmeComponent()
            : base("LEED v4.1 Readme", "Readme",
                "Displays an overview of LEED v4.1 BD+C, including credit categories, scorecard, and certification levels.",
                "Colugo", "LEED IP")
        {
        }

        protected override void RegisterInputParams(GH_InputParamManager pManager)
        {
        }

        protected override void RegisterOutputParams(GH_OutputParamManager pManager)
        {
            pManager.AddTextParameter("Overview", "O", "LEED v4.1 BD+C overview and key goals", GH_ParamAccess.item);
            pManager.AddTextParameter("Certification", "C", "Certification levels and point thresholds", GH_ParamAccess.item);
            pManager.AddTextParameter("Scorecard", "S", "Credit categories and available points", GH_ParamAccess.item);
            pManager.AddTextParameter("Full Readme", "R", "Complete LEED v4.1 BD+C readme", GH_ParamAccess.item);
        }

        protected override void SolveInstance(IGH_DataAccess DA)
        {
            DA.SetData(0, GetOverview());
            DA.SetData(1, GetCertificationLevels());
            DA.SetData(2, GetScorecardSummary());
            DA.SetData(3, GetFullReadme());
        }

        private string GetOverview()
        {
            return
                "LEED v4.1 BD+C (Building Design and Construction) is a rating system developed by USGBC\n" +
                "for high-performance green buildings. It evaluates buildings across multiple environmental\n" +
                "and sustainability categories, with 110 total points available.\n\n" +
                "Four key goals guide LEED v4.1:\n" +
                "  1. Ensure Leadership\n" +
                "  2. Increase Achievability\n" +
                "  3. Measure Performance\n" +
                "  4. Expand the Market\n\n" +
                "The rating system applies to: New Construction, Core and Shell, Schools, Retail,\n" +
                "Data Centers, Warehouses and Distribution Centers, Hospitality, and Healthcare.";
        }

        private string GetCertificationLevels()
        {
            return
                "LEED Certification Levels (Total 110 Points):\n" +
                "  - Certified:  40 - 49 points\n" +
                "  - Silver:     50 - 59 points\n" +
                "  - Gold:       60 - 79 points\n" +
                "  - Platinum:   80+ points";
        }

        private string GetScorecardSummary()
        {
            var categories = new[]
            {
                "[IP] Integrative Process: 1 pts (1 prerequisites)",
                "[LT] Location and Transportation: 16 pts (1 prerequisites)",
                "[SS] Sustainable Sites: 10 pts (2 prerequisites)",
                "[WE] Water Efficiency: 11 pts (3 prerequisites)",
                "[EA] Energy and Atmosphere: 33 pts (4 prerequisites)",
                "[MR] Materials and Resources: 13 pts (2 prerequisites)",
                "[EQ] Indoor Environmental Quality: 16 pts (2 prerequisites)",
                "[IN] Innovation: 6 pts",
                "[RP] Regional Priority: 4 pts"
            };
            return string.Join("\n", categories);
        }

        private string GetFullReadme()
        {
            return
                "============================================================\n" +
                "  LEED v4.1 Building Design and Construction (BD+C)\n" +
                "  v4.1 Beta - January 2019 | U.S. Green Building Council (USGBC)\n" +
                "============================================================\n\n" +
                GetOverview() + "\n\n" +
                "------------------------------------------------------------\n" +
                GetCertificationLevels() + "\n\n" +
                "------------------------------------------------------------\n" +
                "Scorecard (New Construction, 110 pts total):\n\n" +
                GetScorecardSummary() + "\n";
        }

        public override Guid ComponentGuid => new Guid("f26151c3-f5d8-468e-8f72-ed062efb000a");

        protected override Bitmap Icon
        {
            get
            {
                var assembly = System.Reflection.Assembly.GetExecutingAssembly();
                using (var stream = assembly.GetManifestResourceStream("Colugo.Resources.ReadMe_Icon.png"))
                {
                    if (stream == null) return null;
                    return new Bitmap(stream);
                }
            }
        }

        public override GH_Exposure Exposure => GH_Exposure.primary;
    }
}
