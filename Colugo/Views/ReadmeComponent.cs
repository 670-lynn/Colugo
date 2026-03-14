using System;
using System.Drawing;
using Grasshopper.Kernel;
using Colugo.ViewModels;

namespace Colugo.Views
{
    public class ReadmeComponent : GH_Component
    {
        private readonly ReadmeViewModel _viewModel = new ReadmeViewModel();

        public ReadmeComponent()
            : base("LEED v4.1 Readme", "Readme",
                "Displays an overview of LEED v4.1 BD+C, including credit categories, scorecard, and certification levels.",
                "Colugo", "LEED")
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
            DA.SetData(0, _viewModel.Overview);
            DA.SetData(1, _viewModel.CertificationLevels);
            DA.SetData(2, _viewModel.ScorecardSummary);
            DA.SetData(3, _viewModel.FullReadme);
        }

        public override Guid ComponentGuid => new Guid("a1b2c3d4-0001-4000-8000-000000000001");

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
