using System;
using Grasshopper.Kernel;

namespace Colugo
{
    public class ColugoComponent : GH_Component
    {
        public ColugoComponent()
            : base("Colugo", "Col",
                "A sample Colugo component",
                "Colugo", "Utility")
        {
        }

        protected override void RegisterInputParams(GH_InputParamManager pManager)
        {
            pManager.AddTextParameter("Input", "I", "Input text", GH_ParamAccess.item, "Hello");
        }

        protected override void RegisterOutputParams(GH_OutputParamManager pManager)
        {
            pManager.AddTextParameter("Output", "O", "Output text", GH_ParamAccess.item);
        }

        protected override void SolveInstance(IGH_DataAccess DA)
        {
            string input = string.Empty;
            if (!DA.GetData(0, ref input)) return;

            DA.SetData(0, $"Colugo: {input}");
        }

        public override Guid ComponentGuid => new Guid("c4d5e6f7-a8b9-0123-4567-89abcdef0123");

        protected override System.Drawing.Bitmap Icon => null;
    }
}
