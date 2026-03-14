using System;
using System.Collections.Generic;
using Grasshopper.Kernel;
using Grasshopper.Kernel.Data;
using Grasshopper.Kernel.Types;
using Colugo.ViewModels;

namespace Colugo.Views
{
    public class LBDataExportComponent : GH_Component
    {
        private readonly LBDataExportViewModel _viewModel = new LBDataExportViewModel();

        public LBDataExportComponent()
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

            // 比照 Python: headers_flat 展平
            var headersFlat = new List<object>();
            foreach (var goo in headerGoos)
            {
                if (goo is GH_ObjectWrapper wrapper)
                    headersFlat.Add(wrapper.Value);
                else
                    headersFlat.Add(goo);
            }

            // 比照 Python: data_tree.Branch(i) 轉成 List<List<double>>
            var branches = new List<List<double>>();
            for (int i = 0; i < valueTree.PathCount; i++)
            {
                var branch = valueTree.Branches[i];
                var vals = new List<double>();
                foreach (var item in branch)
                    vals.Add(item.Value);
                branches.Add(vals);
            }

            var (message, savedPath) = _viewModel.Execute(filePath, headersFlat, branches);

            DA.SetData(0, message);
            DA.SetData(1, savedPath);
        }

        public override Guid ComponentGuid => new Guid("a1b2c3d4-0002-4000-8000-000000000002");

        protected override System.Drawing.Bitmap Icon => null;

        public override GH_Exposure Exposure => GH_Exposure.primary;
    }
}
