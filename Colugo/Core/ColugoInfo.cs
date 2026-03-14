using System;
using System.Drawing;
using Grasshopper.Kernel;

namespace Colugo.Core
{
    public class ColugoInfo : GH_AssemblyInfo
    {
        public override string Name => "Colugo";
        public override string Description => "A Grasshopper plugin for Rhino";
        public override Guid Id => new Guid("465dc75a-a7cd-4025-b0e0-dd6602d55da2");
        public override string AuthorName => "";
        public override string AuthorContact => "";

        public override Bitmap Icon
        {
            get
            {
                var assembly = System.Reflection.Assembly.GetExecutingAssembly();
                using (var stream = assembly.GetManifestResourceStream("Colugo.Resources.Colugo_Logo.png"))
                {
                    if (stream == null) return null;
                    return new Bitmap(stream);
                }
            }
        }
    }
}
