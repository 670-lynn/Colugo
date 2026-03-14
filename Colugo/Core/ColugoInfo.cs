using System;
using System.Drawing;
using Grasshopper.Kernel;

namespace Colugo.Core
{
    public class ColugoInfo : GH_AssemblyInfo
    {
        public override string Name => "Colugo";
        public override string Description => "A Grasshopper plugin for Rhino";
        public override Guid Id => new Guid("b3a7c1d2-e4f5-6789-0abc-def123456789");
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
