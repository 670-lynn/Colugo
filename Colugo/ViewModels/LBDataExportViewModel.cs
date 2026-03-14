using System.Collections.Generic;
using Colugo.Models;

namespace Colugo.ViewModels
{
    public class LBDataExportViewModel : ViewModelBase
    {
        private readonly LBDataExportModel _model = new LBDataExportModel();

        public (string message, string savedPath) Execute(
            string path,
            List<object> headersFlat,
            List<List<double>> branches)
        {
            if (string.IsNullOrEmpty(path) || headersFlat == null || headersFlat.Count == 0
                || branches == null || branches.Count == 0)
                return ("Check inputs", null);

            return _model.WriteGroupedExcel(path, headersFlat, branches);
        }
    }
}
