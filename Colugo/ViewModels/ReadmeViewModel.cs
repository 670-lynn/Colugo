using System.Collections.Generic;
using System.Linq;
using Colugo.Models;

namespace Colugo.ViewModels
{
    public class ReadmeViewModel : ViewModelBase
    {
        private readonly ReadmeModel _model = new ReadmeModel();

        public string DocumentTitle => _model.DocumentTitle;
        public string Version => _model.Version;
        public string IssuedBy => _model.IssuedBy;
        public int TotalPoints => _model.TotalPoints;

        public string Overview => _model.GetOverview();
        public string CertificationLevels => _model.GetCertificationLevels();

        public List<ScoreCategory> Scorecard => _model.GetScorecard();

        public string ScorecardSummary
        {
            get
            {
                var lines = Scorecard.Select(c => c.ToString());
                return string.Join("\n", lines);
            }
        }

        public string FullReadme
        {
            get
            {
                return
                    $"{'=',-60}\n" +
                    $"  {DocumentTitle}\n" +
                    $"  {Version} | {IssuedBy}\n" +
                    $"{'=',-60}\n\n" +
                    $"{Overview}\n\n" +
                    $"{'-',-60}\n" +
                    $"{CertificationLevels}\n\n" +
                    $"{'-',-60}\n" +
                    $"Scorecard (New Construction, {TotalPoints} pts total):\n\n" +
                    $"{ScorecardSummary}\n";
            }
        }
    }
}
