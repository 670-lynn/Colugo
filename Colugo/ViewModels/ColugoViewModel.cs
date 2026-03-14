using Colugo.Models;

namespace Colugo.ViewModels
{
    public class ColugoViewModel : ViewModelBase
    {
        private readonly ColugoModel _model = new ColugoModel();

        private string _input;
        public string Input
        {
            get => _input;
            set
            {
                if (SetProperty(ref _input, value))
                {
                    _model.Input = value;
                    Output = _model.Process();
                }
            }
        }

        private string _output;
        public string Output
        {
            get => _output;
            private set => SetProperty(ref _output, value);
        }
    }
}
