using Prism.Mvvm;

namespace MarketingCourseTool.ViewModel
{
    
    public class ProgressViewModel : BindableBase
    {
        private float _progressValue;
        private string _progressTitle;
        private float _currentValue;
        private float _maxValue;
        private string _title;

        public ProgressViewModel(string title, float max)
        {
            _progressValue = 0f;
            _currentValue = 0f;
            _maxValue = max;
            _title = title;
            ProgressTitle = $"{title} : {_currentValue}/{_maxValue}";
        }

        public void Progress(float progressStep)
        {
            _currentValue += progressStep;
            ProgressValue = _currentValue / _maxValue;
            ProgressTitle = $"{_title} : {_currentValue}/{_maxValue}";
            RaisePropertyChanged(nameof(ProgressValue));
        }

        public float ProgressValue
        {
            get => _progressValue;
            set => SetProperty(ref _progressValue, value);
        }

        public string ProgressTitle
        {
            get => _progressTitle;
            set => SetProperty(ref _progressTitle, value);
        }
    }
}
