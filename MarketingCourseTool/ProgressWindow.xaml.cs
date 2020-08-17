using System;
using System.ComponentModel;
using System.Windows;
using MarketingCourseTool.ViewModel;

namespace MarketingCourseTool
{
    /// <summary>
    /// ProgressWindow.xaml 的互動邏輯
    /// </summary>
    public partial class ProgressWindow : Window
    {
        private BackgroundWorker _worker;
        public ProgressWindow(string title, float max, Action work)
        {
            InitializeComponent();
            _worker = new BackgroundWorker();
            DataContext = new ProgressViewModel(title, max);

            _worker.ProgressChanged += (s, e) =>
            {
                Progress(e.ProgressPercentage);
            };
            _worker.RunWorkerCompleted += (s, e) =>
            {
                Close();
            };
            _worker.DoWork += (s, e) => { work(); };
        }



        public void Progress(float step)
        {
            ((ProgressViewModel)DataContext).Progress(step);
        }

        public void WorkerProgress(float percentage)
        {
            _worker.ReportProgress((int)(percentage * 100));
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            _worker.RunWorkerAsync();
        }
    }
}
