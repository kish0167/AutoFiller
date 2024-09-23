using System.Collections.ObjectModel;
using System.Windows.Input;
using AutoFiller.InternalLogic.BelTransSat;
using AutoFiller.InternalLogic.Excel;
using DataTracker.Excel;
using DataTracker.Utility;

namespace AutoFiller.UI
{
    public class MainWindowViewModel
    {
        private ExcelFileManager _excelFileManager;
        private MonthlyFileUpdater _monthlyFileUpdater;
        private StatisticsFiller _statisticsFiller;
        private SatDataFiller _satDataFiller;

        public ICommand UpdateCommand { get; }
        public ICommand FillStatisticsCommand { get; }
        public ICommand FillSatDataCommand { get; }

        public MainWindowViewModel(ExcelFileManager excelFileManager, MonthlyFileUpdater monthlyFileUpdater, StatisticsFiller statisticsFiller, SatDataFiller satDataFiller)
        {
            _excelFileManager = excelFileManager;
            _monthlyFileUpdater = monthlyFileUpdater;
            _statisticsFiller = statisticsFiller;
            _satDataFiller = satDataFiller;

            // Define commands for the UI buttons
            UpdateCommand = new RelayCommand(Update);
            FillStatisticsCommand = new RelayCommand(FillStatistics);
            FillSatDataCommand = new RelayCommand(async () => await FillSatData());
        }

        private void Update()
        {
            _monthlyFileUpdater.Update();
            _excelFileManager.SaveExcelFile();
        }

        private void FillStatistics()
        {
            _statisticsFiller.FillStatistics();
            _excelFileManager.SaveExcelFile();
        }

        private async Task FillSatData()
        {
            await _satDataFiller.Fill();
            _excelFileManager.SaveExcelFile();
        }
        
        public ObservableCollection<string> Logs => Logger.Logs;
    }
}