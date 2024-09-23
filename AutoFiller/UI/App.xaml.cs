using System.Windows;
using AutoFiller.InternalLogic.BelTransSat;
using AutoFiller.InternalLogic.Excel;
using DataTracker.Excel;

namespace AutoFiller.UI
{
    public partial class App : Application
    {
        protected override void OnStartup(StartupEventArgs e)
        {
            //base.OnStartup(e);
            ExcelSettings.LoadSettings();

            ExcelFileManager excelFileManager =
                new ExcelFileManager(ExcelSettings.SourceFileLocation, ExcelSettings.ArchieveFolder);
            excelFileManager.LoadExcelFile();

            MonthlyFileUpdater updater = new MonthlyFileUpdater(excelFileManager);
            StatisticsFiller statisticsFiller = new StatisticsFiller(excelFileManager);
            SatDataFiller satDataFiller = new SatDataFiller(excelFileManager);
            
            MainWindow mainWindow = new MainWindow
            {
                DataContext = new MainWindowViewModel(excelFileManager, updater, statisticsFiller, satDataFiller)
            };
            
            
            mainWindow.Show();
        }
    }
}