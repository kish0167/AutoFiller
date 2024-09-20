using System.Windows;
using AutoFiller.InternalLogic.Excel;
using DataTracker.BelTransSat;
using DataTracker.Excel;

namespace AutoFiller
{
    public partial class App : Application
    {
        protected override void OnStartup(StartupEventArgs e)
        {
            base.OnStartup(e);
            
            // Load Excel settings and initialize services
            ExcelSettings.LoadSettings();
            
            ExcelFileManager excelFileManager = new ExcelFileManager(ExcelSettings.SourceFileLocation, ExcelSettings.ArchieveFolder);
            excelFileManager.LoadExcelFile();
            
            MonthlyFileUpdater updater = new MonthlyFileUpdater(excelFileManager);
            StatisticsFiller statisticsFiller = new StatisticsFiller(excelFileManager);
            SatDataFiller satDataFiller = new SatDataFiller(excelFileManager);

            // Passing services to the Main Window ViewModel (you'll create this class next)
            MainWindow mainWindow = new MainWindow
            {
                DataContext = new MainWindowViewModel(excelFileManager, updater, statisticsFiller, satDataFiller)
            };
            
            mainWindow.Show();
        }
    }
}