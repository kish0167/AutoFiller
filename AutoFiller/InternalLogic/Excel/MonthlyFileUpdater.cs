using AutoFiller.InternalLogic.Excel;
using DataTracker.Utility;
using ExcelParser;
using OfficeOpenXml;

namespace DataTracker.Excel
{
    public class MonthlyFileUpdater
    {
        private ExcelFileManager _excelFileManager;
        public MonthlyFileUpdater(ExcelFileManager manager)
        {
            _excelFileManager = manager;
        }
        
        public void Update()
        {
            Logger.Log("Starting updating process...");
            ExcelWorkbook workbook = _excelFileManager.Package.Workbook;

            // Clean old data
            DataCleaner cleaner = new DataCleaner();
            cleaner.CleanOldData(workbook);
            Logger.Log("Old data cleaned.");

            // Update dates
            DateUpdater dateUpdater = new DateUpdater();
            dateUpdater.UpdateDates(workbook);
            Logger.Log("Dates updated.");

            Logger.Log("Updating complete.");
        }
    }
}