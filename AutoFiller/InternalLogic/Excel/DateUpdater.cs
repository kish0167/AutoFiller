using DataTracker.Excel;
using OfficeOpenXml;

namespace AutoFiller.InternalLogic.Excel
{
    public class DateUpdater
    {
        // Update dates in the Excel file
        public void UpdateDates(ExcelFileManager manager)
        {

            ExcelWorkbook workbook = manager.Package.Workbook;
            //GenerateCalcs(manager, workbook);
            //return;
            
            DateTime firstDayOfTheMonth = new DateTime(DateTime.Today.Year, DateTime.Today.Month, 1);
            firstDayOfTheMonth = firstDayOfTheMonth.AddMonths(2);
            foreach (var worksheet in workbook.Worksheets)
            {
                if (!ExcelSettings.IsVehicleSheet(worksheet))
                {
                    continue;
                }

                ExcelSettings.DateCells(worksheet).Value = null;
                
                DateTime dateBuf = firstDayOfTheMonth.AddMonths(-1);
                int i = 0;
                while (dateBuf < firstDayOfTheMonth)
                {
                    if (dateBuf.DayOfWeek.Equals(DayOfWeek.Saturday) || dateBuf.DayOfWeek.Equals(DayOfWeek.Sunday))
                    {
                        dateBuf = dateBuf.AddDays(1);
                        continue;
                    }
                    
                    ExcelSettings.DateCells(worksheet).TakeSingleCell(i,0).Value = dateBuf;
                    dateBuf = dateBuf.AddDays(1);
                    i++;
                }
            }
        }

        private void GenerateCalcs(ExcelFileManager manager, ExcelWorkbook workbook)
        {
            DateTime date = new DateTime(DateTime.Today.Year, DateTime.Today.Month, 1);

            for (int i = 0; i < 60; i++)
            {
                DateTime date2 = date;
                foreach (var worksheet in workbook.Worksheets)
                {
                    if (!ExcelSettings.IsCalcTableSheet(worksheet)) continue;
                    for (int j = 0; j < 31; j++)
                    {
                        ExcelSettings.CalcDateCells(worksheet).TakeSingleCell(0, j).Value = date2;
                        date2 = date.AddDays(j);
                    }
                }
                manager.ArchiveData("Табель на " + date.ToString("yyyy.MM"));
                date = date.AddMonths(1);
            }
        }
    }
}