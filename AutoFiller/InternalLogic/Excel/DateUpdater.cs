using DataTracker.Excel;
using OfficeOpenXml;

namespace AutoFiller.InternalLogic.Excel
{
    public class DateUpdater
    {
        // Update dates in the Excel file
        public void UpdateDates(ExcelWorkbook workbook)
        {
            DateTime firstDayOfTheMonth = new DateTime(DateTime.Today.Year, DateTime.Today.Month, 1);
            firstDayOfTheMonth = firstDayOfTheMonth.AddMonths(2);
            foreach (var worksheet in workbook.Worksheets)
            {
                
                if (!ExcelSettings.IsVehicleSheet(worksheet))
                {
                    continue;
                }
                
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
    }
}