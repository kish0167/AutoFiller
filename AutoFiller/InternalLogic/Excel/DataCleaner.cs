using DataTracker.Excel;
using OfficeOpenXml;

namespace AutoFiller.InternalLogic.Excel
{
    public class DataCleaner
    {
        public void CleanOldData(ExcelWorkbook workbook)
        {
            foreach (var worksheet in workbook.Worksheets)
            {
                if (!ExcelSettings.IsVehicleSheet(worksheet))
                {
                    continue;
                }   
                //ExcelSettings.NumericDataCells(worksheet).Value = null;   TODO:   DONT FORGET
                //ExcelSettings.ConstructionSitesCells(worksheet).Value = "-";
                ExcelSettings.ConsumptionDataCells(worksheet).Value = 0;
                ExcelSettings.SatTravelCells(worksheet).Value = 0;
                ExcelSettings.SatConsumptionCells(worksheet).Value = 0;
                ExcelSettings.ConsumptionDataCells(worksheet).Value = 0;
            }
        }
    }
}