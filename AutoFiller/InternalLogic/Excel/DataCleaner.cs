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
                if (ExcelSettings.IsSatSpecialVehicleSheet(worksheet))
                {
                    CleanSatSpecSheet(worksheet);
                    continue;
                }

                if (ExcelSettings.IsSatDefaultVehicleSheet(worksheet))
                {
                    CleanSatDefaultSheet(worksheet);
                    continue;
                }

                if (ExcelSettings.IsVehicleSheet(worksheet))
                {
                    CleanVehicleSheet(worksheet);
                }
            }
        }
        
        private static void CleanSatSpecSheet(ExcelWorksheet worksheet)
        {
            CleanVehicleSheet(worksheet);
            ExcelSettings.SatRefuelsCells(worksheet).Value = null;
            ExcelSettings.SatSpecConsumptionCells(worksheet).Value = null;
            ExcelSettings.SatMachineHoursCells(worksheet).Value = null;
        }
        
        private static void CleanSatDefaultSheet(ExcelWorksheet worksheet)
        {
            CleanVehicleSheet(worksheet);
            ExcelSettings.SatTravelCells(worksheet).Value = null;
            ExcelSettings.SatConsumptionCells(worksheet).Value = null;
        }

        private static void CleanVehicleSheet(ExcelWorksheet worksheet)
        {
            ExcelSettings.NumericDataCells(worksheet).Value = null;
            ExcelSettings.ConstructionSitesCells(worksheet).Value = "-";
            ExcelSettings.ConsumptionDataCells(worksheet).Value = null;
        }
    }
}