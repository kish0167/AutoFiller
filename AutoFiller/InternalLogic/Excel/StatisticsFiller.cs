using AutoFiller.InternalLogic.Excel;
using DataTracker.Internal;
using DataTracker.Utility;
using ExcelParser;
using OfficeOpenXml;

namespace DataTracker.Excel
{
    public class StatisticsFiller
    {
        private ExcelFileManager _excelFileManager;

        public StatisticsFiller(ExcelFileManager manager)
        {
            _excelFileManager = manager;
        }

        public void FillStatistics()
        {
            ExcelPackage package = _excelFileManager.Package;
            ExcelWorkbook workbook = package.Workbook;
            List<VehicleFuelStatistics> stats = new List<VehicleFuelStatistics>();
            
            foreach (var worksheet in workbook.Worksheets)
            {
                if (!ExcelSettings.IsVehicleSheet(worksheet))
                {
                    continue;
                }

                VehicleFuelStatistics vehicle = new VehicleFuelStatistics();
                TransferDataFromWorksheet(worksheet, vehicle);
                
                vehicle.CalculateTheoryConsumption();
                WriteConsumptionsToWorksheet(worksheet, vehicle);
                //Logger.Log(vehicle.Name + " successfully written");
            }
            
            Logger.Log("Statistics calculated and written.");
        }

        private void TransferDataFromWorksheet(ExcelWorksheet worksheet, VehicleFuelStatistics statistics)
        {
            statistics.Name = (string)ExcelSettings.NameCell(worksheet).Value;
            
            for (int i = 0; i < ExcelSettings.Rows; i++)
            {
                var temp = ExcelSettings.RefuelsDataCells(worksheet).GetCellValue<object>(i,0);
                if (temp==null)
                {
                    statistics.AddRefuel(0);
                }
                else
                {
                    statistics.AddRefuel((double)temp);
                }
                
                temp = ExcelSettings.TravelsDistancesCells(worksheet).GetCellValue<object>(i,0);
                
                if (temp==null)
                {
                    statistics.AddTravel(0);
                }
                else
                {
                    statistics.AddTravel((double)temp);
                }
            }
        }

        private void WriteConsumptionsToWorksheet(ExcelWorksheet worksheet, VehicleFuelStatistics statistics)
        {
            for (int i = 0; i < ExcelSettings.Rows && i<statistics.TheoryConsumptions.Count; i++)
            {
                ExcelSettings.ConsumptionDataCells(worksheet).TakeSingleCell(i,0).Value = statistics.TheoryConsumptions[i];
            }
        }
    }
}