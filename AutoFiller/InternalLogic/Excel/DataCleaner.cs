using DataTracker.Excel;
using OfficeOpenXml;
using static AutoFiller.InternalLogic.Excel.ExcelSettings;

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
                    continue;
                }

                if (ExcelSettings.IsCalcCalcSheet(worksheet))
                {
                    CleanCalcCalcSheet(worksheet);
                    continue;
                }

                if (ExcelSettings.IsCalcTableSheet(worksheet))
                {
                    CleanCalcTableSheet(worksheet);
                    continue;
                }

                if (ExcelSettings.IsCalcObjectSheet(worksheet))
                {
                    CleanCalcObjectSheet(worksheet);
                    continue;
                }
            }
        }

        private void CleanCalcObjectSheet(ExcelWorksheet worksheet)
        {
            CalcObjCells(worksheet).Value = null;
        }

        private void CleanCalcTableSheet(ExcelWorksheet worksheet)
        {
            CalcTableCells(worksheet).Value = null;
        }

        private void CleanCalcCalcSheet(ExcelWorksheet worksheet)
        {
            for(int i=0; i<CalcKtuCells(worksheet).Columns; i+=3)
            {
                int j = 0;
                while (worksheet.Cells[CalcCalcPeople].TakeSingleCell(j, 0).Value != null ||
                       worksheet.Cells[CalcCalcPeople].TakeSingleCell(j + 1, 0).Value != null)
                {
                    if (worksheet.Cells[CalcCalcPeople].TakeSingleCell(j, 0).Value != null)
                    {
                        CalcKtuCells(worksheet).TakeSingleCell(j, i).Value = 1;
                    }

                    j++;
                }
            }
        }

        private static void CleanSatSpecSheet(ExcelWorksheet worksheet)
        {
            CleanVehicleSheet(worksheet);
            SatRefuelsCells(worksheet).Value = null;
            SatSpecConsumptionCells(worksheet).Value = null;
            SatMachineHoursCells(worksheet).Value = null;
        }

        private static void CleanSatDefaultSheet(ExcelWorksheet worksheet)
        {
            CleanVehicleSheet(worksheet);
            SatTravelCells(worksheet).Value = null;
            SatConsumptionCells(worksheet).Value = null;
        }

        private static void CleanVehicleSheet(ExcelWorksheet worksheet)
        {
            NumericDataCells(worksheet).Value = null;
            ConstructionSitesCells(worksheet).Value = "-";
            ConsumptionDataCells(worksheet).Value = null;
        }
    }
}