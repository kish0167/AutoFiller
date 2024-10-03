using System.Windows.Media.Animation;
using DataTracker.Utility;
using OfficeOpenXml;

namespace AutoFiller.InternalLogic.Excel
{
    enum ConfigTypes
    {
        NumericDataCells = 0,
        ConstructionSitesCells = 1,
        DateCells = 2,
        TravelDistancesCells = 3,
        RefuelsDataCells = 4,
        ConsumptionDataCells = 5,
        NameCells = 6,
        SatMachineHoursCells = 7,
        SatTravelCells = 8,
        SatConsumptionCells = 9,
        SatRefuelsCells = 10,
        SatTagCell = 11,
        SatSpecConsumptionCells = 12,
        CalcDateCells = 13,
        CalcTableCells = 14,
        CalcKtuCells = 15,
        CalcTableSheetName = 16,
        CalcObjectSheetName = 17,
        CalcCalcSheetName = 18,
        CalcObjCells = 19
    }

    public static class ExcelSettings
    {
        private static List<string>? _locationsInExcel;
        private static string _sourceFileLocation = null!;

        public static string SourceFileLocation => _sourceFileLocation;
        public static string ArchieveFolder => _archieveFolder;

        private static string _archieveFolder;
        
        private const string ConfigFileName = "config.txt";

        public static ExcelRange NumericDataCells(ExcelWorksheet worksheet)
        {
            return worksheet.Cells[_locationsInExcel[(int)ConfigTypes.NumericDataCells]];
        }

        public static ExcelRange ConstructionSitesCells(ExcelWorksheet worksheet)
        {
            return worksheet.Cells[_locationsInExcel[(int)ConfigTypes.ConstructionSitesCells]];
        }

        public static ExcelRange DateCells(ExcelWorksheet worksheet)
        {
            return worksheet.Cells[_locationsInExcel[(int)ConfigTypes.DateCells]];
        }

        public static ExcelRange TravelsDistancesCells(ExcelWorksheet worksheet)
        {
            return worksheet.Cells[_locationsInExcel[(int)ConfigTypes.TravelDistancesCells]];
        }

        public static ExcelRange RefuelsDataCells(ExcelWorksheet worksheet)
        {
            return worksheet.Cells[_locationsInExcel[(int)ConfigTypes.RefuelsDataCells]];
        }

        public static ExcelRange ConsumptionDataCells(ExcelWorksheet worksheet)
        {
            return worksheet.Cells[_locationsInExcel[(int)ConfigTypes.ConsumptionDataCells]];
        }

        public static ExcelRange NameCell(ExcelWorksheet worksheet)
        {
            return worksheet.Cells[_locationsInExcel[(int)ConfigTypes.NameCells]];
        }
        
        public static ExcelRange SatTravelCells(ExcelWorksheet worksheet)
        {
            return worksheet.Cells[_locationsInExcel[(int)ConfigTypes.SatTravelCells]];
        }

        public static ExcelRange SatConsumptionCells(ExcelWorksheet worksheet)
        {
            return worksheet.Cells[_locationsInExcel[(int)ConfigTypes.SatConsumptionCells]];
        }
        
        public static ExcelRange SatRefuelsCells(ExcelWorksheet worksheet)
        {
            return worksheet.Cells[_locationsInExcel[(int)ConfigTypes.SatRefuelsCells]];
        }
        
        public static ExcelRange SatMachineHoursCells(ExcelWorksheet worksheet)
        {
            return worksheet.Cells[_locationsInExcel[(int)ConfigTypes.SatMachineHoursCells]];
        }

        public static ExcelRange SatSpecConsumptionCells(ExcelWorksheet worksheet)
        {
            return worksheet.Cells[_locationsInExcel[(int)ConfigTypes.SatSpecConsumptionCells]];
        }
        
        public const int Rows = 23;
        public static readonly DateTime OriginDate = new DateTime(2024, 1, 1);
        public const string CalcCalcHeaders = "E6:AZ6";
        public const string CalcCalcPeople = "C8:C100";
        public const string CalcMonthLabel = "C5";
        

        public static bool IsVehicleSheet(ExcelWorksheet worksheet)
        {
            return worksheet.Cells[_locationsInExcel[(int)ConfigTypes.NameCells]].Value != null;
        }
        
        public static bool IsSatDefaultVehicleSheet(ExcelWorksheet worksheet)
        {
            return IsVehicleSheet(worksheet) && worksheet.Cells[_locationsInExcel[(int)ConfigTypes.SatTagCell]].GetCellValue<string>() == "sat-default";
        }
        
        public static bool IsSatSpecialVehicleSheet(ExcelWorksheet worksheet)
        {
            return IsVehicleSheet(worksheet) && worksheet.Cells[_locationsInExcel[(int)ConfigTypes.SatTagCell]].GetCellValue<string>() == "sat-spec";
        }
        
        public static ExcelRange CalcDateCells(ExcelWorksheet worksheet)
        {
            return worksheet.Cells[_locationsInExcel[(int)ConfigTypes.CalcDateCells]];
        }
        
        public static ExcelRange CalcObjCells(ExcelWorksheet worksheet)
        {
            return worksheet.Cells[_locationsInExcel[(int)ConfigTypes.CalcObjCells]];
        }
        
        public static ExcelRange CalcKtuCells(ExcelWorksheet worksheet)
        {
            return worksheet.Cells[_locationsInExcel[(int)ConfigTypes.CalcKtuCells]];
        }
        
        public static ExcelRange CalcTableCells(ExcelWorksheet worksheet)
        {
            return worksheet.Cells[_locationsInExcel[(int)ConfigTypes.CalcTableCells]];
        }

        public static bool IsCalcTableSheet(ExcelWorksheet worksheet)
        {
            return worksheet.Name == _locationsInExcel[(int)ConfigTypes.CalcTableSheetName];
        }
        
        public static bool IsCalcCalcSheet(ExcelWorksheet worksheet)
        {
            return worksheet.Name == _locationsInExcel[(int)ConfigTypes.CalcCalcSheetName];
        }
        
        public static bool IsCalcObjectSheet(ExcelWorksheet worksheet)
        {
            return worksheet.Name == _locationsInExcel[(int)ConfigTypes.CalcObjectSheetName];
        }
        
        
        public static void LoadSettings()
        {
            Logger.Log("Loading config file...");
            string configs = TxtHandler.ReadFile(ConfigFileName);

            string[] lines = configs.Split('\n');

            _locationsInExcel = new List<string>();
            for (int i = 0; i < Enum.GetNames(typeof(ConfigTypes)).Length; i++)
            {
                _locationsInExcel.Add("A1");
            }
            
            foreach (string option in lines)
            {
                foreach (var configType in Enum.GetValues(typeof(ConfigTypes)))
                {
                    if (option.Contains(configType.ToString()))
                    {
                        string[] separateOption = option.Replace('\r', '-').Split('-');
                        ConfigTypes a = (ConfigTypes)Enum.Parse(typeof(ConfigTypes), separateOption[0]);
                        _locationsInExcel[(int)a] = separateOption[1];
                        break;
                    }
                }
                
                if (option.Contains("SourceFileLocation"))
                {
                    string[] separateOption = option.Replace('\r', '-').Split('-');
                    _sourceFileLocation = separateOption[1];
                }
                                    
                if (option.Contains("ArchiveFolder"))
                {
                    string[] separateOption = option.Replace('\r', '-').Split('-');
                    _archieveFolder = separateOption[1];
                }
            }
            Logger.Log("Configs loaded.");
        }
    }
}
