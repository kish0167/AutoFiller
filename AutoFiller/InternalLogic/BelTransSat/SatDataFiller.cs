using AutoFiller.InternalLogic.Excel;
using DataTracker.Excel;
using DataTracker.Utility;
using ExcelParser.BelTransSat;
using OfficeOpenXml;

namespace DataTracker.BelTransSat;

public class SatDataFiller
{
    private ExcelFileManager _excelFileManager;
    private readonly Dictionary<string, string> _vehiclesDictionary;

    public SatDataFiller(ExcelFileManager manager)
    {
        _excelFileManager = manager;
        _vehiclesDictionary = new Dictionary<string, string>()
        {
            {"Case", "Case 695ST"},
            {"Газель", "Газель АК 6826-2"},
            {"Iveco", "ивеко АК 8561-2"},
            {"Крафтер", "Крафтер Аl 6488-2"},
            {"МАЗ", "МАЗ"},
            {"Телескоп.погрузчик", "Подъёмник  8829 ВК-2"},
            {"New Holland", "Экскаватор"}
        };
    }
    
    public async Task Fill()
    {
        ExcelPackage package = _excelFileManager.Package;
        ApiClient client = new ApiClient(GetTokenFromFile());
        RootObject SatDataObject = new RootObject();
        DateTime currentDate;

        foreach (var worksheet in package.Workbook.Worksheets)
        {
            if (!ExcelSettings.IsSatDefaultVehicleSheet(worksheet))
            {
                continue;
            }

            for (int i = 0; i < ExcelSettings.Rows; i++)
            {
                if (IsCellFilled(ExcelSettings.SatTravelCells(worksheet), i, 0))
                {
                    continue;
                }

                string vehicleName = ExcelSettings.NameCell(worksheet).GetCellValue<string>();

                currentDate = ExcelSettings.DateCells(worksheet).GetCellValue<DateTime>(i, 0);
                
                if (!IsValidDate(currentDate))
                {
                    continue;
                }
                
                SatDataObject = await client.GetVehiclesInfo(currentDate);

                if (!_vehiclesDictionary.TryGetValue(vehicleName, out string id))
                {
                    Logger.Log(vehicleName + " not found in api response");
                    continue;
                }
                
                VehicleObject vehicle = SatDataObject.FindWithId(id);
                
                ExcelSettings.SatTravelCells(worksheet).TakeSingleCell(i, 0).Value = vehicle.GetTravelDistance();
                ExcelSettings.SatConsumptionCells(worksheet).TakeSingleCell(i, 0).Value = vehicle.GetFuelUsed();
            }
        }
        Logger.Log("All satellite data loaded to file.\n");
    }
    
    private bool IsCellFilled(ExcelRange cells, int row, int column)
    {
        double cell = cells.GetCellValue<double>(row, column);
        return cell != 0;
    }
    
    private string GetTokenFromFile()
    {
        return TxtHandler.ReadFile("token.txt");
    }

    private bool IsValidDate(DateTime date)
    {
        return DateTime.Compare(date, DateTime.Today) < 0 && DateTime.Compare(date, ExcelSettings.originDate) >= 0;
    }
}