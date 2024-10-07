using AutoFiller.InternalLogic.Excel;
using DataTracker.Excel;
using DataTracker.Utility;
using ExcelParser.BelTransSat;
using OfficeOpenXml;

namespace AutoFiller.InternalLogic.BelTransSat;

public class SatDataFiller(ExcelFileManager manager)
{
    

    public async Task Fill()
    {
        ExcelPackage package = manager.Package;
        ApiClient client = new ApiClient(GetTokenFromFile());
        VehicleDictionary.LoadDictionary();

        foreach (var worksheet in package.Workbook.Worksheets)
        {
            if (ExcelSettings.IsSatDefaultVehicleSheet(worksheet))
            {
                await FillDefaultSheet(worksheet, client);
                continue;
            }

            if (ExcelSettings.IsSatSpecialVehicleSheet(worksheet))
            {
                await FillSpecSheet(worksheet, client);
            }
        }

        Logger.Log("All satellite data loaded to file.\n");
    }

    private async Task FillSpecSheet(ExcelWorksheet worksheet, ApiClient client)
    {
        for (int i = 0; i < ExcelSettings.Rows; i++)
        {
            if (IsCellFilled(ExcelSettings.SatMachineHoursCells(worksheet), i, 0))
            {
                continue;
            }

            await FillRowSpec(worksheet, i, client);
        }
    }

    private async Task FillDefaultSheet(ExcelWorksheet worksheet, ApiClient client)
    {
        for (int i = 0; i < ExcelSettings.Rows; i++)
        {
            if (IsCellFilled(ExcelSettings.SatTravelCells(worksheet), i, 0))
            {
                continue;
            }

            await FillRowDefault(worksheet, i, client);
        }
    }

    private async Task FillRowDefault(ExcelWorksheet worksheet, int i, ApiClient client)
    {
        VehicleObject vehicle = await GetDataForRow(worksheet, i, client);
        ExcelSettings.SatTravelCells(worksheet).TakeSingleCell(i, 0).Value = vehicle.GetTravelDistance();
        ExcelSettings.SatConsumptionCells(worksheet).TakeSingleCell(i, 0).Value = vehicle.GetFuelUsed();
    }

    private async Task FillRowSpec(ExcelWorksheet worksheet, int i, ApiClient client)
    {
        VehicleObject vehicle = await GetDataForRow(worksheet, i, client);
        ExcelSettings.SatSpecConsumptionCells(worksheet).TakeSingleCell(i, 0).Value = vehicle.GetFuelUsed();
        ExcelSettings.SatMachineHoursCells(worksheet).TakeSingleCell(i, 0).Value = vehicle.GetMachineHours();
        ExcelSettings.SatRefuelsCells(worksheet).TakeSingleCell(i, 0).Value = vehicle.GetRefuel();
    }

    private async Task<VehicleObject> GetDataForRow(ExcelWorksheet worksheet, int i, ApiClient client)
    {
        string vehicleName = ExcelSettings.NameCell(worksheet).GetCellValue<string>();
        DateTime currentDate = ExcelSettings.DateCells(worksheet).GetCellValue<DateTime>(i, 0);

        if (!IsValidDate(currentDate))
        {
            return new VehicleObject();
        }

        RootObject satDataObject = await client.GetVehiclesInfo(currentDate);

        if (!VehicleDictionary.Dictionary.TryGetValue(vehicleName, out string? id))
        {
            Logger.Log(vehicleName + " not found in api response");
            return new VehicleObject();
        }

        VehicleObject vehicle = satDataObject.FindWithId(id);
        return vehicle;
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
        return DateTime.Compare(date, DateTime.Today) < 0 && DateTime.Compare(date, ExcelSettings.OriginDate) >= 0;
    }
}