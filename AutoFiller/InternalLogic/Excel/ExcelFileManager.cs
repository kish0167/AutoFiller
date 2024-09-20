using System.IO;
using DataTracker.Utility;
using OfficeOpenXml;

namespace AutoFiller.InternalLogic.Excel
{
    public class ExcelFileManager
    {
        private string _sourceFilePath;
        private string _archiveFilePath;
        private ExcelPackage _package;

        public ExcelPackage Package => _package;

        public ExcelFileManager(string sourcePath, string archiveFolder)
        {
            _sourceFilePath = sourcePath;
            _archiveFilePath = Path.Combine(archiveFolder, DateTime.Now.ToString("yyyy-MM-dd-HH-mm-ss") + ".xlsx");
        }

        public ExcelFileManager(string sourcePath)
        {
            _sourceFilePath = sourcePath;
            _archiveFilePath = "none";
        }
        
        public void LoadExcelFile()
        {
            _package = new ExcelPackage(new FileInfo(_sourceFilePath));
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            Logger.Log(_sourceFilePath + " loaded.");
        }
        
        public void ArchiveData()
        {
            using (var archivePackage = new ExcelPackage(new FileInfo(_archiveFilePath)))
            {
                foreach (var sourceWorksheet in Package.Workbook.Worksheets)
                {
                    archivePackage.Workbook.Worksheets.Add(sourceWorksheet.Name, sourceWorksheet);
                }

                archivePackage.Save();
                Logger.Log(_archiveFilePath + " saved.\n\n");
            }
        }

        public void SaveExcelFile()
        {
            _package.Save();
            Logger.Log(_sourceFilePath + " saved.\n\n");
        }
    }
}