using System.IO;
using DataTracker.Utility;
using OfficeOpenXml;

namespace AutoFiller.InternalLogic.Excel
{
    public class ExcelFileManager
    {
        private string _sourceFilePath;
        private string _archiveFolder;
        private ExcelPackage _package;

        public ExcelPackage Package => _package;

        public string ImString;

        public ExcelFileManager(string sourcePath, string archiveFolder)
        {
            _sourceFilePath = sourcePath;
            _archiveFolder = archiveFolder;
        }

        public ExcelFileManager(string sourcePath)
        {
            _sourceFilePath = sourcePath;
            _archiveFolder = "none";
        }
        
        public void LoadExcelFile()
        {
            _package = new ExcelPackage(new FileInfo(_sourceFilePath));
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            Logger.Log(_sourceFilePath + " loaded.");
        }
        
        public void ArchiveData()
        {
            using (var archivePackage = new ExcelPackage(new FileInfo(Path.Combine(_archiveFolder, DateTime.Now.ToString("yyyy-MM-dd-HH-mm-ss") + ".xlsx"))))
            {
                foreach (var sourceWorksheet in Package.Workbook.Worksheets)
                {
                    archivePackage.Workbook.Worksheets.Add(sourceWorksheet.Name, sourceWorksheet);
                }

                archivePackage.Save();
                Logger.Log(_archiveFolder + " saved.\n\n");
            }
        }

        public void SaveExcelFile()
        {
            _package.Save();
            Logger.Log(_sourceFilePath + " saved.\n\n");
        }
    }
}