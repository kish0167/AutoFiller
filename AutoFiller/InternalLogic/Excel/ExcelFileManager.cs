﻿using System.IO;
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

        public ExcelFileManager(string sourcePath, ExcelPackage package, string imString)
        {
            _sourceFilePath = sourcePath;
            _package = package;
            ImString = imString;
            _archiveFolder = "none";
        }

        public void LoadExcelFile()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            _package = new ExcelPackage(new FileInfo(_sourceFilePath));
            
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            Logger.Log(_sourceFilePath + " loaded.");
        }

        public void ArchiveData()
        {
            string archivePath = Path.Combine(_archiveFolder,"autoSave-" + 
                DateTime.Now.ToString("yyyy-MM-dd-HH-mm-ss") + ExcelSettings.Extension());
            File.WriteAllBytes(archivePath, Package.GetAsByteArray());
            Logger.Log("source archived.");
            return;
            using (var archivePackage = new ExcelPackage(new FileInfo(Path.Combine(_archiveFolder,
                       DateTime.Now.ToString("yyyy-MM-dd-HH-mm-ss") + ExcelSettings.Extension()))))
            {
                foreach (var sourceWorksheet in Package.Workbook.Worksheets)
                {
                    archivePackage.Workbook.Worksheets.Add(sourceWorksheet.Name, sourceWorksheet);
                }

                archivePackage.Save();
                Logger.Log(_archiveFolder + " saved.\n\n");
            }
        }

        public void ArchiveData(string fileName)
        {
            string archivePath = Path.Combine(_archiveFolder, fileName + ExcelSettings.Extension());
            File.WriteAllBytes(archivePath, Package.GetAsByteArray());
            Logger.Log(_archiveFolder + fileName + " saved.");
            return;
            
            
            using (var archivePackage =
                   new ExcelPackage(new FileInfo(Path.Combine(_archiveFolder, fileName + ExcelSettings.Extension()))))
            {
                foreach (var sourceWorksheet in Package.Workbook.Worksheets)
                {
                    archivePackage.Workbook.Worksheets.Add(sourceWorksheet.Name, sourceWorksheet);
                }

                archivePackage.Save();
                Logger.Log(_archiveFolder + fileName + " saved.\n\n");
            }
        }

        public void SaveExcelFile()
        {
            _package.Save();
            Logger.Log(_sourceFilePath + " saved.");
        }
    }
}