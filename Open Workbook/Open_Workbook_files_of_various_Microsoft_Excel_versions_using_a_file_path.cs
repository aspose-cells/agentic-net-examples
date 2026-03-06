using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsOpenVariousFormats
{
    class Program
    {
        static void Main()
        {
            // Create a temporary directory to store sample files
            string tempDir = Path.Combine(Path.GetTempPath(), "AsposeCellsSamples");
            Directory.CreateDirectory(tempDir);

            // Prepare sample workbook to be saved in various formats
            string baseFileName = Path.Combine(tempDir, "Sample");
            CreateSampleWorkbook(baseFileName + ".xlsx");

            var files = new List<string>
            {
                baseFileName + ".xlsx",
                baseFileName + ".xls",
                baseFileName + ".xlsm",
                baseFileName + ".csv",
                baseFileName + ".ods"
            };

            foreach (var filePath in files)
            {
                using (Workbook workbook = OpenWorkbook(filePath))
                {
                    Console.WriteLine($"Opened '{Path.GetFileName(filePath)}' successfully.");
                    Console.WriteLine($"  Worksheets: {workbook.Worksheets.Count}");
                    Console.WriteLine($"  File format: {workbook.FileFormat}");
                    Console.WriteLine($"  Contains macro: {workbook.HasMacro}");
                    Console.WriteLine();
                }
            }
        }

        private static void CreateSampleWorkbook(string xlsxPath)
        {
            // Create a simple workbook and save it in all required formats
            var wb = new Workbook();
            var ws = wb.Worksheets[0];
            ws.Cells["A1"].PutValue("Hello");
            ws.Cells["B1"].PutValue("World");
            wb.Save(xlsxPath, SaveFormat.Xlsx);

            // Save as other formats
            wb.Save(Path.ChangeExtension(xlsxPath, ".xls"), SaveFormat.Excel97To2003);
            wb.Save(Path.ChangeExtension(xlsxPath, ".xlsm"), SaveFormat.Xlsm);
            wb.Save(Path.ChangeExtension(xlsxPath, ".csv"), SaveFormat.CSV);
            wb.Save(Path.ChangeExtension(xlsxPath, ".ods"), SaveFormat.ODS);
        }

        private static Workbook OpenWorkbook(string filePath)
        {
            LoadFormat loadFormat = GetLoadFormatFromExtension(Path.GetExtension(filePath));

            if (loadFormat != LoadFormat.Auto)
            {
                var loadOptions = new LoadOptions(loadFormat);
                return new Workbook(filePath, loadOptions);
            }

            return new Workbook(filePath);
        }

        private static LoadFormat GetLoadFormatFromExtension(string extension)
        {
            if (string.IsNullOrEmpty(extension))
                return LoadFormat.Auto;

            switch (extension.ToLowerInvariant())
            {
                case ".xlsx":
                    return LoadFormat.Xlsx;
                case ".xls":
                    return LoadFormat.Excel97To2003;
                case ".xlsm":
                    return LoadFormat.Xlsx; // Macro-enabled workbooks are loaded using Xlsx format
                case ".csv":
                    return LoadFormat.Csv;
                case ".ods":
                    return LoadFormat.Ods;
                case ".xlsb":
                    return LoadFormat.Xlsb;
                case ".tsv":
                case ".txt":
                    return LoadFormat.Tsv;
                case ".html":
                case ".htm":
                    return LoadFormat.Html;
                case ".mht":
                case ".mhtml":
                    return LoadFormat.MHtml;
                default:
                    return LoadFormat.Auto;
            }
        }
    }
}