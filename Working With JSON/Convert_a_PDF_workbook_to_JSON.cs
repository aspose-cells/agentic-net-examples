using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsPdfToJson
{
    class Program
    {
        static void Main()
        {
            // Path to the source Excel file (generated from an Excel workbook)
            string excelPath = "input.xlsx";

            if (!File.Exists(excelPath))
            {
                Console.WriteLine($"File not found: {Path.GetFullPath(excelPath)}");
                return;
            }

            // Load the Excel file into a Workbook object.
            var loadOptions = new LoadOptions(LoadFormat.Xlsx);
            var workbook = new Workbook(excelPath, loadOptions);

            // Configure JSON save options.
            var jsonOptions = new JsonSaveOptions
            {
                AlwaysExportAsJsonObject = true,
                ExportNestedStructure = true,
                SkipEmptyRows = true
            };

            // Save the workbook as a JSON file.
            string jsonPath = "output.json";
            workbook.Save(jsonPath, jsonOptions);

            Console.WriteLine($"Workbook has been converted to JSON and saved at: {Path.GetFullPath(jsonPath)}");
        }
    }
}