using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsMhtmlToJson
{
    class Program
    {
        static void Main()
        {
            // Path to the source workbook (using XLSX instead of MHTML for compatibility)
            string sourcePath = "input.xlsx";

            // Path for the resulting JSON file
            string jsonPath = "output.json";

            // Ensure the source workbook exists; if not, create a simple workbook and save it as XLSX
            if (!File.Exists(sourcePath))
            {
                var tempWb = new Workbook();
                tempWb.Worksheets[0].Cells["A1"].PutValue("Header");
                tempWb.Worksheets[0].Cells["B1"].PutValue("Value");
                tempWb.Worksheets[0].Cells["A2"].PutValue("Name");
                tempWb.Worksheets[0].Cells["B2"].PutValue("John Doe");
                tempWb.Save(sourcePath, SaveFormat.Xlsx);
            }

            // Load the workbook
            var workbook = new Workbook(sourcePath);

            // Configure JSON save options
            var jsonOptions = new JsonSaveOptions
            {
                AlwaysExportAsJsonObject = true,
                ExportEmptyCells = true,
                HasHeaderRow = true
            };

            // Save the workbook as JSON
            workbook.Save(jsonPath, jsonOptions);

            Console.WriteLine($"Workbook successfully converted to JSON at: {jsonPath}");
        }
    }
}