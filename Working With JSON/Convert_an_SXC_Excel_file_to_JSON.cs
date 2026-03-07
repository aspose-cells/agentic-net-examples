using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsSxcToJson
{
    class Program
    {
        static void Main()
        {
            // Path to the source SXC file (relative to the executable directory)
            string sourcePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "input.sxc");

            // Path where the resulting JSON will be saved
            string jsonPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "output.json");

            Workbook workbook;

            if (File.Exists(sourcePath))
            {
                // Load the existing SXC file
                var loadOptions = new LoadOptions(LoadFormat.Sxc);
                workbook = new Workbook(sourcePath, loadOptions);
            }
            else
            {
                // Create a sample workbook if the SXC file is missing
                workbook = new Workbook();
                var sheet = workbook.Worksheets[0];
                sheet.Name = "Sample";

                // Add header row
                sheet.Cells["A1"].PutValue("ID");
                sheet.Cells["B1"].PutValue("Name");
                sheet.Cells["C1"].PutValue("Score");

                // Add some data rows
                sheet.Cells["A2"].PutValue(1);
                sheet.Cells["B2"].PutValue("Alice");
                sheet.Cells["C2"].PutValue(85);

                sheet.Cells["A3"].PutValue(2);
                sheet.Cells["B3"].PutValue("Bob");
                sheet.Cells["C3"].PutValue(92);
            }

            // Configure JSON save options
            var jsonOptions = new JsonSaveOptions
            {
                // Export the workbook as a JSON object even if it contains a single worksheet
                AlwaysExportAsJsonObject = true,

                // Export empty cells and treat the first row as header
                ExportEmptyCells = true,
                HasHeaderRow = true
            };

            // Save the workbook as JSON using the configured options
            workbook.Save(jsonPath, jsonOptions);

            Console.WriteLine($"Workbook has been converted to JSON at '{jsonPath}'.");
        }
    }
}