using System;
using System.IO;
using Aspose.Cells;

namespace JsonWorkbookConversion
{
    class Program
    {
        static void Main()
        {
            // Path to the source JSON file that represents a workbook
            string inputJsonPath = "inputWorkbook.json";

            Workbook workbook;

            if (File.Exists(inputJsonPath))
            {
                // Load options – enable MultipleWorksheets if the JSON contains multiple arrays
                JsonLoadOptions loadOptions = new JsonLoadOptions
                {
                    MultipleWorksheets = true
                };

                // Load the JSON workbook into an Aspose.Cells Workbook object
                workbook = new Workbook(inputJsonPath, loadOptions);
            }
            else
            {
                // Create a sample workbook if the JSON file does not exist
                workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Name = "SampleSheet";

                // Add some sample data
                sheet.Cells["A1"].PutValue("ID");
                sheet.Cells["B1"].PutValue("Name");
                sheet.Cells["A2"].PutValue(1);
                sheet.Cells["B2"].PutValue("Alice");
                sheet.Cells["A3"].PutValue(2);
                sheet.Cells["B3"].PutValue("Bob");
            }

            // Configure JSON save options for the output
            JsonSaveOptions saveOptions = new JsonSaveOptions
            {
                // Export as a JSON object even if there is only one worksheet
                AlwaysExportAsJsonObject = true,
                // Preserve the hierarchical (parent‑child) structure if present
                ExportNestedStructure = true,
                // Skip rows that contain no data
                SkipEmptyRows = true,
                // Export empty cells as null values
                ExportEmptyCells = true,
                // Treat the workbook as a JSON struct (useful for range exports)
                ToExcelStruct = true
            };

            // Path for the resulting JSON file
            string outputJsonPath = "outputWorkbook.json";

            // Save the workbook back to JSON using the configured options
            workbook.Save(outputJsonPath, saveOptions);

            // Optional: display a confirmation
            Console.WriteLine($"Workbook successfully saved as JSON at: {outputJsonPath}");
        }
    }
}