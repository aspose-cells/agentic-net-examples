using System;
using Aspose.Cells;

namespace AsposeCellsJsonConversion
{
    class Program
    {
        static void Main()
        {
            // Path to the source SpreadsheetML (Excel) file
            string sourcePath = "input.xlsx";

            // Desired output JSON file path
            string jsonPath = "output.json";

            // Load the Excel workbook from the source file
            Workbook workbook = new Workbook(sourcePath);

            // Configure JSON save options
            JsonSaveOptions jsonOptions = new JsonSaveOptions
            {
                // Export the workbook as a JSON object even if it contains a single worksheet
                AlwaysExportAsJsonObject = true,

                // Preserve the hierarchical structure of the workbook (worksheets, tables, etc.)
                ExportNestedStructure = true,

                // Skip empty rows to produce a cleaner JSON output
                SkipEmptyRows = true
            };

            // Save the workbook as JSON using the configured options
            workbook.Save(jsonPath, jsonOptions);

            Console.WriteLine($"Excel file '{sourcePath}' has been successfully converted to JSON at '{jsonPath}'.");
        }
    }
}