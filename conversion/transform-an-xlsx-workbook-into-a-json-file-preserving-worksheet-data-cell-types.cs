using System;
using Aspose.Cells;

namespace AsposeCellsJsonExport
{
    class Program
    {
        static void Main()
        {
            // Path to the source Excel workbook
            string sourcePath = "input.xlsx";

            // Path where the resulting JSON file will be saved
            string jsonPath = "output.json";

            // Load the workbook from the file system
            Workbook workbook = new Workbook(sourcePath);

            // Configure JSON save options to preserve data types, empty cells,
            // and hierarchical (nested) structure of the workbook.
            JsonSaveOptions saveOptions = new JsonSaveOptions
            {
                // Export the workbook as a JSON object even if it contains a single sheet
                AlwaysExportAsJsonObject = true,

                // Preserve empty cells as null values in the JSON output
                ExportEmptyCells = true,

                // Export the workbook using the Excel structural format (preserves hierarchy)
                ToExcelStruct = true,

                // Export data as a parent‑child hierarchy rather than a flat table
                ExportNestedStructure = true,

                // Keep original cell types (numbers, strings, dates, etc.)
                ExportAsString = false,

                // Skip empty rows is set to false to keep the original row order
                SkipEmptyRows = false
            };

            // Save the workbook as a JSON file using the configured options
            workbook.Save(jsonPath, saveOptions);

            Console.WriteLine($"Workbook \"{sourcePath}\" has been exported to JSON at \"{jsonPath}\".");
        }
    }
}