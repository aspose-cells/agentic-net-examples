using System;
using Aspose.Cells;
using Aspose.Cells.Json;

namespace AsposeCellsJsonExport
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source XLTX template file
            string sourcePath = "template.xltx";

            // Path where the resulting JSON file will be saved
            string jsonPath = "template.json";

            // Load the XLTX workbook (lifecycle: load)
            Workbook workbook = new Workbook(sourcePath);

            // Configure JSON save options to preserve Excel structure and formatting
            JsonSaveOptions jsonOptions = new JsonSaveOptions
            {
                // Convert the workbook to a JSON structure that mirrors the Excel file
                ToExcelStruct = true,

                // Include style information for each cell
                ExportStylePool = true,

                // Export empty cells as null to keep the layout intact
                ExportEmptyCells = true,

                // Export cell values as strings to retain original formatting (e.g., dates, numbers)
                ExportAsString = true
            };

            // Save the workbook as JSON using the configured options (lifecycle: save)
            workbook.Save(jsonPath, jsonOptions);

            Console.WriteLine($"Workbook '{sourcePath}' has been exported to JSON at '{jsonPath}'.");
        }
    }
}