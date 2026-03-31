using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsJsonExport
{
    class Program
    {
        static void Main()
        {
            // Path to the source XLSM workbook
            string sourcePath = "input.xlsm";

            // Path where the JSON output will be saved
            string jsonPath = "output.json";

            // Load the XLSM workbook (creates and loads the document)
            Workbook workbook = new Workbook(sourcePath);

            // Configure JSON save options to preserve the full Excel structure
            JsonSaveOptions jsonOptions = new JsonSaveOptions
            {
                // Export the workbook as a JSON representation of the Excel structure
                ToExcelStruct = true,

                // Optional: include empty cells and treat the first row as header
                ExportEmptyCells = true,
                HasHeaderRow = true
            };

            // Save the workbook to JSON using the configured options (uses the provided save rule)
            workbook.Save(jsonPath, jsonOptions);

            Console.WriteLine($"Workbook '{sourcePath}' has been converted to JSON at '{jsonPath}'.");
        }
    }
}