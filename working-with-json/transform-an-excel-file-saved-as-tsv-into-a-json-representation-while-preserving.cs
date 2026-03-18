using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsTsvToJson
{
    class Program
    {
        static void Main()
        {
            // Path to the source TSV file
            string tsvPath = "input.tsv";

            // Path where the resulting JSON will be saved
            string jsonPath = "output.json";

            // Load the TSV file into a workbook using LoadOptions with Tsv format
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Tsv);
            Workbook workbook = new Workbook(tsvPath, loadOptions);

            // Configure JSON save options to preserve data fidelity
            JsonSaveOptions jsonOptions = new JsonSaveOptions
            {
                // Export empty cells as null to keep the exact layout
                ExportEmptyCells = true,
                // Treat the first row as header if present
                HasHeaderRow = true,
                // Export cell values as strings to avoid type conversion loss
                ExportAsString = true,
                // Preserve the Excel structure (sheets, rows, columns) in JSON
                ToExcelStruct = true,
                // Optional: indent for readability
                Indent = "  "
            };

            // Save the workbook as a JSON file using the configured options
            workbook.Save(jsonPath, jsonOptions);

            Console.WriteLine($"TSV file '{tsvPath}' has been converted to JSON at '{jsonPath}'.");
        }
    }
}