using System;
using Aspose.Cells;

namespace AsposeCellsCsvToJson
{
    class Program
    {
        static void Main()
        {
            // Path to the source CSV file (SXC CSV)
            string csvPath = "input.csv";

            // Path for the output JSON file
            string jsonPath = "output.json";

            // Create load options specifying that the source format is CSV
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Csv);

            // Load the CSV file into a Workbook object
            Workbook workbook = new Workbook(csvPath, loadOptions);

            // Configure JSON save options (optional settings can be adjusted as needed)
            JsonSaveOptions jsonOptions = new JsonSaveOptions
            {
                // Export the workbook as a JSON object even if there is only one worksheet
                AlwaysExportAsJsonObject = true,
                // Include empty cells in the JSON output
                ExportEmptyCells = true,
                // Preserve the Excel structure in the JSON (useful for hierarchical data)
                ToExcelStruct = true
            };

            // Save the workbook as a JSON file using the configured options
            workbook.Save(jsonPath, jsonOptions);

            Console.WriteLine($"CSV file '{csvPath}' has been successfully converted to JSON at '{jsonPath}'.");
        }
    }
}