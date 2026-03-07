using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace CsvToJsonExample
{
    class Program
    {
        static void Main()
        {
            // Paths for source CSV and destination JSON files
            string csvPath = "input.csv";
            string jsonPath = "output.json";

            // Load the CSV file into a workbook using CSV load options
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Csv);
            Workbook workbook = new Workbook(csvPath, loadOptions);

            // Configure JSON save options
            JsonSaveOptions jsonOptions = new JsonSaveOptions
            {
                // Export the workbook as a JSON object even if there is only one worksheet
                AlwaysExportAsJsonObject = true,
                // Include header row information if present
                HasHeaderRow = true,
                // Export empty cells as null values
                ExportEmptyCells = true,
                // Export cell values as strings (optional, can be omitted)
                ExportAsString = true
            };

            // Save the workbook as a JSON file
            workbook.Save(jsonPath, jsonOptions);

            Console.WriteLine($"CSV file \"{csvPath}\" has been converted to JSON file \"{jsonPath}\".");
        }
    }
}