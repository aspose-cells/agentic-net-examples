using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace CsvToJsonConversion
{
    class Program
    {
        static void Main()
        {
            // Path to the source CSV file (must exist)
            string csvFilePath = "input.csv";

            // Load the CSV file into a Workbook using LoadOptions for CSV format
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Csv);
            Workbook workbook = new Workbook(csvFilePath, loadOptions);

            // Configure JSON save options (customize as needed)
            JsonSaveOptions jsonOptions = new JsonSaveOptions
            {
                // Treat the first row as header (optional)
                HasHeaderRow = true,
                // Export empty cells as null (optional)
                ExportEmptyCells = true,
                // Indent the JSON output for readability (optional)
                Indent = "  "
            };

            // Path for the resulting JSON file
            string jsonFilePath = "output.json";

            // Save the workbook as a JSON file using the configured options
            workbook.Save(jsonFilePath, jsonOptions);

            Console.WriteLine($"CSV file '{csvFilePath}' has been converted to JSON file '{jsonFilePath}'.");
        }
    }
}