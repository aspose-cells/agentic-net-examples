using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsCsvToJson
{
    class Program
    {
        static void Main()
        {
            // Paths for the source CSV file and the destination JSON file
            string csvPath = "sample.csv";
            string jsonPath = "result.json";

            // Create a sample CSV file (optional, for demonstration)
            File.WriteAllText(csvPath,
                "Name,Age,City\n" +
                "John,30,New York\n" +
                "Alice,25,London\n" +
                "Bob,35,Paris");

            // Load the CSV file into a Workbook using LoadOptions for CSV format
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Csv);
            Workbook workbook = new Workbook(csvPath, loadOptions);

            // Configure JSON save options
            JsonSaveOptions jsonOptions = new JsonSaveOptions
            {
                // Export the workbook as a JSON object even if it contains a single worksheet
                AlwaysExportAsJsonObject = true,
                // Treat the first row as header
                HasHeaderRow = true,
                // Export empty cells as null (optional)
                ExportEmptyCells = true,
                // Export cell values as strings (optional)
                ExportAsString = true,
                // Indent the JSON output for readability
                Indent = "  "
            };

            // Save the workbook as a JSON file
            workbook.Save(jsonPath, jsonOptions);

            // Output the generated JSON to console for verification
            Console.WriteLine("JSON conversion completed. Content:");
            Console.WriteLine(File.ReadAllText(jsonPath));
        }
    }
}