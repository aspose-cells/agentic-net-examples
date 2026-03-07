using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace CsvToJsonConversion
{
    class Program
    {
        static void Main()
        {
            // Paths for input CSV and output JSON
            string csvPath = "input.csv";
            string jsonPath = "output.json";

            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Import CSV data starting at cell A1 (row 0, column 0)
            // Using comma as delimiter and converting numeric data automatically
            cells.ImportCSV(csvPath, ",", true, 0, 0);

            // Determine the range that contains the imported data
            Aspose.Cells.Range dataRange = cells.MaxDisplayRange;

            // Configure JSON export options (customize as needed)
            JsonSaveOptions jsonOptions = new JsonSaveOptions
            {
                ExportEmptyCells = true,   // include empty cells in the output
                HasHeaderRow = true,       // treat the first row as header
                ExportNestedStructure = false
            };

            // Export the range to a JSON string
            string jsonOutput = JsonUtility.ExportRangeToJson(dataRange, jsonOptions);

            // Write the JSON string to the output file
            File.WriteAllText(jsonPath, jsonOutput);

            Console.WriteLine($"CSV file '{csvPath}' has been converted to JSON and saved as '{jsonPath}'.");
        }
    }
}