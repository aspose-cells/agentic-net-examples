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
            // Path to the source CSV file (SXC CSV)
            string csvPath = "input.csv";

            // Path for the resulting JSON file
            string jsonPath = "output.json";

            // Create a new empty workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet and its cells collection
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Import the CSV data starting at cell A1 (row 0, column 0)
            // Using comma as delimiter and converting numeric data where possible
            cells.ImportCSV(csvPath, ",", true, 0, 0);

            // Determine the used range dimensions
            int totalRows = cells.MaxDataRow + 1;      // MaxDataRow is zero‑based
            int totalColumns = cells.MaxDataColumn + 1;

            // Create a range that covers all imported data
            Aspose.Cells.Range dataRange = cells.CreateRange(0, 0, totalRows, totalColumns);

            // Configure JSON export options
            JsonSaveOptions jsonOptions = new JsonSaveOptions
            {
                ExportEmptyCells = true,   // keep empty cells in the output
                HasHeaderRow = true,       // first row contains column names
                ExportNestedStructure = false
            };

            // Export the range to a JSON string
            string jsonResult = JsonUtility.ExportRangeToJson(dataRange, jsonOptions);

            // Write the JSON string to the output file
            File.WriteAllText(jsonPath, jsonResult);

            Console.WriteLine($"CSV data from '{csvPath}' has been converted to JSON and saved to '{jsonPath}'.");
        }
    }
}