using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace CsvToJsonConverter
{
    class Program
    {
        static void Main(string[] args)
        {
            // Input CSV file path (SpreadsheetML CSV)
            string csvPath = "input.csv";

            // Output JSON file path
            string jsonPath = "output.json";

            // 1. Create a new workbook
            Workbook workbook = new Workbook();

            // 2. Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // 3. Import CSV data into the worksheet.
            //    - Use comma as delimiter.
            //    - Enable numeric conversion to preserve data types.
            //    - Start import at cell A1 (row 0, column 0).
            cells.ImportCSV(csvPath, ",", true, 0, 0);

            // 4. Determine the used range that contains the imported data.
            int totalRows = cells.MaxDataRow + 1;
            int totalColumns = cells.MaxDataColumn + 1;
            Aspose.Cells.Range dataRange = cells.CreateRange(0, 0, totalRows, totalColumns);

            // 5. Configure JSON export options to keep headers and empty cells.
            JsonSaveOptions jsonOptions = new JsonSaveOptions
            {
                HasHeaderRow = true,          // Preserve column headers
                ExportEmptyCells = true,      // Include empty cells in the output
                ExportNestedStructure = false // Flat structure matching the worksheet layout
            };

            // 6. Export the range to a JSON string.
            string jsonResult = JsonUtility.ExportRangeToJson(dataRange, jsonOptions);

            // 7. Write the JSON string to the output file.
            File.WriteAllText(jsonPath, jsonResult);

            // Optional: display a confirmation message.
            Console.WriteLine($"CSV data from '{csvPath}' has been converted to JSON and saved to '{jsonPath}'.");
        }
    }
}