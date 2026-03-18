using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsDIFCsvToJson
{
    class Program
    {
        static void Main()
        {
            // Path to the DIF‑formatted CSV file
            string csvPath = "data.csv";

            // Path where the resulting JSON will be saved
            string jsonOutputPath = "data.json";

            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Import the CSV data.
            // Parameters: file name, delimiter, convert numeric data, start row, start column
            // convertNumericData = true ensures numbers and dates are stored as native types.
            cells.ImportCSV(csvPath, ",", true, 0, 0); // lifecycle: load

            // Determine the used range (including header row)
            int totalRows = cells.MaxDataRow + 1;      // MaxDataRow is zero‑based
            int totalCols = cells.MaxDataColumn + 1;   // MaxDataColumn is zero‑based
            Aspose.Cells.Range usedRange = cells.CreateRange(0, 0, totalRows, totalCols);

            // Configure JSON export options.
            JsonSaveOptions jsonOptions = new JsonSaveOptions
            {
                ExportEmptyCells = true,   // keep empty cells in output
                HasHeaderRow = true,       // first row contains column names
                ExportNestedStructure = false
            };

            // Export the range to a JSON string (preserves data types)
            string json = JsonUtility.ExportRangeToJson(usedRange, jsonOptions); // lifecycle: save (export)

            // Write the JSON string to a file
            File.WriteAllText(jsonOutputPath, json);

            // Optionally, save the workbook for verification
            workbook.Save("intermediate.xlsx", SaveFormat.Xlsx);
        }
    }
}