using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace SpreadsheetToJson
{
    class Program
    {
        static void Main(string[] args)
        {
            // Input spreadsheet path (XLSX or CSV)
            string inputPath = "input.xlsx"; // change to your file path

            // Output JSON file path
            string outputJsonPath = "output.json";

            // Load the workbook (Aspose.Cells automatically detects format based on extension)
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (you can modify to select a different sheet)
            Worksheet worksheet = workbook.Worksheets[0];

            // Determine the used range of the worksheet
            int lastRow = worksheet.Cells.MaxDataRow;      // zero‑based index of last row with data
            int lastColumn = worksheet.Cells.MaxDataColumn; // zero‑based index of last column with data

            // Create a range that covers all used cells (add 1 because CreateRange expects count, not index)
            Aspose.Cells.Range usedRange = worksheet.Cells.CreateRange(0, 0, lastRow + 1, lastColumn + 1);

            // Configure JSON export options
            JsonSaveOptions jsonOptions = new JsonSaveOptions
            {
                ExportEmptyCells = true,   // include empty cells in the output
                HasHeaderRow = true,       // treat the first row as header names
                ToExcelStruct = false      // export as simple array of objects
            };

            // Export the range to a JSON string
            string jsonResult = JsonUtility.ExportRangeToJson(usedRange, jsonOptions);

            // Write the JSON string to a file
            File.WriteAllText(outputJsonPath, jsonResult);

            Console.WriteLine($"Spreadsheet data has been exported to JSON file: {outputJsonPath}");
        }
    }
}