using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;
using AsposeRange = Aspose.Cells.Range;

namespace OdsToJsonExample
{
    class Program
    {
        static void Main()
        {
            // Path to the source ODS file
            string sourcePath = "input.ods";

            // Load the ODS file with default load options
            OdsLoadOptions loadOptions = new OdsLoadOptions();
            Workbook workbook = new Workbook(sourcePath, loadOptions);

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Determine the used range (including all populated rows and columns)
            int lastRow = cells.MaxDataRow;          // zero‑based index of the last row with data
            int lastColumn = cells.MaxDataColumn;    // zero‑based index of the last column with data

            // Create a range that covers the entire used area
            AsposeRange range = cells.CreateRange(0, 0, lastRow + 1, lastColumn + 1);

            // Configure JSON export options
            JsonSaveOptions jsonOptions = new JsonSaveOptions
            {
                HasHeaderRow = true,          // treat the first row as header
                ExportEmptyCells = true,      // include empty cells in the output
                ToExcelStruct = false         // export as simple JSON array/object
            };

            // Export the range to a JSON string
            string json = JsonUtility.ExportRangeToJson(range, jsonOptions);

            // Write JSON to a file
            string outputPath = "output.json";
            File.WriteAllText(outputPath, json);

            // Output the JSON to console for verification
            Console.WriteLine("JSON export completed. Content:");
            Console.WriteLine(json);
        }
    }
}