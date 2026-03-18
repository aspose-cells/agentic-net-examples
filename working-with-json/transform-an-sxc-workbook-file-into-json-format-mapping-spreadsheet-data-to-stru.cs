using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsSxcToJson
{
    class Program
    {
        static void Main()
        {
            // Path to the source SXC file
            string sxcPath = "input.sxc";

            // Load the workbook
            Workbook workbook = new Workbook(sxcPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Determine the used range dimensions
            int maxRow = sheet.Cells.MaxDataRow;
            int maxColumn = sheet.Cells.MaxDataColumn;

            if (maxRow < 0 || maxColumn < 0)
            {
                Console.WriteLine("The worksheet is empty. No data to export.");
                return;
            }

            int firstRow = 0;
            int firstColumn = 0;
            int totalRows = maxRow + 1;      // count of rows
            int totalColumns = maxColumn + 1; // count of columns

            // Create a range that covers all used cells
            AsposeRange usedRange = sheet.Cells.CreateRange(firstRow, firstColumn, totalRows, totalColumns);

            // Configure JSON export options
            JsonSaveOptions jsonOptions = new JsonSaveOptions
            {
                ToExcelStruct = true,
                HasHeaderRow = true,
                ExportEmptyCells = true
            };

            // Export the range to a JSON string
            string jsonResult = JsonUtility.ExportRangeToJson(usedRange, jsonOptions);

            // Output the JSON to console
            Console.WriteLine("Exported JSON:");
            Console.WriteLine(jsonResult);

            // Write the JSON to a file
            string jsonPath = "output.json";
            File.WriteAllText(jsonPath, jsonResult);
            Console.WriteLine($"JSON saved to '{jsonPath}'");
        }
    }
}