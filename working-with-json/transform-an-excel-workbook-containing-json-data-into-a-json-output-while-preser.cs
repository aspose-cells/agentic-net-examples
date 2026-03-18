using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsJsonTransform
{
    class Program
    {
        static void Main()
        {
            // Path to the source Excel workbook that already contains JSON data
            string excelPath = "input.xlsx";

            // Load the workbook
            Workbook workbook = new Workbook(excelPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Determine the used range (rows and columns with data)
            int maxRow = sheet.Cells.MaxDataRow;          // zero‑based index of last row with data
            int maxColumn = sheet.Cells.MaxDataColumn;    // zero‑based index of last column with data

            // Create a range that covers all used cells
            Aspose.Cells.Range usedRange = sheet.Cells.CreateRange(0, 0, maxRow + 1, maxColumn + 1);

            // Configure JSON save options to preserve the Excel structure
            JsonSaveOptions jsonOptions = new JsonSaveOptions
            {
                ToExcelStruct = true,
                ExportNestedStructure = true,
                HasHeaderRow = true,
                ExportEmptyCells = true,
                ExportAsString = false,
                Indent = "  "
            };

            // Export the selected range to a JSON string
            string jsonResult = JsonUtility.ExportRangeToJson(usedRange, jsonOptions);

            // Output the JSON to console
            Console.WriteLine("Exported JSON:");
            Console.WriteLine(jsonResult);

            // Save the JSON string to a file
            string jsonPath = "output.json";
            File.WriteAllText(jsonPath, jsonResult);
            Console.WriteLine($"JSON saved to: {jsonPath}");
        }
    }
}