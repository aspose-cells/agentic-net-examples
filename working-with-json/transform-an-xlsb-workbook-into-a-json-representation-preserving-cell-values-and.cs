using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

class XlsbToJsonConverter
{
    static void Main()
    {
        // Path to the source XLSB workbook
        string inputPath = "input.xlsb";

        // Directory where JSON files will be saved
        string outputDir = "JsonOutput";
        Directory.CreateDirectory(outputDir);

        // Load the XLSB workbook
        Workbook workbook = new Workbook(inputPath);

        // Iterate through each worksheet in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Determine the used range of the worksheet
            int maxRow = sheet.Cells.MaxDataRow;      // zero‑based index of the last row with data
            int maxCol = sheet.Cells.MaxDataColumn;   // zero‑based index of the last column with data

            // If the sheet is empty, skip it
            if (maxRow < 0 || maxCol < 0)
                continue;

            // Create a range that covers all used cells
            Aspose.Cells.Range usedRange = sheet.Cells.CreateRange(0, 0, maxRow + 1, maxCol + 1);

            // Configure JSON export options to preserve Excel structure
            JsonSaveOptions jsonOptions = new JsonSaveOptions
            {
                ToExcelStruct = true,          // preserve hierarchical structure
                HasHeaderRow = true,           // treat first row as header (optional)
                ExportEmptyCells = true        // include empty cells in the output
            };

            // Export the range to a JSON string
            string json = JsonUtility.ExportRangeToJson(usedRange, jsonOptions);

            // Save the JSON string to a file named after the worksheet
            string jsonPath = Path.Combine(outputDir, $"{sheet.Name}.json");
            File.WriteAllText(jsonPath, json);

            Console.WriteLine($"Worksheet '{sheet.Name}' exported to JSON at: {jsonPath}");
        }

        // Clean up
        workbook.Dispose();
    }
}