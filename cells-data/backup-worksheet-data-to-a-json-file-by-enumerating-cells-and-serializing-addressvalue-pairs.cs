// Title: Backup an Excel worksheet to a JSON file by enumerating cells with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that opens an .xlsx file with Aspose.Cells, identifies the worksheet's used area, and saves each cell's address and value to a JSON file, ensuring empty cells appear as null. | Implement a C# helper method that receives a worksheet index and an output path, sets JsonSaveOptions to keep original data types, and uses Aspose.Cells utilities to convert the selected range into JSON. | Modify the export routine so that the first row is treated as a header row, turning its values into property names for the subsequent rows in the generated JSON document.
// Common Searches: Aspose.Cells C# export used range of worksheet to JSON file | include empty cells when converting Excel to JSON with Aspose.Cells | C# example for backing up Excel worksheet data as JSON using JsonSaveOptions | how to serialize cell address and value pairs from Excel to JSON in .NET | Aspose.Cells JsonSaveOptions ExportEmptyCells true example
// Tags: Aspose.Cells export worksheet to JSON | JsonSaveOptions include empty cells | C# backup Excel worksheet as JSON | serialize used cell range Aspose.Cells | export address-value pairs to JSON

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

// The example loads an .xlsx workbook, selects the first worksheet's used range, configures JsonSaveOptions to include empty cells and preserve data types, exports the range to a JSON string via Aspose.Cells, and writes the result to a backup file named worksheet_backup.json.
class BackupWorksheetToJson
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "worksheet_backup.json";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (or any specific worksheet)
            Worksheet worksheet = workbook.Worksheets[0];

            // Determine the used range of the worksheet
            int maxRow = worksheet.Cells.MaxDataRow;          // zero‑based index of last used row
            int maxColumn = worksheet.Cells.MaxDataColumn;    // zero‑based index of last used column

            // Create a range that covers all used cells
            // Add 1 because CreateRange expects count, not last index
            Aspose.Cells.Range usedRange = worksheet.Cells.CreateRange(0, 0, maxRow + 1, maxColumn + 1);

            // Configure JSON export options
            JsonSaveOptions jsonOptions = new JsonSaveOptions
            {
                ExportEmptyCells = true,   // include empty cells as null
                HasHeaderRow = false,      // treat first row as data, not header
                ExportAsString = false     // keep original data types
            };

            // Export the range to a JSON string
            string json = JsonUtility.ExportRangeToJson(usedRange, jsonOptions);

            // Write the JSON string to a file (backup)
            File.WriteAllText(outputPath, json);
            Console.WriteLine($"Worksheet backup saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
