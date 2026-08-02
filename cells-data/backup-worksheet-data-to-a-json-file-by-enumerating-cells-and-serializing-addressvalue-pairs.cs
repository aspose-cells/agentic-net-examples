// Title: Backup an Excel worksheet to JSON with Aspose.Cells for .NET (C#)
// Description: Loads a workbook, extracts the first worksheet’s used range, sets JsonSaveOptions to include empty cells, export all values as strings, and pretty‑print with two‑space indentation, then writes the JSON to a file.
// Keywords: Aspose.Cells | C# | Export to JSON | Worksheet backup | JsonSaveOptions | Include empty cells | Pretty print JSON | Excel to JSON | JsonUtility.ExportRangeToJson | Aspose.Cells example
// Common Searches: How to export an Excel worksheet to JSON using Aspose.Cells C# | Aspose.Cells include empty cells in JSON export | Set indentation for JSON output with Aspose.Cells | Backup Excel data to JSON file Aspose.Cells | Export used range to JSON Aspose.Cells .NET
// Developer Intent: Generate a JSON file that contains the full contents of a worksheet’s used range.
// Use Cases: Create a version‑controlled backup of spreadsheet data for disaster recovery. | Provide a JSON snapshot of Excel data for consumption by web APIs or JavaScript front‑ends. | Preserve cell coordinates by exporting empty cells as null while keeping the output human‑readable.
// AI Prompts: Write C# code that uses Aspose.Cells to export a worksheet’s used range to a formatted JSON file, including empty cells as null and using two‑space indentation. | Show how to check for a missing input workbook, log an error, and safely abort before attempting JSON export with Aspose.Cells. | Modify the example to export a named range instead of the entire used range while keeping the same JsonSaveOptions settings.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

// Alias to avoid conflict with System.Range (C# 8+)
using AsposeRange = Aspose.Cells.Range;

// Loads a workbook, extracts the first worksheet’s used range, sets JsonSaveOptions to include empty cells, export all values as strings, and pretty‑print with two‑space indentation, then writes the JSON to a file.
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
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Determine the used range (from A1 to the last non‑empty cell)
            int lastRow = worksheet.Cells.MaxDataRow;
            int lastColumn = worksheet.Cells.MaxDataColumn;
            AsposeRange usedRange = worksheet.Cells.CreateRange(0, 0, lastRow + 1, lastColumn + 1);

            // Set JSON export options
            JsonSaveOptions jsonOptions = new JsonSaveOptions
            {
                ExportEmptyCells = true,   // include empty cells as null
                ExportAsString = true,     // export all values as strings
                HasHeaderRow = false,      // no header row assumed
                Indent = "  "               // pretty‑print with two spaces
            };

            // Export the range to a JSON string
            string json = JsonUtility.ExportRangeToJson(usedRange, jsonOptions);

            // Write the JSON string to a file
            File.WriteAllText(outputPath, json);

            Console.WriteLine($"Worksheet data has been backed up to \"{outputPath}\"");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
