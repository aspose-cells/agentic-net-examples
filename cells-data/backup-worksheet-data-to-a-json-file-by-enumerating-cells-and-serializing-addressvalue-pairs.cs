// Title: Backup Excel Worksheet to JSON with Aspose.Cells for .NET
// Description: Shows how to capture a worksheet’s used range, set JsonSaveOptions, convert the range to a formatted JSON string, and write the result to a file using Aspose.Cells in C#.
// Keywords: Aspose.Cells JSON export | C# export worksheet to JSON | JsonSaveOptions | Export used range to JSON | Aspose.Cells backup Excel | JsonUtility ExportRangeToJson | pretty printed JSON Excel | save Excel as JSON .NET
// Common Searches: Aspose.Cells export range to JSON C# | How to save Excel worksheet as JSON using Aspose.Cells | JsonSaveOptions example C# | Backup Excel data to JSON file .NET | Export used range to JSON Aspose.Cells
// Developer Intent: Create a JSON backup of the data contained in an Excel worksheet using Aspose.Cells.
// Use Cases: Store spreadsheet contents in version‑controlled JSON files for archival or CI pipelines. | Transmit worksheet data to REST APIs that accept JSON payloads. | Produce human‑readable reports from Excel tables while preserving header rows and empty‑cell information.
// AI Prompts: Generate C# code that uses Aspose.Cells to export a worksheet’s used range to a pretty‑printed JSON file, including empty cells as null. | Explain how to configure JsonSaveOptions to retain header rows, export empty cells, and apply custom indentation when converting an Excel range to JSON. | Show how to deserialize the JSON file created by JsonUtility.ExportRangeToJson back into a DataTable or a list of strongly‑typed objects.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;
using AsposeRange = Aspose.Cells.Range;

// Shows how to capture a worksheet’s used range, set JsonSaveOptions, convert the range to a formatted JSON string, and write the result to a file using Aspose.Cells in C#.
class WorksheetBackupToJson
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Sample data – in real scenario the worksheet would already contain data
            cells["A1"].PutValue("Name");
            cells["B1"].PutValue("Age");
            cells["A2"].PutValue("John");
            cells["B2"].PutValue(30);
            cells["A3"].PutValue("Alice");
            cells["B3"].PutValue(25);

            // Determine the used range (rows and columns)
            int maxRow = cells.MaxDataRow;      // zero‑based index of last row with data
            int maxCol = cells.MaxDataColumn;   // zero‑based index of last column with data

            // Create a range that covers the entire used area
            AsposeRange usedRange = cells.CreateRange(0, 0, maxRow + 1, maxCol + 1);

            // Configure JSON export options
            JsonSaveOptions jsonOptions = new JsonSaveOptions
            {
                ExportEmptyCells = true,   // Export empty cells as null (optional)
                HasHeaderRow = true,       // Include header row if present (optional)
                ExportAsString = false,    // Export values as strings (optional)
                Indent = "  "               // Indent for readability
            };

            // Export the range to a JSON string
            string json = JsonUtility.ExportRangeToJson(usedRange, jsonOptions);

            // Define output file path
            string outputPath = "WorksheetBackup.json";

            // Write JSON string to file
            File.WriteAllText(outputPath, json);

            Console.WriteLine($"Worksheet data has been backed up to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
