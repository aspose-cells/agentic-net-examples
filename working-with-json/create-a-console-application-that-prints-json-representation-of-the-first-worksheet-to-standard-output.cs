// Title: Export First Worksheet to JSON and Print to Console with Aspose.Cells for .NET
// Description: A C# console app that creates a workbook, fills the first worksheet with sample data, determines the used range, configures JsonSaveOptions (header row enabled), converts the range to a JSON string via JsonUtility.ExportRangeToJson, and writes the JSON to standard output while handling errors.
// Keywords: Aspose.Cells | C# | .NET | Export worksheet to JSON | JsonUtility | JsonSaveOptions | HasHeaderRow | Console JSON output | Excel range to JSON | Aspose.Cells example
// Common Searches: convert Aspose.Cells worksheet to JSON C# | Aspose.Cells export range as JSON string | print Excel data as JSON in a console app | JsonUtility ExportRangeToJson sample code | Aspose.Cells JsonSaveOptions usage
// Developer Intent: Generate a JSON representation of the first worksheet and display it in the console.
// Use Cases: Stream worksheet data as JSON to a web service without creating a file. | Log Excel content in JSON format for debugging or audit purposes. | Provide on‑the‑fly JSON data to front‑end applications from a lightweight utility.
// AI Prompts: How can I limit the export to columns A‑C only? | Show me code that writes the JSON output to a file instead of the console. | Explain how to preserve data types and handle empty cells when using JsonUtility.ExportRangeToJson.

using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

// A C# console app that creates a workbook, fills the first worksheet with sample data, determines the used range, configures JsonSaveOptions (header row enabled), converts the range to a JSON string via JsonUtility.ExportRangeToJson, and writes the JSON to standard output while handling errors.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data in the first worksheet
            worksheet.Cells["A1"].PutValue("Name");
            worksheet.Cells["B1"].PutValue("Age");
            worksheet.Cells["A2"].PutValue("John");
            worksheet.Cells["B2"].PutValue(30);
            worksheet.Cells["A3"].PutValue("Alice");
            worksheet.Cells["B3"].PutValue(25);

            // Determine the used range of the worksheet
            int lastRow = worksheet.Cells.MaxDataRow;
            int lastColumn = worksheet.Cells.MaxDataColumn;
            // Use Aspose.Cells.Range explicitly to avoid conflict with System.Range
            Aspose.Cells.Range usedRange = worksheet.Cells.CreateRange(0, 0, lastRow + 1, lastColumn + 1);

            // Configure JSON export options
            JsonSaveOptions jsonOptions = new JsonSaveOptions
            {
                HasHeaderRow = true // treat the first row as header
            };

            // Export the range to a JSON string
            string jsonOutput = JsonUtility.ExportRangeToJson(usedRange, jsonOptions);

            // Print the JSON representation to standard output
            Console.WriteLine(jsonOutput);
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}
