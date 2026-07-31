// Title: Export Excel worksheet range to JSON with custom value and style serialization using Aspose.Cells for .NET
// Description: Creates a workbook, fills a product table, applies header and price styles, defines range A1:C3, and uses JsonSaveOptions (ExportAsString, ExportEmptyCells, ExportStylePool, HasHeaderRow, Indent) to generate a pretty‑printed JSON string via JsonUtility.ExportRangeToJson. The JSON is displayed on the console and saved to a file.
// Keywords: Aspose.Cells export to JSON | JsonSaveOptions C# | ExportAsString Aspose | Include empty cells JSON | Excel cell style serialization | ExportRangeToJson example | .NET Excel to JSON
// Common Searches: Aspose.Cells export worksheet to JSON with header row | How to include empty cells as null in JSON export using Aspose.Cells | JsonSaveOptions ExportAsString C# example | Serialize Excel cell styles to JSON with Aspose.Cells | Pretty print JSON from Excel range Aspose
// Developer Intent: Generate a JSON representation of a selected worksheet range while controlling value types, empty‑cell handling, and per‑cell style output.
// Use Cases: Convert a product catalog stored in Excel into a JSON API payload that preserves column names and forces all values to strings. | Create a JSON file for front‑end applications where missing data must appear as null entries. | Export styled spreadsheet data for a reporting engine that reads cell formatting from JSON.
// AI Prompts: Write C# code that uses Aspose.Cells to export a worksheet range to JSON with ExportAsString enabled and empty cells serialized as null. | Show how to modify JsonSaveOptions to include the shared style pool for each cell in the JSON output. | Provide a method that reads the JSON file produced by ExportRangeToJson and rebuilds the original worksheet with its styles using Aspose.Cells.

using System;
using System.IO;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Utility;

// Creates a workbook, fills a product table, applies header and price styles, defines range A1:C3, and uses JsonSaveOptions (ExportAsString, ExportEmptyCells, ExportStylePool, HasHeaderRow, Indent) to generate a pretty‑printed JSON string via JsonUtility.ExportRangeToJson. The JSON is displayed on the console and saved to a file.
class ExportWorksheetToJson
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data with a header row
            cells["A1"].PutValue("Product");
            cells["B1"].PutValue("Price");
            cells["C1"].PutValue("InStock");
            cells["A2"].PutValue("Apple");
            cells["B2"].PutValue(1.2);
            cells["C2"].PutValue(true);
            cells["A3"].PutValue("Banana");
            cells["B3"].PutValue(0.8);
            cells["C3"].PutValue(false);
            // Intentionally leave some cells empty to test ExportEmptyCells

            // Apply distinct styles to demonstrate custom style serialization
            Style headerStyle = workbook.CreateStyle();
            headerStyle.Font.IsBold = true;
            headerStyle.ForegroundColor = Color.LightGray;
            headerStyle.Pattern = BackgroundType.Solid;
            cells["A1"].SetStyle(headerStyle);
            cells["B1"].SetStyle(headerStyle);
            cells["C1"].SetStyle(headerStyle);

            Style priceStyle = workbook.CreateStyle();
            priceStyle.Number = 2; // two decimal places
            cells["B2"].SetStyle(priceStyle);
            cells["B3"].SetStyle(priceStyle);

            // Define the range to export (including the header row)
            Aspose.Cells.Range exportRange = cells.CreateRange("A1:C3");

            // Configure JSON export options for custom serialization
            JsonSaveOptions jsonOptions = new JsonSaveOptions
            {
                ExportAsString = true,      // serialize all cell values as strings
                ExportEmptyCells = true,    // include empty cells as null in JSON
                ExportStylePool = false,    // export style for each cell individually
                HasHeaderRow = true,        // first row is treated as header
                Indent = "    "             // pretty‑print JSON with 4‑space indentation
            };

            // Export the defined range to a JSON string using the configured options
            string json = JsonUtility.ExportRangeToJson(exportRange, jsonOptions);

            // Output the JSON string to the console
            Console.WriteLine(json);

            // Optionally write the JSON to a file
            string outputPath = "WorksheetExport.json";
            File.WriteAllText(outputPath, json);
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
