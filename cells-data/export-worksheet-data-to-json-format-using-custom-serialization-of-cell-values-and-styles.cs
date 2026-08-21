// Title: Export Excel Range to JSON with Custom Value and Style Serialization using Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook, populate a product‑price table, apply header and numeric formatting, and then export the A1:B4 range to a pretty‑printed JSON string. The example uses JsonSaveOptions to serialize all cell values as strings, include empty cells as null, write styles per cell, treat the first row as a header, and control indentation.
// Keywords: Aspose.Cells | C# | .NET | JsonSaveOptions | export range to JSON | Excel to JSON | custom JSON serialization | include empty cells | export cell styles | pretty printed JSON | header row handling | range export
// Common Searches: Aspose.Cells export range to JSON C# | how to include empty cells in JSON export with Aspose.Cells | export Excel sheet as JSON with custom style options | pretty printed JSON from Excel using Aspose.Cells | JsonSaveOptions ExportAsString example
// Developer Intent: Generate a JSON representation of a selected worksheet range with custom serialization of values and styles.
// Use Cases: Create a JSON API payload from an Excel product list while preserving column headers and representing missing prices as null. | Produce a readable, indented JSON file for data exchange with systems that require all values as strings. | Log worksheet content in JSON format for debugging, auditing, or version control.
// AI Prompts: Show how to modify JsonSaveOptions to use a shared style pool instead of per‑cell style export. | Provide code to deserialize the generated JSON back into a DataTable or a list of C# objects. | Explain how to export a nested JSON structure from multiple worksheets using Aspose.Cells.

using System;
using System.IO;
using System.Drawing;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// Demonstrates how to create a workbook, populate a product‑price table, apply header and numeric formatting, and then export the A1:B4 range to a pretty‑printed JSON string. The example uses JsonSaveOptions to serialize all cell values as strings, include empty cells as null, write styles per cell, treat the first row as a header, and control indentation.
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

            // Populate sample data
            cells["A1"].PutValue("Product");
            cells["B1"].PutValue("Price");
            cells["A2"].PutValue("Apple");
            cells["B2"].PutValue(1.5);
            cells["A3"].PutValue("Banana");
            cells["B3"].PutValue(0.8);
            // Leave B4 empty intentionally to test empty‑cell handling
            cells["A4"].PutValue("Cherry");

            // Apply distinct styles to header cells
            Style headerStyle = workbook.CreateStyle();
            headerStyle.Font.IsBold = true;
            headerStyle.Font.Color = Color.White;
            headerStyle.ForegroundColor = Color.DarkBlue;
            headerStyle.Pattern = BackgroundType.Solid;
            cells["A1"].SetStyle(headerStyle);
            cells["B1"].SetStyle(headerStyle);

            // Apply numeric format to price cells
            Style priceStyle = workbook.CreateStyle();
            priceStyle.Number = 2; // two decimal places
            cells["B2"].SetStyle(priceStyle);
            cells["B3"].SetStyle(priceStyle);

            // Configure JSON export options with custom serialization settings
            JsonSaveOptions jsonOptions = new JsonSaveOptions
            {
                ExportAsString = true,          // export all values as strings
                ExportEmptyCells = true,        // include empty cells as null
                ExportStylePool = false,        // export style for each cell individually
                HasHeaderRow = true,            // first row is treated as header
                Indent = "    ",                // pretty‑print with 4‑space indentation
                ExportNestedStructure = false  // flat JSON structure
            };

            // Define the range to export (A1:B4)
            AsposeRange exportRange = sheet.Cells.CreateRange("A1:B4");

            // Convert the range to JSON using the configured options
            string jsonResult = exportRange.ToJson(jsonOptions);

            // Output the JSON string to the console
            Console.WriteLine(jsonResult);

            // Optionally write the JSON to a file
            string outputPath = "ExportedData.json";
            File.WriteAllText(outputPath, jsonResult);
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
