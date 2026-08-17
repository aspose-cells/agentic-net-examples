// Title: Export Excel Range to JSON with Numbers as Formatted Strings – Aspose.Cells for .NET
// Description: Demonstrates creating a workbook, applying a custom number format to a price column, enabling ExportAsString in JsonSaveOptions, and using JsonUtility.ExportRangeToJson to produce indented JSON where numeric values are emitted as formatted string literals.
// Keywords: Aspose.Cells | C# | .NET | JSON export | ExportAsString | numeric to string | custom number format | pretty‑print JSON | Excel to JSON conversion | price column formatting
// Common Searches: Aspose.Cells export numeric cells as strings JSON | JsonSaveOptions ExportAsString example C# | convert Excel numbers to JSON string values | pretty printed JSON from Excel range Aspose | custom number format JSON export Aspose.Cells
// Developer Intent: Generate a JSON representation of a worksheet range where all numeric cells are output as formatted string values.
// Use Cases: Send pricing data to an API that expects currency values as strings with thousand separators. | Create configuration files from Excel where downstream parsers require string‑only values. | Produce human‑readable JSON reports from spreadsheets while preserving numeric formatting.
// AI Prompts: Write C# code with Aspose.Cells that exports a selected range to JSON, forcing numbers to appear as formatted strings. | Explain how JsonSaveOptions.ExportAsString changes the JSON output for numeric cells in Aspose.Cells. | Show an example that applies a custom number format to a column and exports the range to indented JSON.

using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

// Demonstrates creating a workbook, applying a custom number format to a price column, enabling ExportAsString in JsonSaveOptions, and using JsonUtility.ExportRangeToJson to produce indented JSON where numeric values are emitted as formatted string literals.
class JsonNumericToStringDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Add header and numeric data
            cells["A1"].PutValue("Product");
            cells["B1"].PutValue("Price");
            cells["A2"].PutValue("Laptop");
            cells["B2"].PutValue(999.99);
            cells["A3"].PutValue("Phone");
            cells["B3"].PutValue(599.99);

            // Apply a custom number format to the price column
            Style priceStyle = workbook.CreateStyle();
            priceStyle.Custom = "#,##0.00"; // two decimal places with thousand separator
            cells["B2:B3"].SetStyle(priceStyle);

            // Configure JSON export options to export cell values as strings
            JsonSaveOptions jsonOptions = new JsonSaveOptions
            {
                ExportAsString = true,      // forces numeric values to be output as formatted strings
                Indent = "    "             // optional pretty‑print indentation
            };

            // Export the defined range to a JSON string
            Aspose.Cells.Range exportRange = cells.CreateRange("A1:B3");
            string jsonOutput = JsonUtility.ExportRangeToJson(exportRange, jsonOptions);

            // Display the resulting JSON
            Console.WriteLine(jsonOutput);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
