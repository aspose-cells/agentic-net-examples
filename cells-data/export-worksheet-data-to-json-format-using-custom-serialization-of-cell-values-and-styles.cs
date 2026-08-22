// Title: Export an Aspose.Cells worksheet to a formatted JSON file with custom value conversion and per‑cell style information in C#
// AI Prompts: Generate C# code that uses Aspose.Cells JsonSaveOptions to save an entire worksheet as an indented JSON document, converting every cell value to a string and embedding each cell's style data. | Write C# that extracts a specific range from a worksheet and returns its JSON representation using JsonUtility.ExportRangeToJson with options for empty cells, string conversion, and Excel structure. | Show how to adjust JsonSaveOptions to produce a compact JSON output that omits per‑cell style information while preserving header rows.
// Common Searches: how to use Aspose.Cells JsonSaveOptions to export a worksheet as JSON with cell styles in C# | C# Aspose.Cells export selected range to JSON string including formatting | save Excel workbook to a pretty printed JSON file using Aspose.Cells .NET | Aspose.Cells export empty cells as null and values as strings in JSON | customize JSON output from Aspose.Cells to include Excel structure
// Tags: Aspose.Cells JsonSaveOptions export worksheet to JSON | C# export Excel range to JSON with style data | serialize cell values as strings using Aspose.Cells | include per‑cell formatting in JSON output Aspose | pretty‑printed JSON export from workbook Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;
using AsposeRange = Aspose.Cells.Range;

// The example creates a workbook, adds sample data, applies header and numeric styles, configures JsonSaveOptions to serialize all values as strings, include empty cells, embed per‑cell style information, treat the first row as a header, and pretty‑print the JSON with indentation. It then saves the whole worksheet to a JSON file and demonstrates exporting a defined range to a JSON string.
class ExportWorksheetToJson
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data
            worksheet.Cells["A1"].PutValue("Product");
            worksheet.Cells["B1"].PutValue("Price");
            worksheet.Cells["C1"].PutValue("InStock");
            worksheet.Cells["A2"].PutValue("Apple");
            worksheet.Cells["B2"].PutValue(1.2);
            worksheet.Cells["C2"].PutValue(true);
            worksheet.Cells["A3"].PutValue("Banana");
            worksheet.Cells["B3"].PutValue(0.8);
            worksheet.Cells["C3"].PutValue(false);

            // Apply a header style
            Style headerStyle = workbook.CreateStyle();
            headerStyle.Font.IsBold = true;
            headerStyle.ForegroundColor = System.Drawing.Color.LightGray;
            headerStyle.Pattern = BackgroundType.Solid;
            worksheet.Cells["A1:C1"].SetStyle(headerStyle);

            // Apply a numeric style to price column
            Style priceStyle = workbook.CreateStyle();
            priceStyle.Number = 2; // two decimal places
            worksheet.Cells["B2:B3"].SetStyle(priceStyle);

            // Configure JSON export options for custom serialization
            JsonSaveOptions jsonOptions = new JsonSaveOptions
            {
                ExportAsString = true,      // serialize all cell values as strings
                ExportEmptyCells = true,    // include empty cells as null
                ExportStylePool = false,    // export style information per cell
                HasHeaderRow = true,        // first row is treated as header
                Indent = "    ",            // pretty‑print JSON with 4‑space indentation
                ToExcelStruct = true        // include Excel structure (styles) in JSON
            };

            // Export the entire worksheet to a JSON file using the configured options
            string outputPath = "WorksheetExport.json";
            workbook.Save(outputPath, jsonOptions);
            Console.WriteLine($"Workbook exported to JSON file: {Path.GetFullPath(outputPath)}");

            // Obtain the JSON string for a specific range
            AsposeRange range = worksheet.Cells.CreateRange("A1:C3");
            string jsonString = JsonUtility.ExportRangeToJson(range, jsonOptions);
            Console.WriteLine("JSON for selected range:");
            Console.WriteLine(jsonString);
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred:");
            Console.WriteLine(ex.Message);
        }
    }
}
