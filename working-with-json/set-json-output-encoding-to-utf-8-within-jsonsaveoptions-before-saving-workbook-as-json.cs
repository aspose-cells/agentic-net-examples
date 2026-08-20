// Title: Export Excel Range to UTF‑8 JSON with Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create a workbook, fill cells A1:B3, define a range, configure JsonSaveOptions (header row, omit empty cells, custom indentation), convert the range to a JSON string, and write the result to a file using explicit UTF‑8 encoding.
// Keywords: Aspose.Cells | C# | JsonSaveOptions | UTF-8 JSON export | Excel to JSON | export range as JSON | write JSON file with encoding | Unicode JSON Aspose | Aspose.Cells .NET example | Excel data to UTF-8
// Common Searches: Aspose.Cells export range to JSON UTF-8 | C# save JSON with UTF-8 using Aspose.Cells | JsonSaveOptions encoding example | How to write Excel data as UTF-8 JSON in .NET | Aspose.Cells JSON output Unicode characters
// Developer Intent: Generate a UTF‑8 encoded JSON file from a selected Excel range using Aspose.Cells for .NET.
// Use Cases: Create API payloads by converting worksheet tables to UTF‑8 JSON. | Preserve international characters (e.g., accented or Asian scripts) when exporting Excel data. | Produce compact, indented JSON for downstream processing pipelines.
// AI Prompts: Show C# code that sets JsonSaveOptions and saves the JSON output with UTF‑8 encoding using Aspose.Cells. | Explain step‑by‑step how to export a worksheet range to a UTF‑8 JSON file with custom indentation and error handling. | Modify the example to include Unicode text in the cells and verify the saved file remains UTF‑8 encoded.

using System;
using System.IO;
using System.Text;
using Aspose.Cells;
using Aspose.Cells.Utility;

// Demonstrates how to create a workbook, fill cells A1:B3, define a range, configure JsonSaveOptions (header row, omit empty cells, custom indentation), convert the range to a JSON string, and write the result to a file using explicit UTF‑8 encoding.
class JsonUtf8SaveExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data
            sheet.Cells["A1"].PutValue("Name");
            sheet.Cells["B1"].PutValue("Age");
            sheet.Cells["A2"].PutValue("John");
            sheet.Cells["B2"].PutValue(30);
            sheet.Cells["A3"].PutValue("Jane");
            sheet.Cells["B3"].PutValue(25);

            // Define the range to be exported (A1:B3)
            Aspose.Cells.Range exportRange = sheet.Cells.CreateRange("A1:B3");

            // Configure JSON save options
            JsonSaveOptions jsonOptions = new JsonSaveOptions
            {
                HasHeaderRow = true,
                ExportEmptyCells = false,
                ExportAsString = false,
                Indent = "  "
            };

            // Convert the range to a JSON string using the options
            string jsonContent = exportRange.ToJson(jsonOptions);

            // Write the JSON string to a file using UTF-8 encoding explicitly
            string outputPath = "output_utf8.json";
            File.WriteAllText(outputPath, jsonContent, Encoding.UTF8);

            Console.WriteLine($"JSON file saved with UTF-8 encoding at: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}
