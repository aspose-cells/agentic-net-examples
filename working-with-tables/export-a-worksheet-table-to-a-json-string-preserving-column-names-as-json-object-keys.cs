// Title: Export Excel table to JSON with column headers as keys using Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create a workbook, fill a header row and data rows, define a range, set JsonSaveOptions (HasHeaderRow = true, ExportAsString = false, Indent), and call JsonUtility.ExportRangeToJson to generate a formatted JSON string where each column name becomes a JSON property.
// Keywords: Aspose.Cells JSON export C# | Export Excel range to JSON | HasHeaderRow true | JsonUtility ExportRangeToJson | pretty printed JSON from worksheet | Excel to JSON with headers | C# Aspose.Cells example
// Common Searches: Aspose.Cells export range to JSON C# | How to keep column names as keys when converting Excel to JSON | JsonSaveOptions HasHeaderRow usage | Pretty print JSON from Excel worksheet using Aspose.Cells | Export Excel table as JSON string .NET
// Developer Intent: Create a JSON string from a worksheet table, using the first row as property names and preserving original data types.
// Use Cases: Return lookup data (e.g., Name, Age, City) from an API in JSON format without losing numeric types. | Generate readable JSON logs or configuration files directly from Excel worksheets. | Provide front‑end applications with structured data extracted from Excel tables.
// AI Prompts: Write C# code that uses Aspose.Cells to export a selected range to JSON with column headers as keys and pretty‑print indentation. | Explain the impact of JsonSaveOptions properties HasHeaderRow and ExportAsString on the JSON output when exporting an Excel range. | Show how to handle exceptions while exporting a worksheet range to JSON with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsJsonExport
{
    // Demonstrates how to create a workbook, fill a header row and data rows, define a range, set JsonSaveOptions (HasHeaderRow = true, ExportAsString = false, Indent), and call JsonUtility.ExportRangeToJson to generate a formatted JSON string where each column name becomes a JSON property.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Populate the worksheet with a header row and some data rows
                cells["A1"].PutValue("Name");
                cells["B1"].PutValue("Age");
                cells["C1"].PutValue("City");

                cells["A2"].PutValue("John");
                cells["B2"].PutValue(30);
                cells["C2"].PutValue("New York");

                cells["A3"].PutValue("Alice");
                cells["B3"].PutValue(25);
                cells["C3"].PutValue("London");

                // Define the range that includes the header and data (A1:C3)
                Aspose.Cells.Range exportRange = cells.CreateRange("A1:C3");

                // Configure JSON export options:
                // - HasHeaderRow = true ensures column names become JSON object keys
                // - ExportAsString = false (default) keeps original data types
                JsonSaveOptions jsonOptions = new JsonSaveOptions
                {
                    HasHeaderRow = true,
                    ExportAsString = false,
                    Indent = "  " // optional pretty‑print
                };

                // Export the range to a JSON string
                string jsonResult = JsonUtility.ExportRangeToJson(exportRange, jsonOptions);

                // Output the JSON string
                Console.WriteLine("Exported JSON:");
                Console.WriteLine(jsonResult);
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
