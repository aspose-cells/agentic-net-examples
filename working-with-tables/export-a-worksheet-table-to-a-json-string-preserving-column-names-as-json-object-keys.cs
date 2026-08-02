// Title: C# – Export an Excel worksheet table to JSON with column names as keys using Aspose.Cells
// Description: Demonstrates how to create a workbook, fill a header row and data rows, define a range, configure ExportRangeToJsonOptions (header row, numeric values, empty cells, custom indentation), and call JsonUtility.ExportRangeToJson to generate a formatted JSON string printed to the console.
// Keywords: Aspose.Cells | ExportRangeToJson | JsonUtility | C# Excel to JSON | Export worksheet range to JSON | pretty‑printed JSON | header row as keys | .NET Excel JSON conversion | sample code | GitHub example
// Common Searches: Aspose.Cells export range to JSON C# | How to convert Excel table to JSON with column headers | Export Excel range as JSON string using Aspose.Cells | JsonUtility ExportRangeToJsonOptions example | C# pretty printed JSON from worksheet
// Developer Intent: Create a JSON string from a worksheet range where each row becomes an object and the first row supplies the property names.
// Use Cases: Return JSON data directly from an in‑memory workbook for a web API without writing a file. | Serialize tabular data for front‑end JavaScript consumption while preserving original column names. | Generate a readable JSON payload with indentation and empty‑cell placeholders for logging or debugging.
// AI Prompts: Generate C# code that uses Aspose.Cells to export a worksheet range to a formatted JSON string, using the first row as property names. | Show how to modify ExportRangeToJsonOptions to output numeric values as strings and to omit rows that are completely empty. | Provide a snippet that deserializes the JSON produced by JsonUtility.ExportRangeToJson into a List<T> with System.Text.Json.

using System;
using Aspose.Cells;
using Aspose.Cells.Utility;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsJsonExportDemo
{
    // Demonstrates how to create a workbook, fill a header row and data rows, define a range, configure ExportRangeToJsonOptions (header row, numeric values, empty cells, custom indentation), and call JsonUtility.ExportRangeToJson to generate a formatted JSON string printed to the console.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

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

                cells["A4"].PutValue("Bob");
                cells["B4"].PutValue(28);
                cells["C4"].PutValue("Paris");

                // Define the range that includes the header and all data rows
                // startRow = 0 (A1), startColumn = 0 (A), totalRows = 4, totalColumns = 3
                AsposeRange exportRange = cells.CreateRange(0, 0, 4, 3);

                // Configure export options
                ExportRangeToJsonOptions jsonOptions = new ExportRangeToJsonOptions
                {
                    HasHeaderRow = true,
                    ExportAsString = false,
                    ExportEmptyCells = true,
                    Indent = "  "
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
