// Title: C# Example: Export an Aspose.Cells Named Range to JSON (cell addresses & values)
// Description: Demonstrates how to create a workbook, define a named range, iterate its cells, capture each cell's address and raw value, and serialize the collection to a formatted JSON array using System.Text.Json.
// Keywords: Aspose.Cells | C# | named range | export to JSON | cell address | cell value | serialize range | Workbook example | System.Text.Json | Aspose.Cells sample code | range to JSON
// Common Searches: Aspose.Cells export named range to JSON C# | Get cell addresses from a named range Aspose.Cells | Serialize Aspose.Cells range as JSON | C# code to convert named range to JSON | Aspose.Cells JSON array of cell values
// Developer Intent: Generate a JSON array that lists every cell address and its corresponding value from a specified named range in an Aspose.Cells workbook.
// Use Cases: Create API payloads by converting a table area defined as a named range into JSON. | Log or audit specific spreadsheet sections for change tracking. | Feed spreadsheet data to a web front‑end without exposing the whole workbook.
// AI Prompts: Write a reusable C# method that accepts a Workbook and a named range name, then returns a JSON string of address/value pairs with optional indentation. | Add error handling to the sample for missing named ranges, empty cells, and unsupported data types. | Extend the example to allow custom JSON property names and to serialize dates in ISO 8601 format.

using System;
using System.Collections.Generic;
using System.Text.Json;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range; // Alias to avoid conflict with System.Range

namespace ExportNamedRangeToJson
{
    // Simple DTO to hold cell address and its value
    // Demonstrates how to create a workbook, define a named range, iterate its cells, capture each cell's address and raw value, and serialize the collection to a formatted JSON array using System.Text.Json.
    public class CellInfo
    {
        public string? Address { get; set; }
        public object? Value { get; set; }
    }

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

                // Populate sample data
                cells["A1"].PutValue("Name");
                cells["B1"].PutValue("Age");
                cells["A2"].PutValue("John");
                cells["B2"].PutValue(30);
                cells["A3"].PutValue("Alice");
                cells["B3"].PutValue(25);

                // Define a named range that covers the data (A1:B3)
                int nameIndex = workbook.Worksheets.Names.Add("MyRange");
                workbook.Worksheets.Names[nameIndex].RefersTo = "=Sheet1!A1:B3";

                // Retrieve the range using the name
                Name namedRange = workbook.Worksheets.Names["MyRange"];
                AsposeRange range = namedRange.GetRange();

                // Collect address/value pairs from the range
                List<CellInfo> cellInfos = new List<CellInfo>();
                foreach (Cell cell in range)
                {
                    cellInfos.Add(new CellInfo
                    {
                        Address = cell.Name,   // e.g., "A1"
                        Value = cell.Value     // raw value (string, number, etc.)
                    });
                }

                // Serialize the list to a JSON array
                string json = JsonSerializer.Serialize(
                    cellInfos,
                    new JsonSerializerOptions { WriteIndented = true });

                // Output the JSON
                Console.WriteLine(json);
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
