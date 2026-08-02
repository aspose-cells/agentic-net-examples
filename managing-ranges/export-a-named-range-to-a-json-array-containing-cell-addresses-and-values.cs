// Title: Export a Named Range to JSON Using Aspose.Cells for .NET (C#)
// Description: This C# sample shows how to build a workbook with Aspose.Cells, define a named range, retrieve it via GetRangeByName, gather each cell's address and value, and serialize the result into a formatted JSON array with System.Text.Json. The JSON output is ready for API payloads, logging, or data exchange.
// Keywords: Aspose.Cells | C# | named range | JSON export | GetRangeByName | System.Text.Json | cell address | cell value | Excel to JSON | Aspose.Cells example | GitHub sample
// Common Searches: Aspose.Cells export named range to JSON C# | How to get cell address and value from a named range in Aspose.Cells | Serialize Aspose.Cells range as JSON array | GetRangeByName example Aspose.Cells .NET | Convert Excel named range to JSON with System.Text.Json
// Developer Intent: Generate a JSON array that lists every cell’s address and its value from a specified named range.
// Use Cases: Create a JSON payload for a web service by converting a defined range into address/value pairs. | Debug or audit Excel data by printing the contents of a named range as readable JSON. | Transfer spreadsheet information to front‑end applications that consume JSON structures.
// AI Prompts: Write C# code with Aspose.Cells that exports a named range to a JSON array where each element contains the cell address and value. | Show how to handle different data types (strings, numbers, dates) when serializing Aspose.Cells range values to JSON. | Modify the example to omit empty cells from the JSON output.

using System;
using System.Collections.Generic;
using System.Text.Json;
using Aspose.Cells;

// This C# sample shows how to build a workbook with Aspose.Cells, define a named range, retrieve it via GetRangeByName, gather each cell's address and value, and serialize the result into a formatted JSON array with System.Text.Json. The JSON output is ready for API payloads, logging, or data exchange.
class ExportNamedRangeToJson
{
    static void Main()
    {
        try
        {
            // Create a new workbook and add sample data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Name");
            sheet.Cells["B1"].PutValue("Age");
            sheet.Cells["A2"].PutValue("John");
            sheet.Cells["B2"].PutValue(30);
            sheet.Cells["A3"].PutValue("Alice");
            sheet.Cells["B3"].PutValue(25);

            // Define a named range that covers the data
            int nameIdx = workbook.Worksheets.Names.Add("People");
            workbook.Worksheets.Names[nameIdx].RefersTo = "=Sheet1!A1:B3";

            // Retrieve the named range using GetRangeByName
            Aspose.Cells.Range namedRange = workbook.Worksheets.GetRangeByName("People");
            if (namedRange == null)
            {
                Console.WriteLine("Named range not found.");
                return;
            }

            // Collect cell addresses and values into a list
            var cellsInfo = new List<Dictionary<string, object>>();
            foreach (Cell cell in namedRange)
            {
                var entry = new Dictionary<string, object>
                {
                    ["Address"] = cell.Name,   // e.g., "A1"
                    ["Value"] = cell.Value      // actual cell value
                };
                cellsInfo.Add(entry);
            }

            // Serialize the list to a JSON array
            string jsonResult = JsonSerializer.Serialize(cellsInfo, new JsonSerializerOptions { WriteIndented = true });
            Console.WriteLine(jsonResult);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
