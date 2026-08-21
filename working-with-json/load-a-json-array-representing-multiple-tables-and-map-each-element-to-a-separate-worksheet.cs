// Title: C# – Load a JSON Array and Export Each Table to a Separate Excel Worksheet with Aspose.Cells
// Description: Demonstrates how to parse a JSON array where each element defines a table, create an Aspose.Cells workbook, and import each table into its own worksheet using JsonUtility.ImportData with JsonLayoutOptions. The workbook is saved as MultipleTables.xlsx.
// Keywords: Aspose.Cells C# JSON import | JsonUtility ImportData example | JsonLayoutOptions array as table | multiple worksheets from JSON | export JSON to Excel .NET | Aspose.Cells GitHub sample | C# convert JSON array to Excel | Excel workbook from JSON tables | Aspose.Cells JSON to separate sheets | load JSON array Aspose.Cells
// Common Searches: Aspose.Cells import JSON array to separate sheets | C# map each JSON object to its own Excel worksheet | JsonUtility.ImportData multiple tables example | How to use JsonLayoutOptions with Aspose.Cells | Create Excel workbook from JSON using Aspose.Cells .NET
// Developer Intent: The developer needs to read a JSON array where each element represents a distinct table and write each table to a separate worksheet in an Excel file using Aspose.Cells for .NET.
// Use Cases: Generate a multi‑sheet report where each API response segment is placed on its own tab. | Convert a configuration file containing several data tables into an Excel workbook for downstream analysis. | Automate the export of relational data stored in JSON format to a single, distributable Excel file.
// AI Prompts: Show how to set column widths automatically after importing each JSON table with Aspose.Cells. | Provide code that adds a summary worksheet with hyperlinks to all imported tables. | Explain strategies for handling nested JSON objects when using JsonUtility.ImportData.

using System;
using System.Text.Json;
using Aspose.Cells;
using Aspose.Cells.Utility;

// Demonstrates how to parse a JSON array where each element defines a table, create an Aspose.Cells workbook, and import each table into its own worksheet using JsonUtility.ImportData with JsonLayoutOptions. The workbook is saved as MultipleTables.xlsx.
class Program
{
    static void Main()
    {
        // Sample JSON array where each element represents a separate table
        string json = @"[
            {
                ""Headers"": [""Name"", ""Age"", ""City""],
                ""Rows"": [
                    [""Alice"", 30, ""NY""],
                    [""Bob"", 25, ""LA""]
                ]
            },
            {
                ""Headers"": [""Product"", ""Price""],
                ""Rows"": [
                    [""Laptop"", 1200],
                    [""Phone"", 800]
                ]
            }
        ]";

        // Parse the JSON string into a JsonDocument
        JsonDocument doc = JsonDocument.Parse(json);
        JsonElement root = doc.RootElement;

        // Create a new workbook (creation rule)
        Workbook workbook = new Workbook();

        // Define layout options for importing JSON as tables
        JsonLayoutOptions layoutOptions = new JsonLayoutOptions
        {
            ArrayAsTable = true,   // Treat arrays as tables
            IgnoreTitle = true     // Ignore titles to keep data clean
        };

        int sheetIndex = 0;

        // Iterate over each element in the JSON array
        foreach (JsonElement tableElement in root.EnumerateArray())
        {
            // Use the first default worksheet for the first table,
            // otherwise add a new worksheet for each subsequent table
            Worksheet sheet;
            if (sheetIndex == 0)
            {
                sheet = workbook.Worksheets[0];
                sheet.Name = $"Table{sheetIndex + 1}";
            }
            else
            {
                sheet = workbook.Worksheets.Add($"Table{sheetIndex + 1}");
            }

            // Convert the current JSON element back to a JSON string
            string tableJson = tableElement.GetRawText();

            // Import the JSON data into the worksheet starting at cell A1
            JsonUtility.ImportData(tableJson, sheet.Cells, 0, 0, layoutOptions);

            sheetIndex++;
        }

        // Save the workbook to an Excel file (save rule)
        workbook.Save("MultipleTables.xlsx");
    }
}
