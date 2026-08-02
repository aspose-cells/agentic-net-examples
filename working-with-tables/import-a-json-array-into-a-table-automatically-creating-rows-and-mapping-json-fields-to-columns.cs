// Title: Import a JSON Array as an Excel Table with Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create a Workbook, configure JsonLayoutOptions with ArrayAsTable, and use JsonUtility.ImportData to load a JSON array into a worksheet starting at A1. The utility automatically adds a header row and populates rows for each object, then saves the result as an .xlsx file.
// Keywords: Aspose.Cells | C# | .NET | JSON to Excel | JsonUtility ImportData | JsonLayoutOptions | ArrayAsTable | Excel table from JSON | data import | workbook automation
// Common Searches: Aspose.Cells import JSON array as table | JsonUtility ImportData C# example | ArrayAsTable option Aspose.Cells | convert JSON list to Excel worksheet .NET | load JSON into Excel table programmatically
// Developer Intent: Load a JSON array into an Excel worksheet as a structured table with automatic column headers using Aspose.Cells for .NET.
// Use Cases: Turn API response JSON into a ready‑to‑analyze Excel report. | Create a generic data‑import routine that maps JSON fields to worksheet columns without manual mapping. | Automate conversion of flat JSON logs into tabular Excel sheets for business intelligence.
// AI Prompts: Generate C# code that applies bold styling and background color to the header row after importing the JSON array with Aspose.Cells. | Show how to import nested JSON objects into separate worksheets using JsonUtility and Aspose.Cells. | Explain how to change the start cell, worksheet index, or add a custom table name when importing JSON data with JsonLayoutOptions.

using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

// Demonstrates how to create a Workbook, configure JsonLayoutOptions with ArrayAsTable, and use JsonUtility.ImportData to load a JSON array into a worksheet starting at A1. The utility automatically adds a header row and populates rows for each object, then saves the result as an .xlsx file.
class JsonImportExample
{
    static void Main()
    {
        // JSON array to be imported
        string json = @"[
            { ""Name"": ""Alice"", ""Age"": 30, ""City"": ""New York"" },
            { ""Name"": ""Bob"", ""Age"": 25, ""City"": ""Los Angeles"" },
            { ""Name"": ""Charlie"", ""Age"": 35, ""City"": ""Chicago"" }
        ]";

        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Configure layout options to treat the JSON array as a table (adds header row)
        JsonLayoutOptions options = new JsonLayoutOptions
        {
            ArrayAsTable = true
        };

        // Import the JSON data starting at cell A1 (row 0, column 0)
        JsonUtility.ImportData(json, worksheet.Cells, 0, 0, options);

        // Save the workbook to an Excel file
        workbook.Save("JsonArrayTable.xlsx");
    }
}
