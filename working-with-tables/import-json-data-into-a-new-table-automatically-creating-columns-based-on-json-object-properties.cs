// Title: Import JSON Array into Excel Table with Auto‑Generated Columns using Aspose.Cells for .NET
// Description: C# example that creates a workbook, sets JsonLayoutOptions.ArrayAsTable = true, and uses JsonUtility.ImportData to load a JSON array into the first worksheet starting at A1. Columns are created automatically from the JSON object properties, and the result is saved as JsonImported.xlsx.
// Keywords: Aspose.Cells JSON import | JsonUtility ImportData | ArrayAsTable option | auto‑create Excel columns from JSON | C# export JSON to Excel | JSON to Excel table .NET | Aspose.Cells example
// Common Searches: Aspose.Cells import JSON as table C# | JsonLayoutOptions ArrayAsTable example | Create Excel columns from JSON keys using Aspose | How to load JSON array into Excel with Aspose.Cells | C# JsonUtility.ImportData usage
// Developer Intent: Generate an Excel worksheet where each JSON object becomes a row and each property becomes a column without manually defining the schema.
// Use Cases: Convert API response data (JSON) into a sortable Excel report for business analysts. | Transform a list of user profiles stored in JSON into a structured table for HR processing. | Export configuration settings defined in JSON to an Excel document for audit trails.
// AI Prompts: Write C# code that uses Aspose.Cells to import a JSON array into a worksheet as a table with columns generated from the JSON keys. | Explain how JsonLayoutOptions.ArrayAsTable influences the import process and how to change the start cell for the data. | Show how to handle nested JSON objects by splitting each level into separate worksheets with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

// C# example that creates a workbook, sets JsonLayoutOptions.ArrayAsTable = true, and uses JsonUtility.ImportData to load a JSON array into the first worksheet starting at A1. Columns are created automatically from the JSON object properties, and the result is saved as JsonImported.xlsx.
class Program
{
    static void Main()
    {
        // Sample JSON array; each object will become a row and its properties become columns
        string json = @"[
            { ""Name"": ""Alice"", ""Age"": 30, ""City"": ""London"" },
            { ""Name"": ""Bob"",   ""Age"": 25, ""City"": ""Paris""  },
            { ""Name"": ""Carol"", ""Age"": 28, ""City"": ""Berlin"" }
        ]";

        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Configure layout options to treat the JSON array as a table (auto‑create columns)
        JsonLayoutOptions options = new JsonLayoutOptions
        {
            ArrayAsTable = true
        };

        // Import the JSON data starting at cell A1 (row 0, column 0)
        JsonUtility.ImportData(json, worksheet.Cells, 0, 0, options);

        // Save the workbook to an Excel file
        workbook.Save("JsonImported.xlsx");
    }
}
