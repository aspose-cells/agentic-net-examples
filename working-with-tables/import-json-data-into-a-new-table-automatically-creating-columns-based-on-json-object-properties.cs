// Title: Import a Heterogeneous JSON Array into an Excel Table with Automatic Column Creation – Aspose.Cells for .NET (C#)
// Description: Demonstrates how to use Aspose.Cells' JsonUtility.ImportData with JsonLayoutOptions.ArrayAsTable to load a JSON array where each object may have different fields. The code creates a workbook, treats each JSON object as a row, generates columns for all distinct property names, and saves the result as an Excel file.
// Keywords: Aspose.Cells JSON import | JsonUtility ImportData C# | JsonLayoutOptions ArrayAsTable | automatic column generation Excel | convert JSON to Excel table | dynamic JSON schema Excel export | C# Aspose.Cells example
// Common Searches: Aspose.Cells import JSON array as table | C# JsonUtility.ImportData dynamic columns | How to create Excel table from heterogeneous JSON | Aspose.Cells JsonLayoutOptions ArrayAsTable example | Convert API JSON response to Excel with Aspose
// Developer Intent: Load JSON data into a new Excel worksheet and let Aspose.Cells automatically create columns for every unique property found in the JSON objects.
// Use Cases: Transform API responses with varying fields into a ready‑to‑use Excel report without manual column mapping. | Export data from loosely structured JSON files to Excel for analysis, preserving all available attributes. | Build a generic data‑export routine that adapts to any JSON schema by generating columns on‑the‑fly.
// AI Prompts: Generate code to rename the Excel table created after importing JSON with JsonUtility.ImportData. | Show how to apply styling—bold headers, autofit columns, and table formatting—after using JsonLayoutOptions.ArrayAsTable. | Explain the steps to import nested JSON arrays into separate worksheets using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

// Demonstrates how to use Aspose.Cells' JsonUtility.ImportData with JsonLayoutOptions.ArrayAsTable to load a JSON array where each object may have different fields. The code creates a workbook, treats each JSON object as a row, generates columns for all distinct property names, and saves the result as an Excel file.
class Program
{
    static void Main()
    {
        // Sample JSON array with objects that have different properties.
        // Aspose.Cells will automatically create columns for all distinct property names.
        string json = @"[
            { ""Name"": ""Alice"",   ""Age"": 30, ""City"": ""New York"" },
            { ""Name"": ""Bob"",     ""Age"": 25, ""Country"": ""USA"" },
            { ""Name"": ""Charlie"", ""City"": ""Los Angeles"", ""Salary"": 50000 }
        ]";

        // Create a new workbook and get the first worksheet.
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Configure layout options to treat the JSON array as a table.
        JsonLayoutOptions options = new JsonLayoutOptions();
        options.ArrayAsTable = true;   // each object becomes a row, properties become columns

        // Import the JSON data starting at cell A1 (row 0, column 0).
        JsonUtility.ImportData(json, sheet.Cells, 0, 0, options);

        // Save the workbook to an Excel file.
        workbook.Save("JsonImportedTable.xlsx");
    }
}
