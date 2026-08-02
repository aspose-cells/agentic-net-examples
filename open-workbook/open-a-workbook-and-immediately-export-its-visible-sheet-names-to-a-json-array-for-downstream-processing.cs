// Title: Export Visible Worksheet Names to JSON with Aspose.Cells for .NET (C#)
// Description: Loads an Excel workbook using Aspose.Cells, filters worksheets by the IsVisible flag, collects their names, and serializes the list to a JSON array with System.Text.Json. The JSON string is written to the console (or can be saved to a file).
// Keywords: Aspose.Cells | C# export worksheet names | visible Excel sheets | JSON serialization .NET | System.Text.Json Excel | list visible worksheets | Excel workbook to JSON | Aspose.Cells GetVisibleSheets | .NET Excel automation
// Common Searches: Aspose.Cells get visible sheet names C# | serialize Excel worksheet names to JSON | C# export visible worksheets as JSON array | how to list only visible sheets in an Excel file using Aspose | convert Excel sheet names to JSON with .NET
// Developer Intent: Extract the names of all visible worksheets from an Excel file and return them as a JSON array.
// Use Cases: Populate a web UI with tabs that correspond only to visible sheets. | Send a JSON payload of visible sheet names to a downstream service for selective processing. | Create audit logs that record visible worksheet names in a machine‑readable format.
// AI Prompts: Generate C# code that opens an Excel workbook with Aspose.Cells and returns a JSON array of visible worksheet names. | Write a method that accepts a file path, extracts visible sheet names using Aspose.Cells, and returns a JSON string with error handling. | Provide an example that writes the JSON array of visible sheet names to a file instead of the console, using Aspose.Cells and System.Text.Json.

using System;
using System.Collections.Generic;
using System.Text.Json;
using Aspose.Cells;

// Loads an Excel workbook using Aspose.Cells, filters worksheets by the IsVisible flag, collects their names, and serializes the list to a JSON array with System.Text.Json. The JSON string is written to the console (or can be saved to a file).
class ExportVisibleSheetNames
{
    static void Main()
    {
        // Load the workbook from a file
        string workbookPath = "input.xlsx";
        Workbook workbook = new Workbook(workbookPath);

        // Gather the names of all visible worksheets
        List<string> visibleSheetNames = new List<string>();
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            if (sheet.IsVisible)
            {
                visibleSheetNames.Add(sheet.Name);
            }
        }

        // Convert the list of names to a JSON array string
        string jsonArray = JsonSerializer.Serialize(visibleSheetNames);

        // Output the JSON (or write to a file as needed)
        Console.WriteLine(jsonArray);
    }
}
