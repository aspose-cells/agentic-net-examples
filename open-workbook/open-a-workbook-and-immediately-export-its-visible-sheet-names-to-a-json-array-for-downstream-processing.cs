// Title: Export Visible Worksheet Names to JSON with Aspose.Cells for .NET (C#)
// Description: Loads an Excel workbook using Aspose.Cells, iterates through its worksheets, collects the names of those marked as visible, serializes the list to a JSON array with System.Text.Json, and writes the result to a file. Ideal for downstream processing or UI generation.
// Keywords: Aspose.Cells | C# | .NET | export visible sheet names | Excel workbook JSON | worksheet visibility | System.Text.Json | serialize worksheet list | write JSON file | sample code
// Common Searches: Aspose.Cells get visible sheet names C# | export Excel sheet names to JSON .NET | list visible worksheets using Aspose.Cells | C# serialize worksheet names to JSON array | write visible sheet names to file Aspose
// Developer Intent: Retrieve the names of all visible worksheets from an Excel file and save them as a JSON array.
// Use Cases: Provide a JSON payload of visible sheet names to a web service that processes only user‑visible worksheets. | Dynamically generate UI tabs or navigation menus based on the visible worksheets in a user‑uploaded Excel file. | Create an audit log of visible worksheets by storing their names in a JSON file for later review.
// AI Prompts: Generate C# code that opens an Excel workbook with Aspose.Cells, extracts visible worksheet names, and returns a JSON array string. | Write a reusable method that accepts an input path, uses Aspose.Cells to collect visible sheet names, and writes the JSON array to a specified output file. | Explain how to extend the example to filter worksheets by a custom attribute before serializing their names to JSON.

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Cells;

// Loads an Excel workbook using Aspose.Cells, iterates through its worksheets, collects the names of those marked as visible, serializes the list to a JSON array with System.Text.Json, and writes the result to a file. Ideal for downstream processing or UI generation.
class ExportVisibleSheetNames
{
    static void Main()
    {
        // Load the workbook from a file (replace with your actual file path)
        string workbookPath = "input.xlsx";
        Workbook workbook = new Workbook(workbookPath);

        // Gather names of all visible worksheets
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

        // Write the JSON array to a file (or use the string directly downstream)
        string jsonOutputPath = "visibleSheets.json";
        File.WriteAllText(jsonOutputPath, jsonArray);

        // Optional: display the result
        Console.WriteLine("Visible sheet names exported to JSON:");
        Console.WriteLine(jsonArray);
    }
}
