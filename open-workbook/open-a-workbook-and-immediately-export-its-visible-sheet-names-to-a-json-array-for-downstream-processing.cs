// Title: Extract visible worksheet names from an Excel file and output as a JSON array using Aspose.Cells for .NET
// AI Prompts: Generate C# code that opens an .xlsx workbook with Aspose.Cells, selects only the worksheets where IsVisible is true, and produces a JSON array of their names. | Update the sample to write the JSON output to a file called "visibleSheets.json" and add exception handling for missing input files. | Create a reusable function that takes a workbook path, returns a List<string> of visible worksheet names, and serializes it using System.Text.Json.
// Common Searches: C# Aspose.Cells get names of visible worksheets from an Excel file | How to export visible sheet names to a JSON array in .NET | Filter out hidden worksheets when reading Excel with Aspose.Cells | Serialize a list of Excel worksheet names to JSON using System.Text.Json
// Tags: Aspose.Cells visible worksheet extraction | export worksheet names to JSON C# | filter hidden sheets Aspose.Cells | System.Text.Json serialize worksheet list | C# enumerate visible worksheets

using System;
using System.Collections.Generic;
using System.Text.Json;
using Aspose.Cells;

// The example loads an .xlsx workbook with Aspose.Cells, iterates through its worksheets, collects the names of those marked as visible, serializes the list to a JSON array using System.Text.Json, and writes the JSON string to the console.
class Program
{
    static void Main()
    {
        // Load the workbook from a file (replace with your actual path)
        var workbook = new Workbook("input.xlsx");

        // Collect names of visible worksheets
        var visibleSheetNames = new List<string>();
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            if (sheet.IsVisible)
            {
                visibleSheetNames.Add(sheet.Name);
            }
        }

        // Convert the list of names to a JSON array
        string json = JsonSerializer.Serialize(visibleSheetNames);

        // Output the JSON (could be written to a file or passed downstream)
        Console.WriteLine(json);
    }
}
