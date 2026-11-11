// Title: Export only visible worksheet names from an Excel workbook to a JSON array using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells get names of visible worksheets and output as JSON | How to list only visible sheets from an Excel file with Aspose.Cells .NET | Serialize visible worksheet names to a JSON array using Aspose.Cells in C# | Aspose.Cells filter hidden worksheets and export sheet names to JSON | Read an Excel workbook and generate a JSON array of visible sheet names in .NET
// Tags: Aspose.Cells visible worksheet name extraction | export sheet names to JSON C# | filter hidden worksheets Aspose.Cells | System.Text.Json serialize Excel sheet list | list visible worksheets .NET

using System;
using System.Collections.Generic;
using System.Text.Json;
using Aspose.Cells;

// The example loads an Excel file with Aspose.Cells, iterates through its worksheets, collects the names of those marked as visible, serializes the list to a JSON array using System.Text.Json, and writes the JSON string to the console for downstream processing.
class Program
{
    static void Main()
    {
        // Load the workbook from a file
        var workbook = new Workbook("input.xlsx");

        // Gather names of all visible worksheets
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

        // Output the JSON (can be redirected, saved, or passed downstream)
        Console.WriteLine(json);
    }
}
