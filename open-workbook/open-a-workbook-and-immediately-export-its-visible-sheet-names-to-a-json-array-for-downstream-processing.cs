using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Cells;

// Author: Aspose.Cells .NET example – export visible worksheet names to JSON
class ExportVisibleSheetNames
{
    static void Main()
    {
        // Load the workbook (replace with your actual file path)
        string workbookPath = "input.xlsx";
        Workbook workbook = new Workbook(workbookPath);

        // Collect names of visible worksheets
        List<string> visibleSheetNames = new List<string>();
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // NOTE: Adjust the visibility check if the property name differs in your Aspose.Cells version.
            // Common properties: sheet.IsVisible or sheet.Visibility == SheetVisibility.Visible
            if (sheet.IsVisible) // placeholder for visibility check
            {
                visibleSheetNames.Add(sheet.Name);
            }
        }

        // Serialize the list of names to a JSON array
        string json = JsonSerializer.Serialize(visibleSheetNames, new JsonSerializerOptions { WriteIndented = true });

        // Write JSON to a file
        string outputPath = "VisibleSheetNames.json";
        File.WriteAllText(outputPath, json);

        Console.WriteLine($"Exported {visibleSheetNames.Count} visible sheet name(s) to '{outputPath}'.");
    }
}