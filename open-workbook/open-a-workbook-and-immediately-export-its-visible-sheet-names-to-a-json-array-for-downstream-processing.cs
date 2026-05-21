using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Cells;

class ExportVisibleSheetNames
{
    static void Main()
    {
        // Load the workbook from a file (replace with your actual file path)
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

        // Save the JSON array to a file (replace with your desired output path)
        string jsonPath = "visibleSheets.json";
        File.WriteAllText(jsonPath, jsonArray);

        Console.WriteLine($"Visible sheet names have been exported to '{jsonPath}'.");
    }
}