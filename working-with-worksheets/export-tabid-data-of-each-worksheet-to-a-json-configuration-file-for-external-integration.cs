// Title: Export Worksheet TabId and Name to JSON with Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create or load a Workbook, iterate through all worksheets, capture each sheet's Name and TabId, serialize the collection with System.Text.Json, and write the result to a file named WorksheetTabIds.json.
// Keywords: Aspose.Cells | C# | .NET | export worksheet TabId | worksheet metadata JSON | serialize workbook sheets | System.Text.Json | Workbook TabId property | save sheet identifiers | code example
// Common Searches: Aspose.Cells export worksheet TabId to JSON | How to get worksheet TabId in C# | Serialize Aspose.Cells sheet metadata | Write worksheet identifiers to JSON file | C# example for exporting workbook sheet info
// Developer Intent: Generate a JSON file that lists every worksheet's Name and TabId from an Aspose.Cells workbook.
// Use Cases: Provide external systems with a lightweight mapping of sheet names to internal TabId values. | Create version‑controlled configuration files for dynamic sheet selection in reporting pipelines. | Supply client‑side applications with sheet identifiers without loading the full workbook.
// AI Prompts: Write C# code that loads an existing workbook, extracts each worksheet's Name and TabId, and saves the data as formatted JSON using System.Text.Json. | Extend the sample to also include each worksheet's visibility state (Visible, Hidden, VeryHidden) in the exported JSON. | Add robust error handling to the JSON export routine to capture I/O exceptions and permission issues.

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Cells;

// Demonstrates how to create or load a Workbook, iterate through all worksheets, capture each sheet's Name and TabId, serialize the collection with System.Text.Json, and write the result to a file named WorksheetTabIds.json.
class ExportWorksheetTabIds
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook(); // create

        // Add sample worksheets for demonstration
        workbook.Worksheets[0].Name = "Sheet1";
        workbook.Worksheets.Add("Sheet2");
        workbook.Worksheets.Add("Sheet3");

        // Collect TabId and sheet name for each worksheet
        var sheetInfo = new List<object>();
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            sheetInfo.Add(new
            {
                SheetName = sheet.Name,
                TabId = sheet.TabId
            });
        }

        // Serialize the collection to a formatted JSON string
        string json = JsonSerializer.Serialize(sheetInfo, new JsonSerializerOptions
        {
            WriteIndented = true
        });

        // Save the JSON configuration file
        string jsonPath = "WorksheetTabIds.json";
        File.WriteAllText(jsonPath, json);

        Console.WriteLine($"Worksheet TabId data exported to: {jsonPath}");
    }
}
