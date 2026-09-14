// Title: Export worksheet names and TabId values to a JSON file using Aspose.Cells for .NET
// AI Prompts: Write a C# program that loads an Excel workbook with Aspose.Cells, extracts each worksheet's Name and TabId, and saves the collection as a formatted JSON file. | Create a .NET snippet that iterates through Workbook.Worksheets, builds a list of objects containing Name and TabId, and uses System.Text.Json to write the data to a configuration file.
// Common Searches: how to get worksheet TabId with Aspose.Cells in C# | export Excel sheet names and TabId to JSON using .NET | serialize Aspose.Cells worksheet metadata to a JSON configuration file | C# code to write worksheet identifiers from a workbook to JSON
// Tags: aspocells export worksheet tabid to json | c# serialize worksheet metadata with system.text.json | aspocells retrieve worksheet identifiers | json configuration file from excel workbook c# | aspocells worksheet tabid extraction

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using System.Text.Json;

// // Loads an Excel workbook via Aspose.Cells, collects each worksheet's Name and TabId into a list, serializes the list to indented JSON using System.Text.Json, and writes the result to a specified file.
class Program
{
    static void Main()
    {
        // Path to the source Excel workbook
        string excelPath = "input.xlsx";

        // Path where the JSON configuration will be saved
        string jsonPath = "worksheet_tabids.json";

        // Load the workbook from the file system
        Workbook workbook = new Workbook(excelPath);

        // Collection to hold each worksheet's name and TabId
        var worksheetsInfo = new List<Dictionary<string, object>>();

        // Iterate through all worksheets in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Capture the worksheet name and its TabId
            var info = new Dictionary<string, object>
            {
                { "Name", sheet.Name },
                { "TabId", sheet.TabId }
            };

            worksheetsInfo.Add(info);
        }

        // Serialize the collection to a formatted JSON string
        var jsonOptions = new JsonSerializerOptions { WriteIndented = true };
        string json = JsonSerializer.Serialize(worksheetsInfo, jsonOptions);

        // Write the JSON string to the output file
        File.WriteAllText(jsonPath, json);
    }
}
