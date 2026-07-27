using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Cells;
using Aspose.Cells.Properties;

class ExportCustomPropertiesToJson
{
    static void Main()
    {
        // Path to the existing workbook
        string workbookPath = "input.xlsx";

        // Load the workbook from file
        Workbook workbook = new Workbook(workbookPath);

        // Prepare a list to hold property information
        var properties = new List<object>();

        // Enumerate custom document properties
        foreach (DocumentProperty prop in workbook.CustomDocumentProperties)
        {
            // Capture name, type, and value of each property
            properties.Add(new
            {
                Name = prop.Name,
                Type = prop.Type.ToString(),
                Value = prop.Value
            });
        }

        // Serialize the list to JSON with indentation
        string json = JsonSerializer.Serialize(properties, new JsonSerializerOptions { WriteIndented = true });

        // Write JSON to a file
        string jsonPath = "customProperties.json";
        File.WriteAllText(jsonPath, json);

        Console.WriteLine($"Custom properties exported to {jsonPath}");
    }
}