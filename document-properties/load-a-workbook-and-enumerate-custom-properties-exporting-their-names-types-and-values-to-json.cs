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
        // Path to the Excel workbook to be processed
        string workbookPath = "input.xlsx";

        // Load the workbook from the specified file
        Workbook workbook = new Workbook(workbookPath);

        // Prepare a list to hold property information
        var properties = new List<object>();

        // Enumerate all custom document properties
        foreach (DocumentProperty prop in workbook.CustomDocumentProperties)
        {
            // Capture name, type (as string), and value of each property
            properties.Add(new
            {
                Name = prop.Name,
                Type = prop.Type.ToString(),
                Value = prop.Value
            });
        }

        // Serialize the list to JSON with indentation for readability
        string json = JsonSerializer.Serialize(properties, new JsonSerializerOptions { WriteIndented = true });

        // Write the JSON output to a file
        string outputPath = "customProperties.json";
        File.WriteAllText(outputPath, json);

        Console.WriteLine($"Custom properties exported to {outputPath}");
    }
}