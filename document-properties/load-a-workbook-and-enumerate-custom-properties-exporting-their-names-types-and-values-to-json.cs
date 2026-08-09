// Title: Export Excel Custom Document Properties to JSON with Aspose.Cells for .NET
// Description: Load an Excel workbook using Aspose.Cells, iterate through its CustomDocumentProperties collection, capture each property's name, type, and value, serialize the data to indented JSON, and write the output to a file. Ideal for extracting workbook metadata in C#.
// Keywords: Aspose.Cells | C# | Excel custom document properties | export to JSON | Workbook metadata extraction | serialize custom properties | Aspose.Cells .NET example
// Common Searches: how to read custom document properties from Excel using Aspose.Cells C# | export Aspose.Cells custom properties to JSON file | list Excel custom properties with Aspose.Cells .NET | convert workbook custom properties to JSON | Aspose.Cells enumerate custom document properties
// Developer Intent: Retrieve all custom document properties from an Excel workbook and save them as a JSON file.
// Use Cases: Generate a machine‑readable inventory of workbook metadata for audits or reporting. | Transfer Excel custom properties to external systems via a JSON payload. | Validate presence and data type of specific custom properties during automated processing.
// AI Prompts: Create C# code that uses Aspose.Cells to read every custom document property from a Workbook and returns a formatted JSON string with name, type, and value. | Provide a reusable method that accepts a Workbook object and outputs a List of objects (Name, Type, Value), handling null values and serializing to indented JSON. | Write sample code that saves extracted custom properties to a JSON file, logs the file path, and includes error handling for missing workbook or write permissions.

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Cells;
using Aspose.Cells.Properties;

// Load an Excel workbook using Aspose.Cells, iterate through its CustomDocumentProperties collection, capture each property's name, type, and value, serialize the data to indented JSON, and write the output to a file. Ideal for extracting workbook metadata in C#.
class Program
{
    static void Main()
    {
        // Load the workbook from a file (replace with your actual file path)
        string workbookPath = "input.xlsx";
        Workbook workbook = new Workbook(workbookPath);

        // Prepare a list to hold property information
        var properties = new List<PropertyInfo>();

        // Enumerate custom document properties of the workbook
        foreach (DocumentProperty prop in workbook.CustomDocumentProperties)
        {
            properties.Add(new PropertyInfo
            {
                Name = prop.Name,
                // The Type property returns a DocumentPropertyType enum; convert to string for readability
                Type = prop.Type.ToString(),
                // Value may be null, handle accordingly
                Value = prop.Value?.ToString()
            });
        }

        // Serialize the list to JSON with indentation
        var jsonOptions = new JsonSerializerOptions { WriteIndented = true };
        string json = JsonSerializer.Serialize(properties, jsonOptions);

        // Write JSON to a file
        string outputPath = "customProperties.json";
        File.WriteAllText(outputPath, json);

        Console.WriteLine($"Custom properties exported to {outputPath}");
    }

    // Helper class to represent a property in JSON
    private class PropertyInfo
    {
        public string Name { get; set; }
        public string Type { get; set; }
        public string Value { get; set; }
    }
}
