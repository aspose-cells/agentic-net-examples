using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Cells;
using Aspose.Cells.Properties;

class Program
{
    static void Main()
    {
        // Load the workbook from a file
        string workbookPath = "input.xlsx";
        Workbook workbook = new Workbook(workbookPath);

        // Prepare a list to hold custom property details
        var customProps = new List<CustomPropertyInfo>();

        // Enumerate all custom document properties
        foreach (DocumentProperty prop in workbook.CustomDocumentProperties)
        {
            customProps.Add(new CustomPropertyInfo
            {
                Name = prop.Name,
                Type = prop.Type.ToString(),
                Value = prop.Value?.ToString()
            });
        }

        // Serialize the list to formatted JSON
        string json = JsonSerializer.Serialize(customProps, new JsonSerializerOptions { WriteIndented = true });

        // Write the JSON to a file
        File.WriteAllText("customProperties.json", json);
    }

    // Helper class representing a custom property for JSON output
    class CustomPropertyInfo
    {
        public string Name { get; set; }
        public string Type { get; set; }
        public string Value { get; set; }
    }
}