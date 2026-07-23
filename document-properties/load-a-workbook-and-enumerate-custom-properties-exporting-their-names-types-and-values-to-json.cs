// Title: C# – Export Excel Custom Document Properties to JSON with Aspose.Cells
// Description: Loads an Excel workbook using Aspose.Cells, iterates the CustomDocumentProperties collection, captures each property's name, type, and value, serializes the data to indented JSON with System.Text.Json, and writes the result to a file (customProperties.json).
// Keywords: Aspose.Cells | C# | Excel | custom document properties | JSON export | Workbook | System.Text.Json | sample code | GitHub example | API usage
// Common Searches: Aspose.Cells export custom properties to JSON C# | How to read Excel custom document properties with Aspose.Cells | C# code to serialize workbook properties to JSON | Save Excel custom metadata as JSON file | Aspose.Cells example for custom document properties
// Developer Intent: Extract custom document properties from an Excel workbook and save them as a formatted JSON file.
// Use Cases: Create a machine‑readable manifest of workbook metadata for downstream processing. | Generate audit logs of custom properties across multiple Excel files for compliance. | Provide a JSON payload for web APIs that need to expose Excel custom metadata.
// AI Prompts: Generate C# code that uses Aspose.Cells to read all custom document properties from an Excel file and output them as pretty‑printed JSON. | Show how to filter custom document properties by type before exporting them to JSON with Aspose.Cells. | Explain strategies for handling non‑serializable values (e.g., dates, binary data) when converting custom properties to JSON.

using System;
using System.Collections.Generic;
using System.Text.Json;
using Aspose.Cells;
using Aspose.Cells.Properties;

// Loads an Excel workbook using Aspose.Cells, iterates the CustomDocumentProperties collection, captures each property's name, type, and value, serializes the data to indented JSON with System.Text.Json, and writes the result to a file (customProperties.json).
class Program
{
    static void Main()
    {
        // Load the workbook from a file
        string workbookPath = "input.xlsx";
        Workbook workbook = new Workbook(workbookPath);

        // Collect custom property information
        var customProps = new List<object>();
        foreach (DocumentProperty prop in workbook.CustomDocumentProperties)
        {
            customProps.Add(new
            {
                Name = prop.Name,
                Type = prop.Type.ToString(),
                Value = prop.Value
            });
        }

        // Convert the collection to formatted JSON
        var jsonOptions = new JsonSerializerOptions { WriteIndented = true };
        string json = JsonSerializer.Serialize(customProps, jsonOptions);

        // Write JSON to a file
        System.IO.File.WriteAllText("customProperties.json", json);

        Console.WriteLine("Custom properties have been exported to customProperties.json");
    }
}
