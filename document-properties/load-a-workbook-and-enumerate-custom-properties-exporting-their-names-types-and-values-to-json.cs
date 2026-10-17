// Title: Read Excel custom document properties and export them to JSON using Aspose.Cells for .NET (C#)
// Common Searches: how to read custom document properties from an Excel workbook with Aspose.Cells in C# | export Excel custom properties to a JSON file using Aspose.Cells .NET | list all custom workbook properties and serialize to JSON in C# | Aspose.Cells example for enumerating custom document properties and saving as JSON | C# code to convert Excel custom document properties to JSON with Aspose.Cells
// Tags: Aspose.Cells read custom document properties | export Excel custom properties to JSON C# | enumerate workbook custom properties Aspose.Cells | serialize custom document properties JSON .NET | C# Aspose.Cells custom properties extraction

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Cells;

namespace AsposeCellsCustomPropertiesExport
{
    // Model representing a custom property for JSON serialization
    // The sample loads an Excel workbook with Aspose.Cells, iterates through its CustomDocumentProperties collection, captures each property's name, type, and value into a DTO, and then serializes the DTO list to indented JSON using System.Text.Json, outputting the JSON to the console (or optionally to a file).
    public class CustomPropertyInfo
    {
        public string Name { get; set; } = string.Empty;
        public string Type { get; set; } = string.Empty;
        public object Value { get; set; } = new object();
    }

    class Program
    {
        static void Main(string[] args)
        {
            // Path to the Excel workbook (adjust as needed)
            string workbookPath = @"C:\Path\To\YourWorkbook.xlsx";

            // Verify that the workbook file exists to avoid FileNotFoundException
            if (!File.Exists(workbookPath))
            {
                Console.WriteLine($"Error: Workbook file not found at '{workbookPath}'.");
                return;
            }

            try
            {
                // Load the workbook using Aspose.Cells
                Workbook workbook = new Workbook(workbookPath);

                // Collection to hold exported property information
                List<CustomPropertyInfo> exportedProperties = new List<CustomPropertyInfo>();

                // Iterate through all custom document properties
                foreach (var prop in workbook.CustomDocumentProperties)
                {
                    // Create a DTO for each property
                    var info = new CustomPropertyInfo
                    {
                        Name = prop.Name,
                        // Convert the enum value to its string representation
                        Type = prop.Type.ToString(),
                        // Use the raw value; JsonSerializer will handle appropriate conversion
                        Value = prop.Value
                    };

                    exportedProperties.Add(info);
                }

                // Serialize the list to JSON (indented for readability)
                string jsonOutput = JsonSerializer.Serialize(exportedProperties, new JsonSerializerOptions
                {
                    WriteIndented = true
                });

                // Output JSON to console (or write to a file if desired)
                Console.WriteLine(jsonOutput);

                // Optional: write JSON to a file
                // string jsonFilePath = @"C:\Path\To\CustomProperties.json";
                // File.WriteAllText(jsonFilePath, jsonOutput);
            }
            catch (Exception ex)
            {
                // Catch any runtime exceptions and display a friendly message
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
