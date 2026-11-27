// Title: Export Excel custom document properties to JSON using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells export custom document properties from an .xlsx to a JSON file | How to read custom properties of an Excel workbook with Aspose.Cells and serialize them | Aspose.Cells .NET list custom document properties and save as JSON | Sample code for extracting Excel custom properties to JSON in C# | Convert Excel custom document properties to JSON using Aspose.Cells library
// Tags: Aspose.Cells custom document properties to JSON | C# enumerate Excel custom properties | serialize Excel custom properties with System.Text.Json | extract custom properties from .xlsx using Aspose.Cells | JSON file generation for Excel metadata C#

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Cells;

namespace AsposeCellsCustomPropertiesExport
{
    // DTO for JSON serialization of custom document properties
    // The example loads an Excel workbook with Aspose.Cells, iterates through its CustomDocumentProperties collection, captures each property's name, type, and value into a DTO, serializes the collection to indented JSON using System.Text.Json, and writes the output to a file, handling missing files and creating the output directory as needed.
    public class CustomPropertyInfo
    {
        public string? Name { get; set; }
        public string? Type { get; set; }
        public object? Value { get; set; }
    }

    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Path to the source workbook
                string workbookPath = "input.xlsx";

                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(workbookPath))
                {
                    Console.WriteLine($"Error: Workbook file '{workbookPath}' not found.");
                    return;
                }

                // Load the workbook using Aspose.Cells
                Workbook workbook = new Workbook(workbookPath);

                // Collection to store extracted custom property data
                List<CustomPropertyInfo> properties = new List<CustomPropertyInfo>();

                // Iterate over all custom document properties
                foreach (var prop in workbook.CustomDocumentProperties)
                {
                    // Build DTO for each property
                    var info = new CustomPropertyInfo
                    {
                        Name = prop.Name,
                        Type = prop.Type.ToString(),
                        Value = prop.Value
                    };

                    properties.Add(info);
                }

                // Serialize the list to formatted JSON
                string json = JsonSerializer.Serialize(
                    properties,
                    new JsonSerializerOptions { WriteIndented = true });

                // Output JSON to a file
                string outputPath = "customProperties.json";

                // Ensure the directory for the output file exists
                string? outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                File.WriteAllText(outputPath, json);

                Console.WriteLine($"Custom properties exported to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
