// Title: Read custom document properties from an Excel workbook and output them as formatted JSON using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that opens an .xlsx file with Aspose.Cells, loops through workbook.CustomDocumentProperties, and builds a JSON array of each property's Name, Type, and Value with indentation. | Create a .NET snippet that verifies the Excel file exists, extracts all custom metadata from the workbook via Aspose.Cells, and serializes the collection to pretty‑printed JSON using System.Text.Json.
// Common Searches: Aspose.Cells C# retrieve custom document properties from an Excel file | how to export Excel custom properties to JSON using Aspose.Cells | list custom metadata in .xlsx with Aspose.Cells .NET | C# code to serialize workbook custom properties to indented JSON | extract custom properties from Excel workbook via Aspose.Cells API
// Tags: Aspose.Cells read custom properties | C# serialize Excel custom properties to JSON | workbook.CustomDocumentProperties loop | export Excel custom metadata as formatted JSON | Aspose.Cells .NET extract workbook properties

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Cells;

// The example checks for the presence of an input.xlsx file, loads it with Aspose.Cells, iterates over workbook.CustomDocumentProperties, gathers each property's Name, Type, and Value into a list, serializes the list to indented JSON using System.Text.Json, and writes the JSON to the console.
class Program
{
    static void Main()
    {
        try
        {
            // Path to the input workbook
            string workbookPath = "input.xlsx";

            // Verify that the file exists to avoid FileNotFoundException
            if (!File.Exists(workbookPath))
            {
                Console.WriteLine($"Error: The file \"{workbookPath}\" was not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(workbookPath);

            // List to hold custom property information
            var customProperties = new List<Dictionary<string, object>>();

            // Iterate through custom document properties
            for (int i = 0; i < workbook.CustomDocumentProperties.Count; i++)
            {
                // Use dynamic to avoid compile‑time dependency on the exact type name
                dynamic prop = workbook.CustomDocumentProperties[i];

                var propInfo = new Dictionary<string, object>
                {
                    { "Name", prop.Name },
                    { "Type", prop.Type.ToString() },
                    { "Value", prop.Value }
                };
                customProperties.Add(propInfo);
            }

            // Serialize the list to formatted JSON
            string json = JsonSerializer.Serialize(customProperties, new JsonSerializerOptions { WriteIndented = true });

            // Output JSON
            Console.WriteLine(json);
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
