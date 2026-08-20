// Title: Flatten Nested JSON to a Compact File with Aspose.Cells for .NET (C#)
// Description: This C# console app checks for a file named nested_input.json, creates a sample if it does not exist, loads the JSON into an Aspose.Cells Workbook using JsonLoadOptions, and then saves it as flattened_output.json with JsonSaveOptions configured to remove hierarchy, output a single JSON object, and skip empty rows, resulting in a flat, lightweight JSON document.
// Keywords: Aspose.Cells | C# JSON flatten | JsonLoadOptions | JsonSaveOptions | ExportNestedStructure false | compact JSON output | flatten hierarchy Aspose | Workbook.Save JSON | nested JSON conversion .NET | skip empty rows JSON
// Common Searches: C# flatten nested JSON Aspose.Cells | Aspose.Cells JsonSaveOptions ExportNestedStructure example | convert hierarchical JSON to flat JSON .NET | save JSON without nesting using Aspose.Cells | compact JSON file from workbook C#
// Developer Intent: Read a hierarchical JSON file and produce a flat JSON representation using Aspose.Cells in a .NET application.
// Use Cases: Prepare employee or product data for APIs that require flat JSON structures. | Create lightweight JSON reports from Excel‑like workbooks while removing empty rows. | Integrate Aspose.Cells into ETL pipelines to transform nested JSON payloads into flat files for downstream systems.
// AI Prompts: Show a C# snippet that loads a nested JSON into an Aspose.Cells Workbook and saves it as flat JSON with appropriate JsonSaveOptions. | Explain the effect of setting ExportNestedStructure to false on the resulting JSON and how to adjust other options for custom flattening. | Provide performance recommendations for processing large nested JSON files with Aspose.Cells, including memory usage and streaming tips.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsJsonFlattenDemo
{
    // This C# console app checks for a file named nested_input.json, creates a sample if it does not exist, loads the JSON into an Aspose.Cells Workbook using JsonLoadOptions, and then saves it as flattened_output.json with JsonSaveOptions configured to remove hierarchy, output a single JSON object, and skip empty rows, resulting in a flat, lightweight JSON document.
    class Program
    {
        static void Main()
        {
            // Paths for input and output JSON files
            string inputJsonPath = "nested_input.json";
            string outputJsonPath = "flattened_output.json";

            try
            {
                // Ensure the input JSON file exists; create a sample if missing
                if (!File.Exists(inputJsonPath))
                {
                    string sampleJson = @"{
  ""Employee"": {
    ""Name"": ""John Doe"",
    ""Address"": {
      ""Street"": ""123 Main St"",
      ""City"": ""Anytown"",
      ""Zip"": ""12345""
    },
    ""Projects"": [
      { ""Id"": 1, ""Title"": ""Project A"" },
      { ""Id"": 2, ""Title"": ""Project B"" }
    ]
  }
}";
                    File.WriteAllText(inputJsonPath, sampleJson);
                    Console.WriteLine($"Sample input JSON created at: {Path.GetFullPath(inputJsonPath)}");
                }

                // Load the JSON file into a workbook with optional load settings
                JsonLoadOptions loadOptions = new JsonLoadOptions
                {
                    KeptSchema = true // keep original schema (optional)
                };

                Workbook workbook = new Workbook(inputJsonPath, loadOptions);

                // Configure JSON save options to flatten the hierarchy
                JsonSaveOptions saveOptions = new JsonSaveOptions
                {
                    ExportNestedStructure = false,      // flatten nested objects
                    AlwaysExportAsJsonObject = true,    // output as JSON object even for single sheet
                    SkipEmptyRows = true                // omit empty rows for compact output
                };

                // Save the flattened JSON
                workbook.Save(outputJsonPath, saveOptions);
                Console.WriteLine($"Flattened JSON saved to: {Path.GetFullPath(outputJsonPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
