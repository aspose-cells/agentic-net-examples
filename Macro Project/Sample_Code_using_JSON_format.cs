using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsJsonSample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Populate sample data
            Worksheet ws = workbook.Worksheets[0];
            ws.Cells["A1"].PutValue("Name");
            ws.Cells["B1"].PutValue("Age");
            ws.Cells["A2"].PutValue("John");
            ws.Cells["B2"].PutValue(30);
            ws.Cells["A3"].PutValue("Jane");
            ws.Cells["B3"].PutValue(25);

            // Configure JSON save options
            JsonSaveOptions jsonOptions = new JsonSaveOptions
            {
                Indent = "    ",                 // 4 spaces indentation
                ExportNestedStructure = true,   // Export as parent‑child hierarchy
                AlwaysExportAsJsonObject = true,// Force output as JSON object
                ExportStylePool = false,        // Export styles individually per cell
                SkipEmptyRows = true,           // Omit empty rows
                HasHeaderRow = true             // First row contains headers
            };

            // Optional: add a JSON schema for validation
            string schema = @"{
                ""$schema"": ""http://json-schema.org/draft-07/schema#"",
                ""type"": ""object"",
                ""properties"": {
                    ""Name"": { ""type"": ""string"" },
                    ""Age"": { ""type"": ""integer"" }
                },
                ""required"": [""Name"", ""Age""]
            }";
            jsonOptions.Schemas = new string[] { schema };

            // Save workbook as JSON using the configured options
            string outputPath = "sample_output.json";
            workbook.Save(outputPath, jsonOptions);

            Console.WriteLine($"Workbook saved to JSON file: {outputPath}");
        }
    }
}