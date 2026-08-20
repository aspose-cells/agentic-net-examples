// Title: C# – Validate JSON with Aspose.Cells JsonUtility & JsonSaveOptions (Draft‑07 Schema)
// Description: Demonstrates how to import JSON into an Aspose.Cells workbook, retain schema metadata, and enforce Draft‑07 validation on save. The workbook throws an exception if the data does not conform to the defined schema.
// Keywords: Aspose.Cells JSON validation | JsonUtility ImportData C# | JsonSaveOptions schema enforcement | Draft‑07 JSON schema Aspose | C# workbook JSON import | Excel to JSON schema check | .NET JSON schema validation | Aspose.Cells example GitHub | US developers | EU developers
// Common Searches: Aspose.Cells validate JSON against schema .NET | JsonUtility import data with schema preservation | JsonSaveOptions Schemas property usage | C# example for JSON schema validation in Excel | How to catch Aspose.Cells JSON validation errors
// Developer Intent: Ensure imported JSON data matches a predefined Draft‑07 schema and receive an error when it does not.
// Use Cases: Load a product catalog JSON, keep its schema, and verify compliance before exporting back to JSON. | Read configuration files into a worksheet, retain schema metadata, and automatically validate on save. | Implement a data‑exchange pipeline where incoming JSON must meet a contract, using Aspose.Cells to enforce the schema and flag mismatches.
// AI Prompts: Write C# code that catches the Aspose.Cells validation exception and logs detailed error information. | Show how to configure JsonLayoutOptions to ignore extra fields while still requiring mandatory properties. | Provide an example of using multiple schemas in JsonSaveOptions for conditional validation of different JSON sections.

using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

// Demonstrates how to import JSON into an Aspose.Cells workbook, retain schema metadata, and enforce Draft‑07 validation on save. The workbook throws an exception if the data does not conform to the defined schema.
class JsonSchemaValidationDemo
{
    static void Main()
    {
        // Sample JSON data to import
        string jsonData = @"{
            ""Products"": [
                { ""ID"": 101, ""Name"": ""Product A"", ""Price"": 99.99 },
                { ""ID"": 102, ""Name"": ""Product B"", ""Price"": 149.50 }
            ]
        }";

        // JSON schema that the data must conform to
        string schema = @"{
            ""$schema"": ""http://json-schema.org/draft-07/schema#"",
            ""type"": ""object"",
            ""properties"": {
                ""Products"": {
                    ""type"": ""array"",
                    ""items"": {
                        ""type"": ""object"",
                        ""properties"": {
                            ""ID"": { ""type"": ""integer"" },
                            ""Name"": { ""type"": ""string"" },
                            ""Price"": { ""type"": ""number"" }
                        },
                        ""required"": [""ID"", ""Name"", ""Price""]
                    }
                }
            },
            ""required"": [""Products""]
        }";

        // Create a new workbook (lifecycle: create)
        Workbook workbook = new Workbook();

        // Import the JSON data into the first worksheet (lifecycle: load)
        JsonLayoutOptions layoutOptions = new JsonLayoutOptions
        {
            KeptSchema = true   // keep schema information for later validation
        };
        JsonUtility.ImportData(jsonData, workbook.Worksheets[0].Cells, 0, 0, layoutOptions);

        // Configure JSON save options with the predefined schema (validation occurs on save)
        JsonSaveOptions saveOptions = new JsonSaveOptions
        {
            Schemas = new string[] { schema },
            ExportNestedStructure = true,
            SkipEmptyRows = true
        };

        // Save the workbook as JSON; if the data does not match the schema,
        // Aspose.Cells will raise an exception during this operation.
        workbook.Save("validated_output.json", saveOptions);
    }
}
