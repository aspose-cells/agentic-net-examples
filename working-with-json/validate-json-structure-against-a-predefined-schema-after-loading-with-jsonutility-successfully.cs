using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

class JsonSchemaValidationDemo
{
    static void Main()
    {
        // Sample JSON data to be imported
        string jsonData = @"{
            ""Products"": [
                { ""ID"": 101, ""Name"": ""Product A"", ""Price"": 99.99 }
            ]
        }";

        // Predefined JSON schema for validation
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

        // Create a new workbook instance
        Workbook workbook = new Workbook();

        // Configure layout options for JSON import
        JsonLayoutOptions layoutOptions = new JsonLayoutOptions
        {
            // Keep the original schema after import (optional but useful for later validation)
            KeptSchema = true
        };

        // Import JSON data into the first worksheet starting at cell A1 (row 0, column 0)
        JsonUtility.ImportData(jsonData, workbook.Worksheets[0].Cells, 0, 0, layoutOptions);

        // Prepare JSON save options and attach the schema for validation
        JsonSaveOptions saveOptions = new JsonSaveOptions
        {
            Schemas = new string[] { schema },   // Assign the predefined schema
            ExportNestedStructure = true,        // Preserve nested JSON structure
            SkipEmptyRows = true                 // Omit empty rows in the output
        };

        // Save the workbook to a JSON file; Aspose.Cells will validate against the schema
        string outputPath = "validated_output.json";
        workbook.Save(outputPath, saveOptions);

        Console.WriteLine("JSON saved and validated successfully.");
    }
}