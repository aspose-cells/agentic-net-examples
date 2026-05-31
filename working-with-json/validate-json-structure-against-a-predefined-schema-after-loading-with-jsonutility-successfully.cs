using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

class JsonSchemaValidationDemo
{
    static void Main()
    {
        // Sample JSON data to import
        string jsonData = @"{
            ""Products"": [
                {
                    ""ID"": 101,
                    ""Name"": ""Product A"",
                    ""Price"": 99.99
                }
            ]
        }";

        // JSON schema to validate against
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

        // Create a new workbook
        Workbook workbook = new Workbook();

        // Configure layout options (keep schema for later validation)
        JsonLayoutOptions layoutOptions = new JsonLayoutOptions();
        layoutOptions.KeptSchema = true;

        // Import JSON data into the first worksheet
        JsonUtility.ImportData(jsonData, workbook.Worksheets[0].Cells, 0, 0, layoutOptions);

        // Prepare JSON save options with the predefined schema
        JsonSaveOptions saveOptions = new JsonSaveOptions();
        saveOptions.Schemas = new string[] { schema };
        saveOptions.ExportNestedStructure = true;
        saveOptions.SkipEmptyRows = true;

        // Save workbook to JSON; validation occurs during save
        try
        {
            workbook.Save("validated_output.json", saveOptions);
            Console.WriteLine("JSON validated against schema and saved successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("JSON validation failed: " + ex.Message);
        }
    }
}