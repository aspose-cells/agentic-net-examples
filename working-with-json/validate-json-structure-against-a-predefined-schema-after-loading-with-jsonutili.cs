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

        // Create a new workbook instance
        Workbook workbook = new Workbook();

        // Configure layout options to keep the schema during import
        JsonLayoutOptions layoutOptions = new JsonLayoutOptions();
        layoutOptions.KeptSchema = true;

        // Import the JSON string into the first worksheet starting at cell A1
        JsonUtility.ImportData(jsonData, workbook.Worksheets[0].Cells, 0, 0, layoutOptions);

        // Prepare JSON save options and attach the schema for validation
        JsonSaveOptions saveOptions = new JsonSaveOptions();
        saveOptions.Schemas = new string[] { schema };
        saveOptions.ExportNestedStructure = true;
        saveOptions.SkipEmptyRows = true;

        // Attempt to save the workbook as JSON; validation occurs during save
        try
        {
            workbook.Save("validated_output.json", saveOptions);
            Console.WriteLine("JSON saved and validated successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Validation failed: " + ex.Message);
        }
    }
}