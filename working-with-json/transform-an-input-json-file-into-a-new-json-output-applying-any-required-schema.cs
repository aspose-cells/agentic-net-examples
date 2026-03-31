using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsJsonTransform
{
    class Program
    {
        static void Main()
        {
            // Paths for input JSON and output JSON files
            string inputJsonPath = "input.json";
            string outputJsonPath = "output_transformed.json";

            // ------------------- Load JSON into Workbook -------------------
            // Create load options and enable keeping the original schema
            JsonLoadOptions loadOptions = new JsonLoadOptions
            {
                KeptSchema = true,
                // Optional: configure layout options if needed
                LayoutOptions = new JsonLayoutOptions
                {
                    ArrayAsTable = true,
                    ConvertNumericOrDate = true,
                    NumberFormat = "0.00",
                    DateFormat = "yyyy-MM-dd"
                }
            };

            // Load the JSON file into a workbook using the load options
            Workbook workbook = new Workbook(inputJsonPath, loadOptions);

            // ------------------- (Optional) Modify Workbook -------------------
            // Example: add a new column with a computed value
            Worksheet sheet = workbook.Worksheets[0];
            // Assuming the first row contains headers, add a new header
            sheet.Cells[0, sheet.Cells.MaxColumn + 1].PutValue("ProcessedDate");
            // Fill the new column with the current date for each data row
            for (int row = 1; row <= sheet.Cells.MaxDataRow; row++)
            {
                sheet.Cells[row, sheet.Cells.MaxColumn].PutValue(DateTime.Now);
            }

            // ------------------- Prepare JSON Save Options -------------------
            // Sample JSON schema to be attached to the output
            string sampleSchema = @"{
                ""$schema"": ""http://json-schema.org/draft-07/schema#"",
                ""type"": ""object"",
                ""properties"": {
                    ""Data"": {
                        ""type"": ""array"",
                        ""items"": {
                            ""type"": ""object"",
                            ""properties"": {
                                ""Name"": { ""type"": ""string"" },
                                ""Age"": { ""type"": ""integer"" },
                                ""ProcessedDate"": { ""type"": ""string"", ""format"": ""date-time"" }
                            },
                            ""required"": [""Name"", ""Age"", ""ProcessedDate""]
                        }
                    }
                },
                ""required"": [""Data""]
            }";

            // Configure save options
            JsonSaveOptions saveOptions = new JsonSaveOptions
            {
                // Attach the schema
                Schemas = new string[] { sampleSchema },

                // Export as nested (parent‑child) JSON structure
                ExportNestedStructure = true,

                // Skip rows that are completely empty
                SkipEmptyRows = true,

                // Indent the output for readability (2 spaces)
                Indent = "  ",

                // Indicate that the first row contains headers
                HasHeaderRow = true
            };

            // ------------------- Save Workbook as JSON -------------------
            workbook.Save(outputJsonPath, saveOptions);

            Console.WriteLine($"JSON file saved to: {outputJsonPath}");
        }
    }
}