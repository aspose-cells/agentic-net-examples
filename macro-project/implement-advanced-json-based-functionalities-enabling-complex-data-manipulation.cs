using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

class AdvancedJsonDemo
{
    static void Main()
    {
        // Sample JSON data representing a collection of employees
        string jsonData = @"{
            ""Employees"": [
                { ""Id"": 1, ""Name"": ""Alice"", ""Salary"": 50000, ""HireDate"": ""2022-01-15"" },
                { ""Id"": 2, ""Name"": ""Bob"", ""Salary"": 60000, ""HireDate"": ""2021-07-01"" }
            ]
        }";

        // Configure JsonLoadOptions with detailed layout options
        JsonLoadOptions loadOptions = new JsonLoadOptions
        {
            StartCell = "A1",               // Begin import at cell A1
            MultipleWorksheets = false,     // Keep data in a single worksheet
            KeptSchema = true,              // Preserve schema information
            LayoutOptions = new JsonLayoutOptions
            {
                ArrayAsTable = true,                // Treat arrays as tables
                ConvertNumericOrDate = true,        // Auto‑convert numbers and dates
                NumberFormat = "#,##0.00",          // Format for numeric values
                DateFormat = "yyyy-MM-dd",          // Format for date values
                IgnoreNull = true,                  // Skip null values
                IgnoreTitle = false                 // Keep titles
            }
        };

        // Create a new workbook and import the JSON data using the layout options
        Workbook workbook = new Workbook();
        Worksheet ws = workbook.Worksheets[0];
        JsonUtility.ImportData(jsonData, ws.Cells, 0, 0, loadOptions.LayoutOptions);

        // Add a new employee row programmatically
        int newRow = ws.Cells.MaxDataRow + 1;
        ws.Cells[newRow, 0].PutValue(3);                                 // Id
        ws.Cells[newRow, 1].PutValue("Charlie");                        // Name
        ws.Cells[newRow, 2].PutValue(55000);                            // Salary
        ws.Cells[newRow, 3].PutValue(DateTime.Parse("2023-03-10"));    // HireDate

        // Define a JSON schema to be attached to the saved JSON file
        string schema = @"{
            ""$schema"": ""http://json-schema.org/draft-07/schema#"",
            ""type"": ""object"",
            ""properties"": {
                ""Employees"": {
                    ""type"": ""array"",
                    ""items"": {
                        ""type"": ""object"",
                        ""properties"": {
                            ""Id"": { ""type"": ""integer"" },
                            ""Name"": { ""type"": ""string"" },
                            ""Salary"": { ""type"": ""number"" },
                            ""HireDate"": { ""type"": ""string"", ""format"": ""date"" }
                        },
                        ""required"": [""Id"", ""Name"", ""Salary"", ""HireDate""]
                    }
                }
            },
            ""required"": [""Employees""]
        }";

        // Configure JsonSaveOptions to export with nested structure and embed the schema
        JsonSaveOptions saveOptions = new JsonSaveOptions
        {
            ExportNestedStructure = true,   // Preserve parent‑child hierarchy
            Schemas = new string[] { schema },
            SkipEmptyRows = true,
            ExportAsString = false,
            HasHeaderRow = true,
            Indent = "  "
        };

        // Save the workbook as an Excel file for visual verification
        workbook.Save("AdvancedJsonDemo.xlsx");

        // Save the workbook as a JSON file using the configured options
        workbook.Save("AdvancedJsonDemo.json", saveOptions);
    }
}