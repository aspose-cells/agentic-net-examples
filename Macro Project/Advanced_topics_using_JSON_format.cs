using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Utility;

class JsonAdvancedDemo
{
    static void Main()
    {
        // 1. Create a new workbook
        Workbook workbook = new Workbook();

        // 2. Sample JSON data representing employees
        string jsonData = @"{
            ""Employees"": [
                { ""Id"": 1, ""Name"": ""Alice"", ""JoinDate"": ""2023-01-15"", ""Salary"": 70000 },
                { ""Id"": 2, ""Name"": ""Bob"", ""JoinDate"": ""2023-03-20"", ""Salary"": 85000 }
            ]
        }";

        // 3. Set import options: treat arrays as tables, apply a custom title style
        JsonLayoutOptions importOptions = new JsonLayoutOptions
        {
            ArrayAsTable = true,
            TitleStyle = CreateTitleStyle(workbook)
        };

        // 4. Import JSON into the first worksheet starting at cell A1
        JsonUtility.ImportData(jsonData, workbook.Worksheets[0].Cells, 0, 0, importOptions);

        // 5. Define a JSON schema for validation
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
                            ""JoinDate"": { ""type"": ""string"", ""format"": ""date"" },
                            ""Salary"": { ""type"": ""number"" }
                        },
                        ""required"": [""Id"", ""Name"", ""JoinDate"", ""Salary""]
                    }
                }
            },
            ""required"": [""Employees""]
        }";

        // 6. Configure JSON save options: nested structure, pretty indentation, attach schema, force object output
        JsonSaveOptions saveOptions = new JsonSaveOptions
        {
            ExportNestedStructure = true,
            Indent = "    ", // 4 spaces
            Schemas = new string[] { schema },
            AlwaysExportAsJsonObject = true
        };

        // 7. Export a specific range (A1:D3) to a JSON string
        Aspose.Cells.Range exportRange = workbook.Worksheets[0].Cells.CreateRange(0, 0, 3, 4);
        string jsonResult = JsonUtility.ExportRangeToJson(exportRange, saveOptions);
        Console.WriteLine("Exported JSON from range:");
        Console.WriteLine(jsonResult);

        // 8. Save the entire workbook as a JSON file using the same options
        string outputJsonPath = "EmployeesOutput.json";
        workbook.Save(outputJsonPath, saveOptions);
        Console.WriteLine($"Workbook saved to JSON file: {outputJsonPath}");

        // 9. Load the JSON file back into a workbook
        Workbook loadedWorkbook = new Workbook(outputJsonPath);

        // 10. Save the loaded workbook to Excel to verify round‑trip conversion
        string roundTripExcel = "RoundTripOutput.xlsx";
        loadedWorkbook.Save(roundTripExcel);
        Console.WriteLine($"Loaded JSON saved back to Excel: {roundTripExcel}");
    }

    // Helper method to create a style for column titles
    private static Style CreateTitleStyle(Workbook wb)
    {
        Style style = wb.CreateStyle();
        style.HorizontalAlignment = TextAlignmentType.Center;
        style.Font.IsBold = true;
        style.Font.Color = Color.DarkBlue;
        return style;
    }
}