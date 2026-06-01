using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

class JsonMultipleWorksheetsDemo
{
    static void Main()
    {
        // Sample JSON containing multiple tables (arrays)
        string json = @"{
            ""Employees"": [
                { ""Name"": ""John"", ""Age"": 30 },
                { ""Name"": ""Jane"", ""Age"": 25 }
            ],
            ""Departments"": [
                { ""DeptId"": 1, ""DeptName"": ""HR"" },
                { ""DeptId"": 2, ""DeptName"": ""IT"" }
            ]
        }";

        // Write JSON to a temporary file (required for Workbook constructor)
        string jsonPath = "data.json";
        File.WriteAllText(jsonPath, json);

        // Configure JSON load options:
        // - MultipleWorksheets = true creates a separate worksheet for each top‑level array
        // - LayoutOptions.ArrayAsTable = true treats each array as a table
        JsonLoadOptions loadOptions = new JsonLoadOptions
        {
            MultipleWorksheets = true,
            LayoutOptions = new JsonLayoutOptions
            {
                ArrayAsTable = true
            }
        };

        // Load the JSON file into a workbook using the specified options
        Workbook workbook = new Workbook(jsonPath, loadOptions);

        // Save the workbook to an Excel file
        workbook.Save("output.xlsx");
    }
}