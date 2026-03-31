using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

class PrnJsonToXlsx
{
    static void Main()
    {
        // Path to the input PRN‑formatted JSON file
        string inputJsonPath = @"C:\Data\input.prn.json";

        // Path for the output XLSX workbook
        string outputXlsxPath = @"C:\Data\output.xlsx";

        // Read the JSON content; if the file does not exist, use a sample JSON string
        string jsonContent;
        if (File.Exists(inputJsonPath))
        {
            jsonContent = File.ReadAllText(inputJsonPath);
        }
        else
        {
            // Sample JSON representing a simple table
            jsonContent = @"{
                ""Employees"": [
                    { ""Id"": 1, ""Name"": ""John Doe"", ""Department"": ""HR"" },
                    { ""Id"": 2, ""Name"": ""Jane Smith"", ""Department"": ""IT"" },
                    { ""Id"": 3, ""Name"": ""Bob Johnson"", ""Department"": ""Finance"" }
                ]
            }";
        }

        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Configure JSON import options.
        JsonLayoutOptions importOptions = new JsonLayoutOptions
        {
            ArrayAsTable = true,
            IgnoreTitle = false
        };

        // Import JSON data starting at cell A1 (row 0, column 0)
        JsonUtility.ImportData(jsonContent, sheet.Cells, 0, 0, importOptions);

        // Save the workbook to XLSX
        workbook.Save(outputXlsxPath, SaveFormat.Xlsx);
    }
}