using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace JsonImportExample
{
    class Program
    {
        static void Main()
        {
            // Sample JSON representing an array of objects.
            string json = @"[
                { ""ProductId"": 101, ""ProductName"": ""Laptop"", ""Price"": 999.99, ""InStock"": true },
                { ""ProductId"": 102, ""ProductName"": ""Smartphone"", ""Price"": 599.49, ""InStock"": false },
                { ""ProductId"": 103, ""ProductName"": ""Tablet"", ""Price"": 299.00, ""InStock"": true }
            ]";

            // Create a new workbook and get the first worksheet.
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Configure layout options: treat JSON arrays as Excel tables.
            JsonLayoutOptions options = new JsonLayoutOptions
            {
                ArrayAsTable = true   // Each array element becomes a row in a table.
            };

            // Import the JSON data starting at cell A1 (row 0, column 0).
            JsonUtility.ImportData(json, sheet.Cells, 0, 0, options);

            // Save the workbook to an XLSX file.
            workbook.Save("ImportedJsonData.xlsx");
        }
    }
}