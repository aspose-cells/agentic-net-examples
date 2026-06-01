using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

class Program
{
    static void Main()
    {
        // Sample JSON array where each object will become a row and its properties become columns
        string json = @"[
            { ""Name"": ""John"", ""Age"": 30, ""City"": ""New York"" },
            { ""Name"": ""Jane"", ""Age"": 25, ""City"": ""London"" },
            { ""Name"": ""Bob"",  ""Age"": 40, ""City"": ""Paris"" }
        ]";

        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Configure layout options to treat the JSON array as a table
        JsonLayoutOptions options = new JsonLayoutOptions
        {
            ArrayAsTable = true   // Enables automatic row creation and column mapping
        };

        // Import the JSON data starting at cell A1 (row index 0, column index 0)
        JsonUtility.ImportData(json, worksheet.Cells, 0, 0, options);

        // Save the workbook to an Excel file
        workbook.Save("JsonArrayTable.xlsx");
    }
}