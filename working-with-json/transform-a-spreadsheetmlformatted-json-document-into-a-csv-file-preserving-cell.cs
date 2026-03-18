using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

class JsonToCsvConverter
{
    static void Main()
    {
        // JSON string representing a spreadsheet (replace with actual JSON content)
        string json = @"[
            { ""Name"": ""John"", ""Age"": 30, ""City"": ""New York"" },
            { ""Name"": ""Alice"", ""Age"": 25, ""City"": ""London"" },
            { ""Name"": ""Bob"", ""Age"": 35, ""City"": ""Paris"" }
        ]";

        // Destination CSV file path
        string csvPath = "output.csv";

        // Create a new workbook (empty spreadsheet)
        Workbook workbook = new Workbook();

        // Get the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Configure JSON import options: treat JSON array as a table
        JsonLayoutOptions importOptions = new JsonLayoutOptions
        {
            ArrayAsTable = true
        };

        // Import JSON data into the worksheet starting at cell A1 (row 0, column 0)
        JsonUtility.ImportData(json, worksheet.Cells, 0, 0, importOptions);

        // Save the workbook as CSV, preserving the cell data and layout
        workbook.Save(csvPath, SaveFormat.Csv);

        Console.WriteLine($"JSON data has been converted to CSV at: {csvPath}");
    }
}