using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

class IncludeColumnNamesDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Add header row (column names)
        cells["A1"].PutValue("Name");
        cells["B1"].PutValue("Age");

        // Add sample data rows
        cells["A2"].PutValue("John");
        cells["B2"].PutValue(30);
        cells["A3"].PutValue("Alice");
        cells["B3"].PutValue(25);

        // Define the range that includes the header and data
        Aspose.Cells.Range range = cells.CreateRange("A1:B3");

        // Configure export options to treat the first row as header,
        // which causes column names to be used as JSON keys
        ExportRangeToJsonOptions exportOptions = new ExportRangeToJsonOptions
        {
            HasHeaderRow = true
        };

        // Export the range to a JSON string
        string jsonOutput = JsonUtility.ExportRangeToJson(range, exportOptions);

        // Display the resulting JSON
        Console.WriteLine(jsonOutput);
    }
}