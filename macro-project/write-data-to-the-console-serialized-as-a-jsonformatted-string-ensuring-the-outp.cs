using System;
using Aspose.Cells;
using Aspose.Cells.Utility;
using AsposeRange = Aspose.Cells.Range;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data with a header row
        worksheet.Cells["A1"].PutValue("Name");
        worksheet.Cells["B1"].PutValue("Age");
        worksheet.Cells["A2"].PutValue("John");
        worksheet.Cells["B2"].PutValue(30);
        worksheet.Cells["A3"].PutValue("Alice");
        worksheet.Cells["B3"].PutValue(25);

        // Define the range to be exported (including header)
        AsposeRange range = worksheet.Cells.CreateRange("A1:B3");

        // Configure JSON export options
        JsonSaveOptions jsonOptions = new JsonSaveOptions
        {
            HasHeaderRow = true,          // Use first row as header
            Indent = "    "               // 4‑space indentation for pretty JSON
        };

        // Export the range to a JSON string
        string jsonOutput = JsonUtility.ExportRangeToJson(range, jsonOptions);

        // Write the JSON string to the console
        Console.WriteLine(jsonOutput);
    }
}