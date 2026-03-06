using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data
        sheet.Cells["A1"].PutValue("Name");
        sheet.Cells["B1"].PutValue("Age");
        sheet.Cells["A2"].PutValue("John");
        sheet.Cells["B2"].PutValue(30);
        sheet.Cells["A3"].PutValue("Jane");
        sheet.Cells["B3"].PutValue(25);

        // Define the range to be exported
        Aspose.Cells.Range range = sheet.Cells.CreateRange("A1:B3");

        // Configure JSON export options
        JsonSaveOptions options = new JsonSaveOptions
        {
            HasHeaderRow = true,
            Indent = "    "
        };

        // Export the range to a JSON string
        string jsonOutput = JsonUtility.ExportRangeToJson(range, options);

        // Write the JSON string to the console
        Console.WriteLine(jsonOutput);
    }
}