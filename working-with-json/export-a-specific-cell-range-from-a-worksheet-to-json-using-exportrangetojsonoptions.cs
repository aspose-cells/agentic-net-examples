using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;
using AsposeRange = Aspose.Cells.Range;

class ExportRangeToJsonDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data in the worksheet
            sheet.Cells["A1"].PutValue("Name");
            sheet.Cells["B1"].PutValue("Age");
            sheet.Cells["C1"].PutValue("City");
            sheet.Cells["A2"].PutValue("John");
            sheet.Cells["B2"].PutValue(30);
            sheet.Cells["C2"].PutValue("New York");
            sheet.Cells["A3"].PutValue("Alice");
            sheet.Cells["B3"].PutValue(25);
            sheet.Cells["C3"].PutValue("London");

            // Define the range that will be exported (A1:C3)
            AsposeRange range = sheet.Cells.CreateRange("A1:C3");

            // Set up export options
            ExportRangeToJsonOptions options = new ExportRangeToJsonOptions
            {
                HasHeaderRow = true,      // First row contains column headers
                ExportAsString = true,    // Export all cell values as strings
                ExportEmptyCells = true,  // Include empty cells as null in JSON
                Indent = "    "           // Pretty‑print JSON with 4‑space indentation
            };

            // Export the defined range to a JSON string
            string json = JsonUtility.ExportRangeToJson(range, options);

            // Output the JSON string to the console
            Console.WriteLine(json);

            // Optionally write the JSON to a file
            string outputPath = "ExportedRange.json";
            File.WriteAllText(outputPath, json);
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}