using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;
using AsposeRange = Aspose.Cells.Range;

class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data
            worksheet.Cells["A1"].PutValue("Name");
            worksheet.Cells["B1"].PutValue("Age");
            worksheet.Cells["A2"].PutValue("John");
            worksheet.Cells["B2"].PutValue(30);
            worksheet.Cells["A3"].PutValue("Alice");
            worksheet.Cells["B3"].PutValue(25);

            // Define the range that covers the used cells
            int lastRow = worksheet.Cells.MaxDataRow;
            int lastColumn = worksheet.Cells.MaxDataColumn;
            AsposeRange range = worksheet.Cells.CreateRange(0, 0, lastRow + 1, lastColumn + 1);

            // Configure JSON export options
            JsonSaveOptions jsonOptions = new JsonSaveOptions
            {
                HasHeaderRow = true,          // first row contains column names
                ExportEmptyCells = true,      // include empty cells as null
                Indent = "    "               // pretty‑print with 4‑space indentation
            };

            // Export the defined range to a JSON string
            string json = JsonUtility.ExportRangeToJson(range, jsonOptions);

            // Output the JSON representation of the first worksheet
            Console.WriteLine(json);
        }
        catch (FileNotFoundException ex)
        {
            Console.Error.WriteLine($"File not found: {ex.FileName}");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}