using System;
using Aspose.Cells;
using Aspose.Cells.Utility;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsJsonExport
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data (header + two rows)
            worksheet.Cells["A1"].PutValue("Name");
            worksheet.Cells["B1"].PutValue("Age");
            worksheet.Cells["A2"].PutValue("John");
            worksheet.Cells["B2"].PutValue(30);
            worksheet.Cells["A3"].PutValue("Alice");
            worksheet.Cells["B3"].PutValue(25);

            // Define a range that covers the populated cells (A1:B3)
            AsposeRange range = worksheet.Cells.CreateRange(0, 0, 3, 2); // startRow, startColumn, rowCount, columnCount

            // Configure JSON export options
            JsonSaveOptions jsonOptions = new JsonSaveOptions
            {
                HasHeaderRow = true,          // first row is header
                ExportEmptyCells = true       // include empty cells as null
            };

            // Export the range to a JSON string
            string json = JsonUtility.ExportRangeToJson(range, jsonOptions);

            // Print the JSON representation to standard output
            Console.WriteLine(json);
        }
    }
}