using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsOffsetCopyDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Populate original range D1:E2 with sample data
                cells["D1"].PutValue("Item 1");
                cells["E1"].PutValue(100);
                cells["D2"].PutValue("Item 2");
                cells["E2"].PutValue(200);

                // Create the original range (D1:E2)
                AsposeRange originalRange = cells.CreateRange("D1", "E2");

                // Get a new range that is three columns to the left (A1:B2)
                // Row offset = 0, Column offset = -3
                AsposeRange offsetRange = originalRange.GetOffset(0, -3);

                // Copy values from the original range to the offset range
                offsetRange.CopyValue(originalRange);

                // Save the workbook
                workbook.Save("OffsetCopyDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}