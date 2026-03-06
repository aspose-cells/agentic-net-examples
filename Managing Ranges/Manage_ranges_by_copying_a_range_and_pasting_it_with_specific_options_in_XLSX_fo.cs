using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsRangeCopyDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Fill sample data into a source range (5 rows x 5 columns)
            for (int row = 0; row < 5; row++)
            {
                for (int col = 0; col < 5; col++)
                {
                    sheet.Cells[row, col].PutValue($"R{row}C{col}");
                }
            }

            // Create source and destination Range objects
            AsposeRange sourceRange = sheet.Cells.CreateRange(0, 0, 5, 5);
            AsposeRange destinationRange = sheet.Cells.CreateRange(6, 0, 5, 5);

            // Configure paste options
            PasteOptions pasteOptions = new PasteOptions
            {
                PasteType = PasteType.All,
                SkipBlanks = true,
                Transpose = false
            };

            // Copy the source range into the destination range with the specified options
            destinationRange.Copy(sourceRange, pasteOptions);

            // Save the workbook
            workbook.Save("RangeCopyWithPasteOptions.xlsx");
        }
    }
}