using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsRangeCopyDemo
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Fill source range (A1:E5) with sample data
            for (int row = 0; row < 5; row++)
            {
                for (int col = 0; col < 5; col++)
                {
                    cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                }
            }

            // Create source and destination Range objects
            // Source: rows 0-4, columns 0-4 (A1:E5)
            AsposeRange sourceRange = cells.CreateRange(0, 0, 5, 5);
            // Destination: rows 7-11, columns 0-4 (A8:E12)
            AsposeRange destinationRange = cells.CreateRange(7, 0, 5, 5);

            // Set paste options – copy everything (values, formulas, formats, etc.)
            PasteOptions pasteOptions = new PasteOptions
            {
                PasteType = PasteType.All
            };

            // Copy source range to destination range using the paste options
            destinationRange.Copy(sourceRange, pasteOptions);

            // Save the workbook to an XLSX file
            workbook.Save("RangeCopyResult.xlsx");
        }
    }
}