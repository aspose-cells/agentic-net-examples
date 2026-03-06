using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace RangeCopyExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Fill sample data in the source range (A1:C3)
            for (int row = 0; row < 3; row++)
            {
                for (int col = 0; col < 3; col++)
                {
                    cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                }
            }

            // Define source and destination ranges
            // Source: A1:C3 (rows 0-2, columns 0-2)
            AsposeRange sourceRange = cells.CreateRange(0, 0, 3, 3);
            // Destination: E5:G7 (rows 4-6, columns 4-6)
            AsposeRange destinationRange = cells.CreateRange(4, 4, 3, 3);

            // Copy the source range to the destination range (including data, formulas, formatting, etc.)
            destinationRange.Copy(sourceRange);

            // Save the workbook in XLSX format
            workbook.Save("RangeCopyDemo.xlsx");
        }
    }
}