using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsRangeCopyDemo
{
    class Program
    {
        static void Main()
        {
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            for (int row = 0; row < 5; row++)
            {
                for (int col = 0; col < 5; col++)
                {
                    cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                }
            }

            AsposeRange sourceRange = cells.CreateRange(0, 0, 5, 5);
            AsposeRange destRange = cells.CreateRange(0, 6, 5, 5);

            PasteOptions options = new PasteOptions
            {
                PasteType = PasteType.All
            };

            destRange.Copy(sourceRange, options);

            workbook.Save("RangeCopyResult.xlsx");
        }
    }
}