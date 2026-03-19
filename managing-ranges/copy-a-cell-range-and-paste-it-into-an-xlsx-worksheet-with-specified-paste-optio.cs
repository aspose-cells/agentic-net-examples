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

            // Populate source range (A1:E5) with sample data
            for (int row = 0; row < 5; row++)
            {
                for (int col = 0; col < 5; col++)
                {
                    sheet.Cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                }
            }

            // Define source and destination ranges
            AsposeRange sourceRange = sheet.Cells.CreateRange(0, 0, 5, 5);
            AsposeRange destinationRange = sheet.Cells.CreateRange(0, 6, 5, 5);

            // Configure paste options
            PasteOptions pasteOptions = new PasteOptions
            {
                PasteType = PasteType.All,
                SkipBlanks = true,
                OnlyVisibleCells = false,
                Transpose = false
            };

            // Perform the copy with the specified paste options
            destinationRange.Copy(sourceRange, pasteOptions);

            // Save the workbook to an XLSX file
            workbook.Save("RangeCopyWithPasteOptions.xlsx");
        }
    }
}