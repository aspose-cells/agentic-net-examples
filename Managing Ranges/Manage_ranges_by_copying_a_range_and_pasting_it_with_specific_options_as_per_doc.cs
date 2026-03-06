using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsRangeCopyDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate source data in the range A1:C5
            for (int row = 0; row < 5; row++)
            {
                for (int col = 0; col < 3; col++)
                {
                    sheet.Cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                }
            }

            // Create source and destination Range objects
            // Source: rows 0-4, columns 0-2 (A1:C5)
            AsposeRange sourceRange = sheet.Cells.CreateRange(0, 0, 5, 3);
            // Destination: rows 0-4, columns 4-6 (E1:G5)
            AsposeRange destRange = sheet.Cells.CreateRange(0, 4, 5, 3);

            // Configure paste options
            PasteOptions pasteOptions = new PasteOptions
            {
                // Copy only values (no formulas, formats, etc.)
                PasteType = PasteType.Values,
                // Skip blank cells in the source range
                SkipBlanks = true,
                // Do not transpose rows/columns
                Transpose = false,
                // Include hidden cells as well
                OnlyVisibleCells = false,
                // No special operation type
                OperationType = PasteOperationType.None
            };

            // Perform the copy with the specified paste options
            destRange.Copy(sourceRange, pasteOptions);

            // Save the workbook in XLSX format
            workbook.Save("RangeCopyWithOptions.xlsx");
        }
    }
}