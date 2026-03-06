using System;
using Aspose.Cells;

namespace AsposeCellsRangePasteDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate source range with sample data (A1:C3)
            for (int row = 0; row < 3; row++)
            {
                for (int col = 0; col < 3; col++)
                {
                    cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                }
            }

            // Define source and destination ranges
            // Source: rows 0-2, columns 0-2 (A1:C3)
            Aspose.Cells.Range sourceRange = cells.CreateRange(0, 0, 3, 3);
            // Destination: rows 5-7, columns 0-2 (A6:C8)
            Aspose.Cells.Range destinationRange = cells.CreateRange(5, 0, 3, 3);

            // Configure paste options
            PasteOptions pasteOptions = new PasteOptions
            {
                PasteType = PasteType.Values,      // copy only values
                SkipBlanks = true,                // ignore blank cells in source
                Transpose = false,                // keep original orientation
                OnlyVisibleCells = false,         // include hidden cells if any
                OperationType = PasteOperationType.None,
                IgnoreLinksToOriginalFile = true
            };

            // Perform the paste operation with the specified options
            destinationRange.Copy(sourceRange, pasteOptions);

            // Save the workbook in XLSX format
            workbook.Save("RangePasteWithOptions.xlsx");
        }
    }
}