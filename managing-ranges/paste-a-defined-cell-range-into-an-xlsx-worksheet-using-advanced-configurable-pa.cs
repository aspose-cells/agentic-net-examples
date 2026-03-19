using System;
using Aspose.Cells;

namespace AsposeCellsPasteRangeDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate source data (A1:E5)
            for (int row = 0; row < 5; row++)
            {
                for (int col = 0; col < 5; col++)
                {
                    cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                }
            }

            // Define source and destination ranges
            // Source: A1:E5 (0,0,5,5)
            Aspose.Cells.Range sourceRange = cells.CreateRange(0, 0, 5, 5);
            // Destination: G1:K5 (0,6,5,5)
            Aspose.Cells.Range destRange = cells.CreateRange(0, 6, 5, 5);

            // Configure advanced paste options
            PasteOptions pasteOptions = new PasteOptions
            {
                // Copy only values and number formats
                PasteType = PasteType.ValuesAndNumberFormats,
                // Skip blank cells in the source range
                SkipBlanks = true,
                // Do not transpose rows/columns
                Transpose = false,
                // Add values when both source and destination contain numbers
                OperationType = PasteOperationType.Add,
                // Keep any existing tables in the destination range
                KeepOldTables = true,
                // Do not ignore links to the original file
                IgnoreLinksToOriginalFile = false,
                // Copy only visible cells (if there are hidden rows/columns)
                OnlyVisibleCells = true,
                // Do not shift formulas of shapes
                ShiftFormulasOfShapes = false
            };

            // Perform the paste operation with the configured options
            destRange.Copy(sourceRange, pasteOptions);

            // Save the workbook to an XLSX file
            workbook.Save("AdvancedPasteOptionsDemo.xlsx");
        }
    }
}