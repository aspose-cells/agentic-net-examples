using System;
using Aspose.Cells;

namespace AsposeCellsDemo
{
    class DefineSourceRangeWithCellArea
    {
        static void Main()
        {
            // Create a new workbook (XLSX format by default)
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate some sample data in the worksheet (A1 to D5)
            for (int row = 0; row < 5; row++)
            {
                for (int col = 0; col < 4; col++)
                {
                    cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                }
            }

            // Define the source range using CellArea.
            // This range corresponds to cells A1:C5 (rows 0‑4, columns 0‑2)
            CellArea sourceRange = CellArea.CreateCellArea(0, 0, 4, 2);

            // Example usage: clear the defined source range (contents + formatting)
            sheet.Cells.ClearRange(sourceRange);

            // Save the workbook as an XLSX file
            workbook.Save("SourceRangeDefined.xlsx");
        }
    }
}