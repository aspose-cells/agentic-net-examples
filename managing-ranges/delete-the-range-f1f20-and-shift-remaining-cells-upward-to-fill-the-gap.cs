using System;
using Aspose.Cells;

namespace AsposeCellsDeleteRangeExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data in column F (index 5) for rows 1 to 20
            for (int row = 0; row < 20; row++)
            {
                cells[row, 5].PutValue($"F{row + 1}");
            }

            // Delete the range F1:F20 and shift remaining cells upward
            // startRow = 0 (F1), startColumn = 5 (column F)
            // endRow = 19 (F20), endColumn = 5 (column F)
            cells.DeleteRange(0, 5, 19, 5, ShiftType.Up);

            // Save the modified workbook
            workbook.Save("DeletedRangeUpShift.xlsx");
        }
    }
}