using System;
using Aspose.Cells;

class DeleteRangeExample
{
    static void Main()
    {
        // Load an existing workbook (provide the correct path to your file)
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Delete the range F1:F20.
        // Row and column indexes are zero‑based, so:
        //   startRow = 0 (row 1), startColumn = 5 (column F)
        //   endRow   = 19 (row 20), endColumn = 5 (column F)
        // Shift remaining cells upward to fill the gap.
        sheet.Cells.DeleteRange(0, 5, 19, 5, ShiftType.Up);

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}