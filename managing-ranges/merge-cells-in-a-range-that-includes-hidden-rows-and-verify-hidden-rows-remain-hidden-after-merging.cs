using System;
using Aspose.Cells;

class MergeHiddenRowsDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Fill some sample data in rows 0-4 and columns 0-2
        for (int row = 0; row < 5; row++)
        {
            for (int col = 0; col < 3; col++)
            {
                cells[row, col].PutValue($"R{row}C{col}");
            }
        }

        // Hide rows 1 to 3 (zero‑based indices)
        cells.HideRows(1, 3);

        // Merge a range that includes the hidden rows (rows 0‑4, columns 0‑1)
        cells.Merge(0, 0, 5, 2);

        // Verify that the rows that were hidden remain hidden after merging
        for (int row = 0; row < 5; row++)
        {
            bool isHidden = worksheet.Cells.Rows[row].IsHidden;
            Console.WriteLine($"Row {row} hidden: {isHidden}");
        }

        // Save the workbook
        workbook.Save("MergeHiddenRowsDemo.xlsx");
    }
}