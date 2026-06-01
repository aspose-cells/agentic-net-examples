using System;
using Aspose.Cells;

namespace MergeHiddenRowsDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate some data for visual reference
            for (int row = 0; row < 10; row++)
            {
                cells[row, 0].PutValue($"Row {row + 1}");
            }

            // Hide rows 3 to 5 (zero‑based indices 2,3,4)
            cells.HideRows(2, 3);

            // Merge cells A1:A5 (rows 0‑4, column 0) – this range includes the hidden rows
            cells.Merge(0, 0, 5, 1);

            // Verify that the hidden rows are still hidden after merging
            for (int i = 2; i <= 4; i++)
            {
                bool isHidden = worksheet.Cells.Rows[i].IsHidden;
                Console.WriteLine($"Row {i + 1} hidden status after merge: {isHidden}");
            }

            // Save the workbook to verify the result manually if needed
            workbook.Save("MergeHiddenRowsDemo.xlsx");
        }
    }
}