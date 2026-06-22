using System;
using Aspose.Cells;

namespace FreezePaneAdjustmentDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data (10 rows, 3 columns)
            for (int r = 0; r < 10; r++)
                for (int c = 0; c < 3; c++)
                    cells[r, c].PutValue($"R{r + 1}C{c + 1}");

            // Freeze panes at row index 5 (6th row) and column index 0 (first column)
            // This freezes the first 5 rows (0‑4) and the first column (0)
            sheet.FreezePanes(5, 0, 5, 1);

            // Retrieve current freeze pane settings
            int frozenRow, frozenColumn, frozenRows, frozenColumns;
            bool hasFreeze = sheet.GetFreezedPanes(out frozenRow, out frozenColumn, out frozenRows, out frozenColumns);

            if (!hasFreeze)
            {
                Console.WriteLine("No frozen panes detected.");
                return;
            }

            // Define how many rows above the freeze point we want to delete
            int rowsToDelete = 3; // delete rows 0,1,2 (first three rows)

            // Delete the rows
            cells.DeleteRows(0, rowsToDelete);

            // Adjust the freeze pane position:
            // New freeze row index = original frozenRow - rowsToDelete
            int newFrozenRow = frozenRow - rowsToDelete;
            // Column index remains the same
            int newFrozenColumn = frozenColumn;

            // Unfreeze first to avoid conflicts
            sheet.UnFreezePanes();

            // Apply the updated freeze pane settings
            sheet.FreezePanes(newFrozenRow, newFrozenColumn, frozenRows, frozenColumns);

            // Save the workbook
            workbook.Save("AdjustedFreezePane.xlsx");
        }
    }
}