using System;
using Aspose.Cells;

namespace AsposeCellsFreezePanesDemo
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook (in‑memory)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // -------------------------------------------------
            // 1. Freeze panes using row/column indices
            // Freeze at cell C3 (row index 2, column index 2) with
            // 2 rows and 2 columns frozen (top‑left pane size)
            // -------------------------------------------------
            sheet.FreezePanes(2, 2, 2, 2);

            // -------------------------------------------------
            // 2. Retrieve freeze pane information
            // -------------------------------------------------
            int frozenRow, frozenColumn, frozenRows, frozenColumns;
            bool hasFreeze = sheet.GetFreezedPanes(out frozenRow, out frozenColumn, out frozenRows, out frozenColumns);
            Console.WriteLine($"Has Freeze: {hasFreeze}");
            Console.WriteLine($"Freeze Position - Row: {frozenRow}, Column: {frozenColumn}");
            Console.WriteLine($"Frozen Rows: {frozenRows}, Frozen Columns: {frozenColumns}");

            // -------------------------------------------------
            // 3. Unfreeze the panes
            // -------------------------------------------------
            sheet.UnFreezePanes();

            // Verify that panes are no longer frozen
            hasFreeze = sheet.GetFreezedPanes(out frozenRow, out frozenColumn, out frozenRows, out frozenColumns);
            Console.WriteLine($"After UnFreeze - Has Freeze: {hasFreeze}");

            // -------------------------------------------------
            // 4. Freeze panes using cell address overload
            // Freeze at cell "E5" with 4 frozen rows and 3 frozen columns
            // -------------------------------------------------
            sheet.FreezePanes("E5", 4, 3);

            // -------------------------------------------------
            // 5. Access pane collection to adjust visible pane offsets
            // -------------------------------------------------
            PaneCollection panes = sheet.GetPanes();
            // Scroll the bottom pane down so that row 10 becomes the first visible row
            panes.FirstVisibleRowOfBottomPane = 10;
            // Scroll the right pane right so that column 2 becomes the first visible column
            panes.FirstVisibleColumnOfRightPane = 2;

            // -------------------------------------------------
            // 6. Save the workbook to a file
            // -------------------------------------------------
            workbook.Save("FreezePanesDemo.xlsx");
        }
    }
}