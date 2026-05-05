using System;
using Aspose.Cells;

namespace WorksheetViewDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (uses the provided creation rule)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // -------------------------------------------------
            // Configure view options
            // -------------------------------------------------

            // 1. Set the view type to Normal (required before setting Zoom)
            sheet.ViewType = ViewType.NormalView;

            // 2. Set zoom level to 150%
            sheet.Zoom = 150;

            // 3. Freeze panes at cell C4 (row index 3, column index 2) with 3 frozen rows and 3 frozen columns
            sheet.FreezePanes("C4", 3, 3);

            // 4. Split the window (creates a split view without freezing)
            sheet.Split();

            // -------------------------------------------------
            // Query view options
            // -------------------------------------------------

            // Query zoom level
            Console.WriteLine($"Current Zoom: {sheet.Zoom}%");

            // Query pane state (Frozen, Split, etc.)
            Console.WriteLine($"Pane State: {sheet.PaneState}");

            // Retrieve freeze pane details
            int frozenRow, frozenColumn, frozenRowsCount, frozenColumnsCount;
            bool hasFreeze = sheet.GetFreezedPanes(out frozenRow, out frozenColumn, out frozenRowsCount, out frozenColumnsCount);
            Console.WriteLine($"Has Freeze Panes: {hasFreeze}");
            if (hasFreeze)
            {
                Console.WriteLine($"Freeze Position - Row: {frozenRow}, Column: {frozenColumn}");
                Console.WriteLine($"Frozen Rows: {frozenRowsCount}, Frozen Columns: {frozenColumnsCount}");
            }

            // Retrieve split pane information
            PaneCollection panes = sheet.GetPanes();
            Console.WriteLine($"First Visible Row of Bottom Pane: {panes.FirstVisibleRowOfBottomPane}");
            Console.WriteLine($"First Visible Column of Right Pane: {panes.FirstVisibleColumnOfRightPane}");

            // -------------------------------------------------
            // Clean up view settings (optional)
            // -------------------------------------------------

            // Remove split view
            sheet.RemoveSplit();

            // Unfreeze panes
            sheet.UnFreezePanes();

            // -------------------------------------------------
            // Save the workbook (uses the provided saving rule)
            // -------------------------------------------------
            workbook.Save("WorksheetViewDemo.xlsx");
        }
    }
}