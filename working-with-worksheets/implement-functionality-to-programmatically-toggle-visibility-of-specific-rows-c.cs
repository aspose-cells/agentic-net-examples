using System;
using Aspose.Cells;

namespace AsposeCellsVisibilityDemo
{
    public class VisibilityToggle
    {
        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // -------------------------------------------------
            // 1. Toggle row visibility
            // -------------------------------------------------
            // Hide row 2 (zero‑based index)
            sheet.Cells.HideRow(1);
            // Verify that the row is hidden
            Console.WriteLine($"Row 2 hidden: {sheet.Cells.IsRowHidden(1)}");

            // Unhide the same row and set its height to default (15 points)
            sheet.Cells.UnhideRow(1, 15);
            Console.WriteLine($"Row 2 hidden after unhide: {sheet.Cells.IsRowHidden(1)}");

            // -------------------------------------------------
            // 2. Toggle column visibility
            // -------------------------------------------------
            // Hide column B (index 1)
            sheet.Cells.HideColumn(1);
            Console.WriteLine($"Column B hidden: {sheet.Cells.IsColumnHidden(1)}");

            // Unhide column B and set its width to default (8.43 characters)
            sheet.Cells.UnhideColumn(1, 8.43);
            Console.WriteLine($"Column B hidden after unhide: {sheet.Cells.IsColumnHidden(1)}");

            // -------------------------------------------------
            // 3. Scroll to a specific row and column
            // -------------------------------------------------
            // Populate some data to see the effect
            for (int i = 0; i < 30; i++)
            {
                sheet.Cells[i, 0].PutValue($"Row {i + 1}");
                sheet.Cells[i, 1].PutValue($"Data {i + 1}");
            }

            // Set the first visible row to row 15 (zero‑based index 14)
            sheet.FirstVisibleRow = 14;
            // Set the first visible column to column C (index 2)
            sheet.FirstVisibleColumn = 2;

            Console.WriteLine($"First visible row index: {sheet.FirstVisibleRow}");
            Console.WriteLine($"First visible column index: {sheet.FirstVisibleColumn}");

            // -------------------------------------------------
            // 4. Toggle worksheet scrollbars visibility
            // -------------------------------------------------
            // Hide both horizontal and vertical scrollbars
            workbook.Settings.IsHScrollBarVisible = false;
            workbook.Settings.IsVScrollBarVisible = false;
            Console.WriteLine($"Horizontal scrollbar visible: {workbook.Settings.IsHScrollBarVisible}");
            Console.WriteLine($"Vertical scrollbar visible: {workbook.Settings.IsVScrollBarVisible}");

            // Show scrollbars again
            workbook.Settings.IsHScrollBarVisible = true;
            workbook.Settings.IsVScrollBarVisible = true;

            // -------------------------------------------------
            // 5. Additional visibility options (optional)
            // -------------------------------------------------
            // Hide row and column headers
            sheet.IsRowColumnHeadersVisible = false;
            // Hide gridlines
            sheet.IsGridlinesVisible = false;

            // -------------------------------------------------
            // Save the workbook to verify the changes
            // -------------------------------------------------
            workbook.Save("VisibilityToggleDemo.xlsx", SaveFormat.Xlsx);
            Console.WriteLine("Workbook saved as VisibilityToggleDemo.xlsx");
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            VisibilityToggle.Run();
        }
    }
}