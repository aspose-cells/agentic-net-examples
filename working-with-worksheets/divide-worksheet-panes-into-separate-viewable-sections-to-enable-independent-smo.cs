using System;
using Aspose.Cells;
using Aspose.Cells.Drawing; // For RectangleAlignmentType

namespace AsposeCellsPaneDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data to make scrolling noticeable
            for (int row = 0; row < 100; row++)
            {
                for (int col = 0; col < 10; col++)
                {
                    sheet.Cells[row, col].Value = $"R{row + 1}C{col + 1}";
                }
            }

            // Split the window into four panes (no parameters needed)
            sheet.Split();

            // Retrieve the pane collection to adjust each pane independently
            PaneCollection panes = sheet.GetPanes();

            // Set the first visible row of the bottom pane (scroll offset for bottom pane)
            panes.FirstVisibleRowOfBottomPane = 20; // rows start from 0, so row 21 becomes first visible

            // Set the first visible column of the right pane (scroll offset for right pane)
            panes.FirstVisibleColumnOfRightPane = 5; // column F becomes first visible in right pane

            // Optionally set the active pane to BottomRight for focus
            panes.AcitvePaneType = RectangleAlignmentType.BottomRight;

            // Save the workbook (lifecycle save)
            workbook.Save("PaneSplitDemo.xlsx");
        }
    }
}