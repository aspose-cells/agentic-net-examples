using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate the sheet with sample data to visualize the panes
        for (int row = 0; row < 100; row++)
        {
            for (int col = 0; col < 20; col++)
            {
                sheet.Cells[row, col].PutValue($"R{row + 1}C{col + 1}");
            }
        }

        // Split the window into four panes (horizontal and vertical split)
        sheet.Split();

        // Retrieve the collection that represents the panes
        PaneCollection panes = sheet.GetPanes();

        // Set the first visible row of the bottom pane (e.g., start scrolling from row 20)
        panes.FirstVisibleRowOfBottomPane = 20;

        // Set the first visible column of the right pane (e.g., start scrolling from column 5)
        panes.FirstVisibleColumnOfRightPane = 5;

        // Optionally set the active pane to BottomRight for demonstration
        panes.AcitvePaneType = RectangleAlignmentType.BottomRight;

        // Save the workbook with the configured pane splits
        workbook.Save("PaneSplitDemo.xlsx", SaveFormat.Xlsx);
    }
}