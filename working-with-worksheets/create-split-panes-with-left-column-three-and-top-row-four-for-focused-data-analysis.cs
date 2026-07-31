// Title: Split Worksheet into Left‑Pane (3 Columns) and Top‑Pane (4 Rows) with Aspose.Cells for .NET
// Description: Demonstrates how to create a new workbook, optionally fill it with sample data, split the worksheet window, and configure the pane collection so the left pane shows the first three columns and the top pane shows the first four rows, then save as SplitPanesDemo.xlsx.
// Keywords: Aspose.Cells split panes | C# freeze columns | freeze top rows Aspose.Cells | FirstVisibleColumnOfRightPane | FirstVisibleRowOfBottomPane | Excel pane configuration .NET | split worksheet window programmatically | Aspose.Cells pane collection
// Common Searches: Aspose.Cells split pane three columns left | set first visible row of bottom pane Aspose.Cells | freeze first 3 columns and 4 rows using Aspose.Cells C# | how to split worksheet window with Aspose.Cells .NET | configure pane positions in Excel with Aspose.Cells
// Developer Intent: The developer needs to split a worksheet so that the left pane displays three columns and the top pane displays four rows, keeping those sections static while scrolling.
// Use Cases: Keep identifier columns visible while scrolling through large datasets. | Maintain summary rows at the top of a report for constant reference. | Create a four‑quadrant view for data analysis, with independent scrolling in each pane.
// AI Prompts: Show C# code to freeze the first two columns and the first row using Aspose.Cells. | Provide an example that sets custom pixel‑based pane sizes after splitting a worksheet with Aspose.Cells. | Explain how to read, modify, and persist existing pane settings when loading an existing workbook.

using System;
using Aspose.Cells;

namespace AsposeCellsSplitPanesDemo
{
    // Demonstrates how to create a new workbook, optionally fill it with sample data, split the worksheet window, and configure the pane collection so the left pane shows the first three columns and the top pane shows the first four rows, then save as SplitPanesDemo.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some sample data (optional, just for demonstration)
            for (int i = 0; i < 20; i++)
            {
                for (int j = 0; j < 10; j++)
                {
                    sheet.Cells[i, j].PutValue($"R{i + 1}C{j + 1}");
                }
            }

            // Split the worksheet window
            sheet.Split();

            // Get the pane collection to configure split positions
            PaneCollection panes = sheet.GetPanes();

            // Set the first visible column of the right pane (zero‑based index)
            // Left pane will show columns 0,1,2 (three columns)
            panes.FirstVisibleColumnOfRightPane = 3;

            // Set the first visible row of the bottom pane (zero‑based index)
            // Top pane will show rows 0,1,2,3 (four rows)
            panes.FirstVisibleRowOfBottomPane = 4;

            // Save the workbook
            workbook.Save("SplitPanesDemo.xlsx");
        }
    }
}
