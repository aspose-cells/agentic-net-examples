// Title: C# – Split worksheet into left 3 columns and top 4 rows using Aspose.Cells
// Description: Demonstrates how to create a workbook, populate sample data, split the view, and set the first visible column of the right pane to column D and the first visible row of the bottom pane to row 5, then save as SplitPaneDemo.xlsx.
// Keywords: Aspose.Cells split pane C# | first visible column of right pane | first visible row of bottom pane | freeze left columns Aspose.Cells | freeze top rows Aspose.Cells | Excel pane configuration .NET | split worksheet window Aspose
// Common Searches: Aspose.Cells split pane left three columns | set first visible column of right pane example | freeze top four rows with Aspose.Cells | C# split worksheet view Aspose.Cells | how to configure split panes in Excel using .NET
// Developer Intent: Configure a worksheet view so the left pane shows three columns and the top pane shows four rows.
// Use Cases: Keep identifier columns and header rows visible while scrolling large tables. | Build a dashboard where the first rows and columns stay static for quick reference. | Prepare financial or inventory reports that require frozen top rows and left columns for context.
// AI Prompts: Write C# code with Aspose.Cells that splits a worksheet so the left pane displays the first three columns and the top pane displays the first four rows, then saves the file. | Explain the zero‑based indexing of FirstVisibleColumnOfRightPane and FirstVisibleRowOfBottomPane in Aspose.Cells. | Show how to adjust split pane positions dynamically based on the size of a data range in a .NET workbook.

using Aspose.Cells;
using System;

// Demonstrates how to create a workbook, populate sample data, split the view, and set the first visible column of the right pane to column D and the first visible row of the bottom pane to row 5, then save as SplitPaneDemo.xlsx.
class SplitPaneDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Optional: populate some data for demonstration
        for (int i = 0; i < 20; i++)
        {
            for (int j = 0; j < 10; j++)
            {
                worksheet.Cells[i, j].PutValue($"Row {i + 1}, Col {j + 1}");
            }
        }

        // Split the worksheet window
        worksheet.Split();

        // Access the pane collection to configure split positions
        PaneCollection panes = worksheet.GetPanes();

        // Set the first visible column of the right pane (left pane will show 3 columns)
        panes.FirstVisibleColumnOfRightPane = 3; // zero‑based index, column D becomes first visible in right pane

        // Set the first visible row of the bottom pane (top pane will show 4 rows)
        panes.FirstVisibleRowOfBottomPane = 4; // zero‑based index, row 5 becomes first visible in bottom pane

        // Save the workbook
        workbook.Save("SplitPaneDemo.xlsx");
    }
}
