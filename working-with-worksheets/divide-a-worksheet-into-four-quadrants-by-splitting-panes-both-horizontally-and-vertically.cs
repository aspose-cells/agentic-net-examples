// Title: Split an Excel worksheet into four quadrants with Aspose.Cells for .NET (C#)
// Description: C# example that creates a workbook, populates sample data, enables pane splitting, defines horizontal and vertical split positions, selects a specific quadrant as the active pane, and saves the file as an XLSX document using Aspose.Cells.
// Keywords: Aspose.Cells | C# | .NET | split worksheet panes | Excel quadrants | horizontal split | vertical split | first visible row | first visible column | active pane | PaneCollection
// Common Searches: Aspose.Cells split worksheet into panes C# | how to create four quadrants in Excel with Aspose | set active pane after splitting Excel sheet Aspose.Cells | horizontal and vertical split positions Aspose.Cells | C# example for pane splitting in Excel workbook
// Developer Intent: Produce a workbook that displays four independent view areas by applying both a horizontal and a vertical split and optionally activate one of the quadrants.
// Use Cases: Navigate large tables while keeping header rows and columns visible in separate panes. | Design a dashboard where each quadrant shows a different chart or summary table. | Allow simultaneous scrolling of distinct data sections for side‑by‑side comparison.
// AI Prompts: Generate C# code with Aspose.Cells that splits a worksheet at row 15 and column 8, then sets the top‑left pane as active. | Explain how to retrieve and modify PaneCollection properties such as FirstVisibleRowOfBottomPane and FirstVisibleColumnOfRightPane after calling Worksheet.Split(). | Provide a complete example that creates quadrants, inserts a chart in each pane, and exports the workbook to PDF using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// C# example that creates a workbook, populates sample data, enables pane splitting, defines horizontal and vertical split positions, selects a specific quadrant as the active pane, and saves the file as an XLSX document using Aspose.Cells.
class SplitWorksheetQuadrants
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data (optional, just to visualize the quadrants)
        for (int i = 0; i < 30; i++)
        {
            for (int j = 0; j < 15; j++)
            {
                sheet.Cells[i, j].PutValue($"R{i + 1}C{j + 1}");
            }
        }

        // Split the worksheet window
        sheet.Split();

        // Retrieve the pane collection to configure split positions
        PaneCollection panes = sheet.GetPanes();

        // Set split after row 10 (zero‑based) and column 5
        panes.FirstVisibleRowOfBottomPane = 10;   // Top pane shows rows 0‑9, bottom pane starts at row 10
        panes.FirstVisibleColumnOfRightPane = 5; // Left pane shows columns 0‑4, right pane starts at column 5

        // Optionally set the active pane (e.g., bottom‑right quadrant)
        panes.AcitvePaneType = RectangleAlignmentType.BottomRight;

        // Save the workbook
        workbook.Save("QuadrantsSplit.xlsx", SaveFormat.Xlsx);
    }
}
