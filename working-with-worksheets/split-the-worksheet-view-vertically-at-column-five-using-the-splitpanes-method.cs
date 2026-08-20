// Title: Split Worksheet Vertically at Column 5 with Aspose.Cells for .NET (C#)
// Description: This example shows how to create a vertical split in an Excel worksheet using Aspose.Cells for .NET. After creating a workbook, the code calls Split() to enable pane splitting, sets the first visible column of the right pane to index 5 (column F) via the PaneCollection, and saves the file as SplitAtColumnFive.xlsx.
// Keywords: Aspose.Cells vertical split | split worksheet column 5 | FirstVisibleColumnOfRightPane | C# Excel pane split | Split() Aspose.Cells | Excel pane collection example
// Common Searches: Aspose.Cells split worksheet vertically at column 5 | C# set first visible column of right pane Aspose.Cells | How to create a vertical split view in Excel using Aspose.Cells | Split() method example Aspose.Cells .NET
// Developer Intent: Create a vertical pane split so that column 5 becomes the first column of the right pane in the worksheet view.
// Use Cases: Freeze a navigation column on the left while allowing the rest of the sheet to scroll from column F onward. | Build a reporting workbook where labels stay static on the left and data starts at column 5 in a scrollable pane. | Design a presentation sheet with a fixed left pane for headings and a movable right pane beginning at column F.
// AI Prompts: Generate C# code with Aspose.Cells that splits a worksheet vertically at column 5 and saves the workbook. | Explain how to use Split() and PaneCollection to set the first visible column of the right pane to index 5 in Aspose.Cells. | Provide a step‑by‑step tutorial for creating a vertical split view in an Excel file using Aspose.Cells for .NET.

using Aspose.Cells;

// This example shows how to create a vertical split in an Excel worksheet using Aspose.Cells for .NET. After creating a workbook, the code calls Split() to enable pane splitting, sets the first visible column of the right pane to index 5 (column F) via the PaneCollection, and saves the file as SplitAtColumnFive.xlsx.
class SplitWorksheetVertically
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Split the worksheet window
        sheet.Split();

        // Obtain the pane collection and set the first visible column of the right pane to column index 5 (zero‑based)
        PaneCollection panes = sheet.GetPanes();
        panes.FirstVisibleColumnOfRightPane = 5;

        // Save the workbook
        workbook.Save("SplitAtColumnFive.xlsx", SaveFormat.Xlsx);
    }
}
