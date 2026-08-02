// Title: C# – Split an Excel worksheet vertically at column 5 using Aspose.Cells
// Description: Shows how to enable pane splitting, set the first visible column of the right pane to column 5 (after column E), and save the workbook with Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# split pane | vertical split column 5 | FirstVisibleColumnOfRightPane | Excel pane splitting .NET
// Common Searches: Aspose.Cells split pane at specific column | How to set vertical split after column E in C# | GetPanes FirstVisibleColumnOfRightPane example | Split Excel view vertically using Aspose.Cells
// Developer Intent: Create a vertical split so the right pane starts at column 5.
// Use Cases: Freeze the first five columns while scrolling horizontally | Compare data side‑by‑side in large worksheets | Design a template where left columns stay visible during data entry
// AI Prompts: Write C# code to split a worksheet at column 10 and freeze the left pane with Aspose.Cells. | Show how to read and modify the current vertical split position at runtime. | Provide an example that adds both a vertical split at column 5 and a horizontal split at row 15 using Aspose.Cells.

using Aspose.Cells;

// Shows how to enable pane splitting, set the first visible column of the right pane to column 5 (after column E), and save the workbook with Aspose.Cells for .NET.
class SplitWorksheetVertically
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // (Optional) Add sample data
        for (int i = 0; i < 20; i++)
        {
            sheet.Cells[i, 0].PutValue($"Row {i + 1}");
        }

        // Split the window
        sheet.Split();

        // Get the pane collection and set the vertical split at column 5 (zero‑based index)
        PaneCollection panes = sheet.GetPanes();
        panes.FirstVisibleColumnOfRightPane = 5; // splits after column E

        // Save the workbook
        workbook.Save("SplitAtColumn5.xlsx", SaveFormat.Xlsx);
    }
}
