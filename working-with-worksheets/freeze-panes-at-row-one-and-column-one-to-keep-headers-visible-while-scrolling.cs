// Title: Freeze first row and column in Excel using Aspose.Cells for .NET (C#)
// Description: Creates a new Workbook, accesses the first Worksheet, calls FreezePanes(1,1,1,1) to lock the top row and leftmost column so headers remain visible while scrolling, and saves the file as FreezePanesFirstRowColumn.xlsx.
// Keywords: Aspose.Cells | FreezePanes | C# Excel freeze panes | lock header row Aspose.Cells | freeze first column .NET | Excel freeze panes example | Aspose.Cells tutorial | GitHub Aspose.Cells code | global
// Common Searches: Aspose.Cells freeze top row and first column C# | How to lock header row in Excel with Aspose.Cells .NET | Freeze panes example Aspose.Cells for .NET | C# code to keep first row and column visible in generated Excel | Aspose.Cells FreezePanes method usage
// Developer Intent: Apply FreezePanes to keep the worksheet’s header row and left‑most column static while the user scrolls.
// Use Cases: Generating large reports where column headers and row titles must stay in view. | Building a spreadsheet template that anchors the top‑left cell for easier data entry. | Exporting data tables with frozen headers to improve readability on desktop and web viewers.
// AI Prompts: Write C# code with Aspose.Cells to freeze the first two rows and the first column of a worksheet. | Show how to unfreeze panes and then refreeze them at a new position using Aspose.Cells for .NET. | Create an example that determines the number of header rows at runtime and freezes panes accordingly.

using System;
using Aspose.Cells;

// Creates a new Workbook, accesses the first Worksheet, calls FreezePanes(1,1,1,1) to lock the top row and leftmost column so headers remain visible while scrolling, and saves the file as FreezePanesFirstRowColumn.xlsx.
class FreezePanesExample
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Freeze the first row and first column (row index 1, column index 1)
        sheet.FreezePanes(1, 1, 1, 1);

        // Save the workbook
        workbook.Save("FreezePanesFirstRowColumn.xlsx");
    }
}
