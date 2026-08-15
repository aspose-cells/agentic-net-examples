// Title: Reorder a Worksheet and Freeze Column A using Aspose.Cells for .NET (C#)
// Description: Create a workbook, add three sheets, move "Sheet3" to the second position with MoveTo, freeze column A on that sheet via FreezePanes, and save the file as MovedAndFrozen.xlsx.
// Keywords: Aspose.Cells | C# | MoveTo | FreezePanes | reorder worksheet | freeze first column | Excel automation | worksheet order | freeze column A
// Common Searches: Aspose.Cells move worksheet to specific index C# | Freeze column A after moving sheet Aspose.Cells | How to reorder worksheets and apply freeze panes in .NET | C# example MoveTo and FreezePanes Aspose.Cells | Change sheet order and lock first column Aspose.Cells
// Developer Intent: Reorder a worksheet within a workbook and apply a freeze pane so column A remains visible while scrolling.
// Use Cases: Place a summary sheet at the start of a report, then keep identifier column fixed on the data sheet for quick navigation. | Generate a financial workbook where the detailed data sheet is positioned deliberately and account IDs in column A stay static. | Programmatically build an Excel file that requires a specific sheet sequence and a frozen first column before distribution.
// AI Prompts: Show a C# snippet that moves a worksheet to index 1 and freezes column A using Aspose.Cells. | Provide an Aspose.Cells example that reorders sheets and applies FreezePanes to the first column. | Explain the FreezePanes parameters needed to lock only column A after moving a sheet in Aspose.Cells.

using System;
using Aspose.Cells;

// Create a workbook, add three sheets, move "Sheet3" to the second position with MoveTo, freeze column A on that sheet via FreezePanes, and save the file as MovedAndFrozen.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (contains a default sheet named "Sheet1")
            Workbook workbook = new Workbook();

            // Remove the default sheet to avoid duplicate name errors
            if (workbook.Worksheets.Count > 0)
                workbook.Worksheets.RemoveAt(0);

            // Add sample worksheets with unique names
            workbook.Worksheets.Add("Sheet1");
            workbook.Worksheets.Add("Sheet2");
            workbook.Worksheets.Add("Sheet3");

            // Move "Sheet3" to the second position (index 1)
            Worksheet sheetToMove = workbook.Worksheets["Sheet3"];
            sheetToMove.MoveTo(1);

            // Freeze the first column (A) on the moved sheet
            // FreezePanes(rowIndex, columnIndex, freezedRows, freezedColumns)
            // Setting columnIndex to 1 (B) freezes column A; no rows are frozen.
            sheetToMove.FreezePanes(0, 1, 0, 1);

            // Save the workbook
            workbook.Save("MovedAndFrozen.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
