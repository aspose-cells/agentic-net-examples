// Title: C# – Move a Worksheet to the First Tab and Freeze Column A with Aspose.Cells
// Description: Shows how to reorder worksheets by moving a sheet to index 0 and then lock the first column (column A) using FreezePanes in Aspose.Cells for .NET.
// Keywords: Aspose.Cells move worksheet C# | freeze column A Aspose.Cells | reorder worksheets .NET | FreezePanes C# example | move sheet to first position Aspose | lock first column Excel .NET | Aspose.Cells worksheet ordering | freeze panes only column Aspose
// Common Searches: move worksheet to first tab Aspose.Cells C# | freeze first column in a specific sheet Aspose.Cells | Aspose.Cells reorder sheets and freeze panes | how to use FreezePanes for column A in C# | Aspose.Cells move sheet to index 0
// Developer Intent: Reorder a worksheet to the leftmost tab and keep column A fixed while scrolling.
// Use Cases: Place a summary or index sheet at the beginning of a workbook and keep identifiers visible during horizontal scrolling. | Create a dashboard where the primary data sheet is moved to the front and the key ID column stays locked. | Organize multi‑sheet reports by moving the most important sheet to the first position and freezing its first column for quick reference.
// AI Prompts: Generate C# code that moves a worksheet to index 0 and freezes column A using Aspose.Cells. | Explain the parameters of FreezePanes when only the first column should be locked after moving a sheet. | Show an Aspose.Cells example that reorders worksheets and applies FreezePanes to column A in the moved sheet.

using System;
using Aspose.Cells;

// Shows how to reorder worksheets by moving a sheet to index 0 and then lock the first column (column A) using FreezePanes in Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Add sample worksheets
        workbook.Worksheets.Add("First");
        workbook.Worksheets.Add("Second");
        workbook.Worksheets.Add("Third");

        // Move the worksheet named "Third" to the first position (index 0)
        Worksheet sheetToMove = workbook.Worksheets["Third"];
        sheetToMove.MoveTo(0);

        // Freeze the first column (column A) in the moved worksheet
        // Freeze at cell B1 with 0 frozen rows and 1 frozen column
        sheetToMove.FreezePanes("B1", 0, 1);

        // Save the workbook
        workbook.Save("MovedAndFrozen.xlsx");
    }
}
