// Title: Aspose.Cells .NET: Move a Range While Preserving Frozen Panes
// Description: C# example that freezes panes at C3, moves the range A1:B3 to D6:E8 with Cells.MoveRange, then restores the original FreezePanes settings so the frozen rows and columns stay unchanged.
// Keywords: Aspose.Cells | MoveRange | preserve frozen panes | FreezePanes | C# worksheet manipulation | reapply freeze settings | range relocation
// Common Searches: move range without affecting freeze panes Aspose.Cells | keep frozen rows after moving cells .NET | Aspose.Cells reapply FreezePanes after MoveRange | preserve freeze pane positions when shifting data | C# Aspose.Cells move cells keep freeze
// Developer Intent: Retain the existing FreezePanes configuration while moving a block of cells to a new location in a worksheet.
// Use Cases: Re‑arrange a report section without losing header rows/columns that are frozen for scrolling. | Shift a data table in a dashboard while maintaining user‑defined frozen panes for consistent view. | Copy a financial summary to another area of the sheet and restore the original frozen rows and columns.
// AI Prompts: Generate C# code that moves a cell range with Aspose.Cells and automatically restores the original frozen pane settings. | Explain step‑by‑step how to capture FreezePanes parameters, use Cells.MoveRange, and reapply the freeze without altering other worksheet properties. | Create a reusable method in .NET that preserves FreezePanes whenever any range is moved with Aspose.Cells.

using System;
using Aspose.Cells;

// C# example that freezes panes at C3, moves the range A1:B3 to D6:E8 with Cells.MoveRange, then restores the original FreezePanes settings so the frozen rows and columns stay unchanged.
class MoveRangeWithoutAffectingFrozenPanes
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate some sample data in the source range (A1:B3)
        worksheet.Cells["A1"].PutValue("Header");
        worksheet.Cells["A2"].PutValue(10);
        worksheet.Cells["B2"].PutValue(20);
        worksheet.Cells["A3"].PutValue(30);
        worksheet.Cells["B3"].PutValue(40);

        // Freeze panes at cell C3 (row index 2, column index 2) with 2 frozen rows and 2 frozen columns
        worksheet.FreezePanes(2, 2, 2, 2);

        // Capture the current frozen pane settings
        int frozenRow, frozenColumn, frozenRows, frozenColumns;
        bool hasFreeze = worksheet.GetFreezedPanes(out frozenRow, out frozenColumn, out frozenRows, out frozenColumns);

        // Define the source area to move (A1:B3)
        CellArea sourceArea = new CellArea();
        sourceArea.StartRow = 0;      // Row 0 (A)
        sourceArea.StartColumn = 0;   // Column 0 (A)
        sourceArea.EndRow = 2;        // Row 2 (A3)
        sourceArea.EndColumn = 1;     // Column 1 (B)

        // Move the range down 5 rows and right 3 columns (to D6:E8)
        worksheet.Cells.MoveRange(sourceArea, 5, 3);

        // Re‑apply the frozen panes after the move to keep them unchanged
        if (hasFreeze)
        {
            worksheet.FreezePanes(frozenRow, frozenColumn, frozenRows, frozenColumns);
        }

        // Save the workbook
        workbook.Save("MovedRange_NoFreezeShift.xlsx");
    }
}
