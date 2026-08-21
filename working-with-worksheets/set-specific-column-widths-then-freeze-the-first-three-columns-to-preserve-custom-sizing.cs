// Title: Set column widths and freeze the first three columns with Aspose.Cells for .NET (C#)
// Description: Creates a workbook, sets column A, B, and C to 20, 30, and 15 character units, freezes columns A‑C, and saves the file as ColumnWidthsAndFreeze.xlsx using Aspose.Cells for .NET.
// Keywords: Aspose.Cells column width C# | Aspose.Cells freeze panes C# | set column width Aspose.Cells | freeze first columns Aspose.Cells | Aspose.Cells FreezePanes example | Aspose.Cells worksheet formatting | Aspose.Cells .NET column sizing
// Common Searches: Aspose.Cells set column width and freeze columns | C# freeze first three columns Aspose.Cells | How to use FreezePanes after setting column widths in Aspose.Cells | Aspose.Cells column width example C# | Freeze panes while preserving custom column sizes Aspose.Cells
// Developer Intent: Define exact column widths and lock the first three columns in a worksheet.
// Use Cases: Financial statements where identifier columns stay visible while scrolling through rows. | Data‑entry templates with fixed widths for product codes, names, and quantities, keeping them static during navigation. | Printable reports where reference columns remain on screen as users scroll horizontally.
// AI Prompts: Generate an Aspose.Cells for .NET code snippet that sets column widths based on content length and then freezes a configurable number of columns. | Show how to apply a style to frozen columns after setting custom widths in C# with Aspose.Cells. | Explain how to determine the number of columns to freeze dynamically while preserving previously set column widths using Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsColumnWidthAndFreezeDemo
{
    // Creates a workbook, sets column A, B, and C to 20, 30, and 15 character units, freezes columns A‑C, and saves the file as ColumnWidthsAndFreeze.xlsx using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Set specific column widths (in character units)
            // Column A (index 0)
            cells.SetColumnWidth(0, 20.0);
            // Column B (index 1)
            cells.SetColumnWidth(1, 30.0);
            // Column C (index 2)
            cells.SetColumnWidth(2, 15.0);

            // Freeze the first three columns.
            // Freeze at column index 3 (i.e., column D) with 0 frozen rows and 3 frozen columns.
            worksheet.FreezePanes(0, 3, 0, 3);

            // Save the workbook
            workbook.Save("ColumnWidthsAndFreeze.xlsx");
        }
    }
}
