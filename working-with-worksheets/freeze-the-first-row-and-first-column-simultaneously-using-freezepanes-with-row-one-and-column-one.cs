// Title: C# Example: Freeze First Row and First Column Simultaneously with Aspose.Cells
// Description: This Aspose.Cells for .NET sample creates an in‑memory workbook, accesses the first worksheet, and uses worksheet.FreezePanes(1,1,1,1) to lock the top row and leftmost column together. It then reads back the frozen pane settings with GetFreezedPanes and saves the file as FreezeFirstRowAndColumn.xlsx.
// Keywords: Aspose.Cells | C# | .NET | FreezePanes | freeze first row | freeze first column | freeze panes Excel | worksheet freeze | Excel automation | Aspose.Cells example | GitHub sample
// Common Searches: How to freeze the top row and left column with Aspose.Cells C# | Aspose.Cells FreezePanes method parameters tutorial | Retrieve frozen pane settings using Aspose.Cells | Freeze first row and column Excel Aspose.Cells example | Aspose.Cells C# code to lock header row and index column
// Developer Intent: Apply a simultaneous freeze to the first row and first column of a worksheet using Aspose.Cells.
// Use Cases: Generate reports where header rows and index columns stay visible while scrolling. | Create data‑entry templates that keep the first row and column fixed for easier navigation. | Programmatically verify frozen pane configuration before distributing a workbook.
// AI Prompts: Write C# code that freezes the first two rows and first three columns with Aspose.Cells. | Show how to unfreeze panes and then apply a new FreezePanes setting in an existing workbook. | Explain how to read, modify, and re‑apply frozen pane settings in an Excel file using Aspose.Cells.

using System;
using Aspose.Cells;

namespace FreezeFirstRowAndColumn
{
    // This Aspose.Cells for .NET sample creates an in‑memory workbook, accesses the first worksheet, and uses worksheet.FreezePanes(1,1,1,1) to lock the top row and leftmost column together. It then reads back the frozen pane settings with GetFreezedPanes and saves the file as FreezeFirstRowAndColumn.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (in‑memory)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Freeze the first row and first column simultaneously.
            // Parameters: row index, column index, number of frozen rows, number of frozen columns.
            // Row and column indices are zero‑based, so (1,1) freezes the area above row 1 and left of column 1.
            worksheet.FreezePanes(1, 1, 1, 1);

            // Optionally verify the freeze settings
            int row, column, frozenRows, frozenColumns;
            bool hasFreeze = worksheet.GetFreezedPanes(out row, out column, out frozenRows, out frozenColumns);
            Console.WriteLine($"Freeze applied: {hasFreeze} (Row={row}, Column={column}, FrozenRows={frozenRows}, FrozenColumns={frozenColumns})");

            // Save the workbook to a file
            workbook.Save("FreezeFirstRowAndColumn.xlsx");
        }
    }
}
