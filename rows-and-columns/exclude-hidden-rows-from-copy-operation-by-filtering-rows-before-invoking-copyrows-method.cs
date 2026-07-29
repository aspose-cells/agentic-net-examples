// Title: Copy only visible rows with Aspose.Cells for .NET (C#) – skip hidden rows
// Description: This C# example shows how to create a source workbook, hide specific rows, and copy only the visible rows to a new workbook. It uses the IsRowHidden check inside a loop and the CopyRows method to transfer each non‑hidden row, then saves the result as VisibleRowsCopyResult.xlsx.
// Keywords: Aspose.Cells copy visible rows | exclude hidden rows C# | CopyRows hidden rows filter | .NET spreadsheet copy rows | skip hidden rows Aspose.Cells | copy visible data Aspose.Cells
// Common Searches: Aspose.Cells copy only visible rows | How to ignore hidden rows when copying in .NET | CopyRows method skip hidden rows | C# copy visible rows Aspose.Cells | filter hidden rows before copying spreadsheet
// Developer Intent: Copy non‑hidden rows from one worksheet to another using Aspose.Cells.
// Use Cases: Generate a report that contains only rows the user left visible after filtering. | Create a clean data export by omitting rows hidden for internal calculations. | Synchronize visible data between workbooks while preserving formatting.
// AI Prompts: Show me a C# code snippet that copies only visible rows with Aspose.Cells without manual row checks. | Explain how to use IsRowHidden and CopyRows together to skip hidden rows in Aspose.Cells for .NET. | Provide an example that copies visible rows and retains cell styles using Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsCopyVisibleRows
{
    // This C# example shows how to create a source workbook, hide specific rows, and copy only the visible rows to a new workbook. It uses the IsRowHidden check inside a loop and the CopyRows method to transfer each non‑hidden row, then saves the result as VisibleRowsCopyResult.xlsx.
    class Program
    {
        static void Main()
        {
            // Create source workbook and fill sample data
            Workbook sourceWorkbook = new Workbook();
            Worksheet sourceSheet = sourceWorkbook.Worksheets[0];
            Cells srcCells = sourceSheet.Cells;

            for (int i = 0; i < 10; i++)
            {
                srcCells[i, 0].PutValue($"Row {i + 1}");
            }

            // Hide some rows (zero‑based indexes 2, 4, 6)
            srcCells.HideRow(2);
            srcCells.HideRow(4);
            srcCells.HideRow(6);

            // Create destination workbook
            Workbook destinationWorkbook = new Workbook();
            Cells destCells = destinationWorkbook.Worksheets[0].Cells;

            // Copy only visible rows from source to destination
            int sourceRowCount = srcCells.MaxDataRow + 1; // total rows with data
            int destRowIndex = 0; // where to paste in destination

            for (int i = 0; i < sourceRowCount; i++)
            {
                // Skip hidden rows
                if (srcCells.IsRowHidden(i))
                    continue;

                // Copy a single visible row
                destCells.CopyRows(srcCells, i, destRowIndex, 1);
                destRowIndex++;
            }

            // Save the result
            destinationWorkbook.Save("VisibleRowsCopyResult.xlsx");
        }
    }
}
