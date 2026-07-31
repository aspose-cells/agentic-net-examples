// Title: Copy Only Visible Rows with Aspose.Cells for .NET (C#)
// Description: Demonstrates how to hide specific rows, detect hidden rows using IsRowHidden, and copy only the visible rows from a source worksheet to a new workbook with the CopyRows method, then save the result as VisibleRowsCopy.xlsx.
// Keywords: Aspose.Cells copy visible rows | exclude hidden rows Aspose.Cells | CopyRows filter hidden rows | IsRowHidden C# example | Aspose.Cells .NET row visibility | copy rows without hidden rows
// Common Searches: Aspose.Cells copy only visible rows C# | How to skip hidden rows when using CopyRows | Exclude hidden rows in Aspose.Cells .NET | Copy rows from one worksheet to another ignoring hidden rows | Aspose.Cells IsRowHidden usage
// Developer Intent: Copy rows that are not hidden from a source worksheet to a destination worksheet.
// Use Cases: Create a report that includes only rows the user left visible after filtering or manual hiding. | Export a clean data set to a new file while preserving the original workbook's hidden rows. | Share a worksheet's visible content with collaborators without exposing hidden information.
// AI Prompts: Generate C# code using Aspose.Cells that copies only visible rows from one worksheet to another, keeping formatting intact. | Explain how to filter out hidden rows before invoking CopyRows in Aspose.Cells for .NET. | Show a step‑by‑step example of using IsRowHidden with CopyRows to skip hidden rows during a copy operation.

using System;
using Aspose.Cells;

// Demonstrates how to hide specific rows, detect hidden rows using IsRowHidden, and copy only the visible rows from a source worksheet to a new workbook with the CopyRows method, then save the result as VisibleRowsCopy.xlsx.
class ExcludeHiddenRowsCopy
{
    static void Main()
    {
        // Create source workbook and add sample data
        Workbook sourceWb = new Workbook();
        Worksheet srcSheet = sourceWb.Worksheets[0];
        Cells srcCells = srcSheet.Cells;

        for (int i = 0; i < 10; i++)
        {
            srcCells[i, 0].PutValue($"Row {i + 1}");
        }

        // Hide some rows (zero‑based indexes 1, 3, 5)
        srcCells.HideRow(1);
        srcCells.HideRow(3);
        srcCells.HideRow(5);

        // Create destination workbook
        Workbook destWb = new Workbook();
        Worksheet destSheet = destWb.Worksheets[0];
        Cells destCells = destSheet.Cells;

        int destRowIndex = 0;

        // Iterate through source rows, copy only visible rows
        for (int srcRowIndex = 0; srcRowIndex <= srcCells.MaxDataRow; srcRowIndex++)
        {
            // Skip hidden rows
            if (srcCells.IsRowHidden(srcRowIndex))
                continue;

            // Copy a single visible row from source to destination
            destCells.CopyRows(srcCells, srcRowIndex, destRowIndex, 1);
            destRowIndex++;
        }

        // Save the destination workbook
        destWb.Save("VisibleRowsCopy.xlsx");
    }
}
