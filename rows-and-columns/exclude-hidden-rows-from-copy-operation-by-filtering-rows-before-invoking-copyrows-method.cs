// Title: Copy Only Visible Rows with Aspose.Cells for .NET (C#)
// Description: Loads a source workbook, iterates through each row up to the last data row, skips rows where IsRowHidden returns true, copies the remaining rows to a new worksheet using CopyRow, and saves the result as a separate file.
// Keywords: Aspose.Cells | CopyRow | visible rows | skip hidden rows | C# Excel automation | filter rows before copy | exclude hidden rows | Aspose.Cells .NET | Excel row copy | CopyRows alternative
// Common Searches: Aspose.Cells copy only visible rows C# | skip hidden rows when copying Excel with Aspose | filter hidden rows before CopyRows Aspose.Cells | copy visible rows to new workbook using Aspose.Cells | C# code to ignore hidden rows in Excel export
// Developer Intent: Copy all non‑hidden rows from a source worksheet to a destination workbook using Aspose.Cells.
// Use Cases: Create a clean report that contains only rows the user can see after manual hiding or filter operations. | Export a lightweight Excel file for downstream processing, omitting rows that are hidden for readability. | Synchronize visible data between two workbooks while preserving original row order.
// AI Prompts: Generate C# code with Aspose.Cells that copies only visible rows from one worksheet to another, ignoring hidden rows. | Explain how IsRowHidden and CopyRow can be combined to filter out hidden rows before a bulk copy in Aspose.Cells. | Show how to maintain correct destination row indexing when copying only visible rows with Aspose.Cells for .NET.

using System;
using Aspose.Cells;

namespace AsposeCellsCopyVisibleRows
{
    // Loads a source workbook, iterates through each row up to the last data row, skips rows where IsRowHidden returns true, copies the remaining rows to a new worksheet using CopyRow, and saves the result as a separate file.
    class Program
    {
        static void Main()
        {
            // Load the source workbook (replace with actual path)
            Workbook sourceWorkbook = new Workbook("source.xlsx");
            Worksheet sourceSheet = sourceWorkbook.Worksheets[0];
            Cells sourceCells = sourceSheet.Cells;

            // Create a new workbook for the destination
            Workbook destinationWorkbook = new Workbook();
            Worksheet destinationSheet = destinationWorkbook.Worksheets[0];
            Cells destinationCells = destinationSheet.Cells;

            // Destination row index that will be incremented only for visible rows
            int destRowIndex = 0;

            // Determine the last row that contains data in the source sheet
            int lastRow = sourceCells.MaxDataRow;

            // Iterate through each row in the source sheet
            for (int srcRowIndex = 0; srcRowIndex <= lastRow; srcRowIndex++)
            {
                // Skip hidden rows
                if (sourceCells.IsRowHidden(srcRowIndex))
                    continue;

                // Copy the visible row to the destination sheet
                // Using CopyRow (copies a single row) for simplicity
                destinationCells.CopyRow(sourceCells, srcRowIndex, destRowIndex);

                // Move to the next destination row
                destRowIndex++;
            }

            // Save the result (replace with desired output path)
            destinationWorkbook.Save("output.xlsx", SaveFormat.Xlsx);
        }
    }
}
