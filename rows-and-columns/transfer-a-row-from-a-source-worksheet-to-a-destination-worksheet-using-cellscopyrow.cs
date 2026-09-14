// Title: How to copy a specific row from one worksheet to another using Aspose.Cells Cells.CopyRow in C#
// AI Prompts: Use Aspose.Cells in C# to copy row 0 from a worksheet named "SourceSheet" to a newly created worksheet "DestinationSheet" and save the workbook as an .xlsx file. | Write C# code that transfers a given row index from a source worksheet to a target worksheet in the same workbook using Cells.CopyRow, then persist the changes. | Adapt the sample to copy multiple consecutive rows from the source sheet to the destination sheet with Aspose.Cells for .NET.
// Common Searches: aspnet copy first row from one sheet to another using Aspose.Cells | C# Aspose.Cells Cells.CopyRow method example for transferring rows | how to duplicate a row between worksheets in an Excel file with Aspose.Cells .NET | copy row between worksheets and save workbook using Aspose.Cells C#
// Tags: Aspose.Cells Cells.CopyRow usage C# | copy row between worksheets Aspose.Cells | programmatic Excel row transfer .NET | save workbook after copying row Aspose | duplicate row to new worksheet Aspose.Cells

using System;
using Aspose.Cells;

namespace AsposeCellsRowTransfer
{
    // The example creates a workbook, adds sample data to the first row of a source worksheet, creates a destination worksheet, copies the specified row from the source to the destination using Cells.CopyRow, and saves the result as RowTransferResult.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook which will contain both source and destination worksheets
            Workbook workbook = new Workbook();

            // -------------------- Source Worksheet --------------------
            // Access the first worksheet (default) and add sample data to the first row
            Worksheet sourceSheet = workbook.Worksheets[0];
            sourceSheet.Name = "SourceSheet";
            Cells sourceCells = sourceSheet.Cells;
            sourceCells["A1"].PutValue("Item");
            sourceCells["B1"].PutValue(123);
            sourceCells["C1"].PutValue(DateTime.Now);

            // -------------------- Destination Worksheet --------------------
            // Add a new worksheet to act as the destination
            Worksheet destSheet = workbook.Worksheets.Add("DestinationSheet");
            Cells destCells = destSheet.Cells;

            // Copy the entire first row from source to destination (row index 0 -> 0)
            destCells.CopyRow(sourceCells, 0, 0);

            // -------------------- Save the Workbook --------------------
            // Save the result to an Excel file
            workbook.Save("RowTransferResult.xlsx");
        }
    }
}
