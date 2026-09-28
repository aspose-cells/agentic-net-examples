// Title: How to copy rows safely in Aspose.Cells for .NET without exceeding the worksheet row limit
// AI Prompts: Generate C# code that copies a block of rows from a source worksheet to a destination worksheet using Aspose.Cells, first checking Workbook.Settings.MaxRow and trimming the row count if the copy would surpass the allowed rows. | Show an example of adding error handling around Cells.CopyRows in Aspose.Cells to gracefully handle cases where the destination start index is beyond the maximum row capacity.
// Common Searches: Aspose.Cells copy rows when destination exceeds max rows .NET | C# Aspose.Cells prevent exception for row overflow during CopyRows | adjust rowsToCopy based on Settings.MaxRow Aspose.Cells example | safe row copy between workbooks using Aspose.Cells C# | how to validate worksheet row limit before copying rows Aspose.Cells
// Tags: Aspose.Cells copy rows with max row check | C# Aspose.Cells row limit handling | Aspose.Cells Settings.MaxRow validation | Aspose.Cells safe row copy between worksheets | Aspose.Cells prevent row overflow exception

using System;
using Aspose.Cells;

namespace AsposeCellsRowCopyExample
{
    // The example creates a source workbook with sample data, then copies its rows to a destination workbook while checking the destination worksheet's Settings.MaxRow. If the copy would exceed the allowed rows, the code adjusts the number of rows to copy or aborts, preventing runtime exceptions. The result is saved as RowCopySafeResult.xlsx.
    class Program
    {
        static void Main()
        {
            // ---------- Create source workbook and add sample data ----------
            Workbook sourceWorkbook = new Workbook();
            Worksheet sourceSheet = sourceWorkbook.Worksheets[0];
            Cells sourceCells = sourceSheet.Cells;

            // Populate 10 rows of data in the source sheet
            for (int i = 0; i < 10; i++)
            {
                sourceCells[i, 0].PutValue($"Row {i + 1} - Col A");
                sourceCells[i, 1].PutValue($"Row {i + 1} - Col B");
            }

            // ---------- Create destination workbook ----------
            Workbook destWorkbook = new Workbook();
            Worksheet destSheet = destWorkbook.Worksheets[0];
            Cells destCells = destSheet.Cells;

            // Desired copy parameters
            int sourceRowIndex = 0;               // start copying from first row of source
            int destinationRowIndex = 0;          // start pasting at first row of destination
            int rowsToCopy = sourceCells.MaxRow + 1; // total rows that contain data in source

            // ---------- Error‑handling logic ----------
            // MaxRow is zero‑based; it returns the highest row index allowed by the file format.
            int maxAllowedRowIndex = destWorkbook.Settings.MaxRow;

            // Calculate the last row index that would be written after the copy operation
            int lastDestinationRowIndex = destinationRowIndex + rowsToCopy - 1;

            if (lastDestinationRowIndex > maxAllowedRowIndex)
            {
                // Adjust the number of rows to copy so we do not exceed the limit
                rowsToCopy = maxAllowedRowIndex - destinationRowIndex + 1;

                if (rowsToCopy <= 0)
                {
                    Console.WriteLine("Cannot copy rows: destination start index is beyond the worksheet row limit.");
                }
                else
                {
                    Console.WriteLine($"Adjusted rows to copy to {rowsToCopy} to stay within the row limit.");
                    destCells.CopyRows(sourceCells, sourceRowIndex, destinationRowIndex, rowsToCopy);
                }
            }
            else
            {
                // Safe to copy all requested rows
                destCells.CopyRows(sourceCells, sourceRowIndex, destinationRowIndex, rowsToCopy);
            }

            // ---------- Save the result ----------
            destWorkbook.Save("RowCopySafeResult.xlsx");
            Console.WriteLine("Workbook saved as RowCopySafeResult.xlsx");
        }
    }
}
