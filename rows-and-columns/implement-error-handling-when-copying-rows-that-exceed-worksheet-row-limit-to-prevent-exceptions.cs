// Title: Safely copy rows with Aspose.Cells .NET – avoid exceeding worksheet row limit
// Description: C# example that checks the destination worksheet's Settings.MaxRow, calculates the allowable row count, copies only the safe number of rows with Cells.CopyRows, handles overflow cases, and saves the result.
// Keywords: Aspose.Cells | C# CopyRows | row limit handling | Settings.MaxRow | worksheet overflow protection | safe row copy | Excel row capacity
// Common Searches: Aspose.Cells copy rows without exceeding max rows | How to prevent row limit exception in Aspose.Cells .NET | Check worksheet row capacity before copying data | CopyRows safe example Aspose.Cells | Handle Excel row overflow Aspose.Cells
// Developer Intent: Add logic that copies only the rows that fit within the destination worksheet’s maximum row count to prevent runtime exceptions.
// Use Cases: Migrate data between workbooks while respecting XLS/XLSX row limits. | Validate large data transfers in ETL pipelines to ensure the target sheet can accommodate the rows. | Create format‑agnostic utilities that adapt copying behavior based on Settings.MaxRow for different Excel versions.
// AI Prompts: Show me a C# method that safely copies rows between Aspose.Cells worksheets and returns the actual rows copied. | Generate code that checks Settings.MaxRow before calling Cells.CopyRows and logs a warning if rows are truncated. | Explain how to handle a destination worksheet that is near its row capacity when using Aspose.Cells for .NET.

using System;
using Aspose.Cells;

namespace AsposeCellsRowCopySafetyDemo
{
    // C# example that checks the destination worksheet's Settings.MaxRow, calculates the allowable row count, copies only the safe number of rows with Cells.CopyRows, handles overflow cases, and saves the result.
    class Program
    {
        static void Main()
        {
            // Create source workbook and fill it with sample data
            Workbook sourceWorkbook = new Workbook();
            Worksheet sourceSheet = sourceWorkbook.Worksheets[0];
            Cells sourceCells = sourceSheet.Cells;

            // Populate 10 rows of data
            for (int i = 0; i < 10; i++)
            {
                sourceCells[i, 0].PutValue($"Row {i + 1} - Col A");
                sourceCells[i, 1].PutValue($"Row {i + 1} - Col B");
            }

            // Create destination workbook (empty)
            Workbook destWorkbook = new Workbook();
            Worksheet destSheet = destWorkbook.Worksheets[0];
            Cells destCells = destSheet.Cells;

            // Parameters for copying
            int sourceRowIndex = 0;                     // start from first row in source
            int destinationRowIndex = 0;                // start from first row in destination
            int rowsToCopy = sourceCells.MaxRow + 1;    // total rows in source (MaxRow is zero‑based)

            // Determine the maximum row index allowed by the destination workbook format
            int maxAllowedRowIndex = destWorkbook.Settings.MaxRow; // zero‑based

            // Calculate how many rows can actually be copied without exceeding the limit
            int availableRows = maxAllowedRowIndex - destinationRowIndex + 1;
            int safeRowsToCopy = Math.Min(rowsToCopy, availableRows);

            if (safeRowsToCopy > 0)
            {
                // Perform the copy safely
                destCells.CopyRows(sourceCells, sourceRowIndex, destinationRowIndex, safeRowsToCopy);
                Console.WriteLine($"Copied {safeRowsToCopy} rows successfully.");
            }
            else
            {
                // No rows can be copied; handle according to business logic
                Console.WriteLine("Destination row index exceeds worksheet row limit. No rows were copied.");
            }

            // Save the result
            destWorkbook.Save("RowCopySafeResult.xlsx");
        }
    }
}
