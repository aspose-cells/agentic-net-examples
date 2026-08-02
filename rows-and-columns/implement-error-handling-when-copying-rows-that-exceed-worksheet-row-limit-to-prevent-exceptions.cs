// Title: Safely Copy Rows Near Worksheet MaxRow Limit with Aspose.Cells for .NET
// Description: Demonstrates how to copy rows from a source worksheet to a destination worksheet while checking the destination's MaxRow setting. The example adjusts the copy range when the requested rows would exceed the worksheet limit, logs the adjustment, and saves the workbook without throwing an exception.
// Keywords: Aspose.Cells copy rows | row limit handling | MaxRow exception | safe row copy .NET | prevent worksheet overflow
// Common Searches: Aspose.Cells copy rows without exceeding max rows | how to avoid MaxRow exception in Aspose.Cells | safe row copy method Aspose.Cells C# | adjust row copy count for worksheet limit
// Developer Intent: Implement row‑copy logic that validates the destination worksheet’s maximum row count, trims the copy range if necessary, and prevents runtime exceptions.
// Use Cases: Copy a block of rows when the destination start row is close to the worksheet’s maximum row index. | Automatically truncate the copy operation to fit within the allowed row range. | Skip copying and log a warning if the start row is already beyond the worksheet limit.
// AI Prompts: Create a reusable C# method using Aspose.Cells that copies rows safely, checks the destination MaxRow, adjusts the row count, and returns the number of rows actually copied. | Provide C# code that wraps Aspose.Cells’ CopyRows with error handling for row overflow, includes logging, and optionally falls back to a no‑copy scenario.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Demonstrates how to copy rows from a source worksheet to a destination worksheet while checking the destination's MaxRow setting. The example adjusts the copy range when the requested rows would exceed the worksheet limit, logs the adjustment, and saves the workbook without throwing an exception.
    public class SafeCopyRowsDemo
    {
        public static void Main()
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a source workbook and add sample rows
            Workbook sourceWb = new Workbook();
            Worksheet srcSheet = sourceWb.Worksheets[0];
            for (int i = 0; i < 10; i++)
            {
                srcSheet.Cells[i, 0].PutValue($"Source Row {i + 1}");
            }

            // Create a destination workbook
            Workbook destWb = new Workbook();

            // Define copy parameters – intentionally set destination near the row limit
            int sourceStartRow = 0;
            int destinationStartRow = destWb.Settings.MaxRow - 5; // close to the maximum allowed row
            int rowsToCopy = 10; // this would exceed the limit without handling

            // Perform safe copy
            SafeCopyRows(
                sourceWorksheet: srcSheet,
                sourceRowIndex: sourceStartRow,
                destinationWorksheet: destWb.Worksheets[0],
                destinationRowIndex: destinationStartRow,
                rowNumber: rowsToCopy);

            // Save the result
            destWb.Save("SafeCopyRowsOutput.xlsx");
        }

        private static void SafeCopyRows(Worksheet sourceWorksheet, int sourceRowIndex,
                                         Worksheet destinationWorksheet, int destinationRowIndex,
                                         int rowNumber)
        {
            // Maximum row index allowed by the destination workbook/worksheet format
            int maxRowIndex = destinationWorksheet.Workbook.Settings.MaxRow;

            // Determine the last row that would be written by the copy operation
            int lastRowToWrite = destinationRowIndex + rowNumber - 1;

            if (lastRowToWrite > maxRowIndex)
            {
                // Calculate how many rows can actually be copied without exceeding the limit
                int allowedRows = maxRowIndex - destinationRowIndex + 1;

                if (allowedRows <= 0)
                {
                    Console.WriteLine("Destination start row is beyond the worksheet limit. No rows will be copied.");
                    return;
                }

                Console.WriteLine($"Requested copy exceeds row limit. Adjusting rows to copy from {rowNumber} to {allowedRows}.");
                rowNumber = allowedRows;
            }

            // Perform the copy using the built‑in CopyRows method
            destinationWorksheet.Cells.CopyRows(sourceWorksheet.Cells, sourceRowIndex, destinationRowIndex, rowNumber);
        }
    }
}
