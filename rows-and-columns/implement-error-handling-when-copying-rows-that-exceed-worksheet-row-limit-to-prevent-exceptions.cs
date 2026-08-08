// Title: Safely copy rows with row‑limit validation and error handling in Aspose.Cells for .NET
// Description: Demonstrates how to copy a range of rows from one worksheet to another while checking the workbook's maximum row index (Settings.MaxRow). The example automatically trims the row count to stay within XLS/XLSX limits, wraps the CopyRows and Save calls in try‑catch blocks, and logs adjustments or exceptions.
// Keywords: Aspose.Cells CopyRows | row limit validation | Settings.MaxRow | C# Excel row overflow protection | exception handling Aspose.Cells | safe row copy | Excel worksheet max rows
// Common Searches: Aspose.Cells prevent row overflow when copying rows | C# copy rows near bottom of worksheet without exception | How to use Settings.MaxRow in Aspose.Cells | Wrap Aspose.Cells CopyRows in try catch | Adjust rowsToCopy based on worksheet size Aspose
// Developer Intent: Copy rows without exceeding the worksheet's row capacity and handle any runtime errors gracefully.
// Use Cases: Copy a block of rows that starts close to the sheet's bottom, automatically truncating the copy to fit the format's row limit. | Programmatically determine the maximum row index for XLS or XLSX files and validate copy operations before execution. | Log adjustments or failures during row copying and workbook saving to aid debugging and user feedback.
// AI Prompts: Generate C# code that copies rows with Aspose.Cells, checks Settings.MaxRow, and reduces the copy size if it would exceed the limit. | Show how to wrap Aspose.Cells CopyRows and Save methods in try‑catch blocks that output error messages. | Explain the steps to retrieve the maximum row index for a workbook format using Aspose.Cells and use it to prevent CopyRows exceptions.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Demonstrates how to copy a range of rows from one worksheet to another while checking the workbook's maximum row index (Settings.MaxRow). The example automatically trims the row count to stay within XLS/XLSX limits, wraps the CopyRows and Save calls in try‑catch blocks, and logs adjustments or exceptions.
    public class SafeCopyRowsDemo
    {
        public static void Run()
        {
            // Create a source workbook and add sample data
            Workbook sourceWorkbook = new Workbook();
            Worksheet sourceSheet = sourceWorkbook.Worksheets[0];
            for (int i = 0; i < 10; i++)
            {
                sourceSheet.Cells[i, 0].PutValue($"Source Row {i + 1}");
            }

            // Create a destination workbook
            Workbook destWorkbook = new Workbook();
            Worksheet destSheet = destWorkbook.Worksheets[0];

            // Define copy parameters
            int sourceStartRow = 0;               // zero‑based index of first row to copy
            int destinationStartRow = 65000;      // intentionally near the limit to trigger handling
            int rowsToCopy = 100;                 // number of rows we want to copy

            // Get the maximum allowed row index for the workbook format
            int maxRowIndex = destWorkbook.Settings.MaxRow; // e.g., 65535 for XLS, 1048575 for XLSX

            // Calculate the last row index after copying
            int lastRowIndexAfterCopy = destinationStartRow + rowsToCopy - 1;

            // Adjust rowsToCopy if it would exceed the worksheet limit
            if (lastRowIndexAfterCopy > maxRowIndex)
            {
                // Reduce the number of rows to copy so that we stay within the limit
                rowsToCopy = maxRowIndex - destinationStartRow + 1;
                Console.WriteLine($"Adjusted rows to copy to {rowsToCopy} to avoid exceeding the max row limit ({maxRowIndex}).");
            }

            // Perform the copy inside a try‑catch block to handle any unexpected errors
            try
            {
                // Use the CopyRows method (sourceCells, sourceRowIndex, destinationRowIndex, rowNumber)
                destSheet.Cells.CopyRows(
                    sourceSheet.Cells,
                    sourceStartRow,
                    destinationStartRow,
                    rowsToCopy);

                Console.WriteLine("Rows copied successfully.");
            }
            catch (Exception ex)
            {
                // Log the exception – in a real application you might rethrow or handle differently
                Console.WriteLine($"Error during row copy: {ex.Message}");
            }

            // Save the result (using the standard Save method)
            try
            {
                destWorkbook.Save("SafeCopyRowsOutput.xlsx");
                Console.WriteLine("Workbook saved as SafeCopyRowsOutput.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error saving workbook: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            try
            {
                SafeCopyRowsDemo.Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unhandled exception: {ex.Message}");
            }
        }
    }
}
