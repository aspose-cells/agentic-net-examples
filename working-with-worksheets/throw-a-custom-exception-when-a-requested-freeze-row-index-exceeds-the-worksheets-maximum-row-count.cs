// Title: Throw a custom FreezeRowIndexOutOfRangeException when a requested freeze row index exceeds Excel's maximum row count using Aspose.Cells for .NET
// AI Prompts: Generate C# code that declares a FreezeRowIndexOutOfRangeException class, verifies a supplied freeze row index against the worksheet's maximum rows, and raises the exception before invoking Worksheet.FreezePanes. | Refactor the example to obtain the worksheet's maximum row count from the Aspose.Cells API instead of a hard‑coded constant, while keeping the custom exception behavior for out‑of‑range values. | Add structured logging that records the message from FreezeRowIndexOutOfRangeException and rethrows it wrapped in a higher‑level ApplicationException.
// Common Searches: how to validate freeze pane row index with Aspose.Cells before calling FreezePanes | C# custom exception for freeze row index out of range in Excel workbook | retrieve maximum number of rows supported by Aspose.Cells worksheet | prevent FreezePanes error when row index exceeds Excel limit using Aspose.Cells | Aspose.Cells .NET check row index against 1,048,576 limit
// Tags: freeze row index validation exception Aspose.Cells | freeze panes row limit validation .NET | worksheet row capacity Aspose.Cells | throw out‑of‑range exception before FreezePanes | C# Aspose.Cells FreezePanes parameter validation

using System;
using System.IO;
using Aspose.Cells;

namespace FreezeRowExample
{
    // Custom exception for invalid freeze row index
    // The sample creates a workbook, defines a requested freeze row index, checks it against Excel's 1,048,576 row limit, throws a custom FreezeRowIndexOutOfRangeException when the index is too high, applies FreezePanes, saves the file, and handles both the custom and generic exceptions.
    public class FreezeRowIndexOutOfRangeException : Exception
    {
        public FreezeRowIndexOutOfRangeException(string message) : base(message) { }
    }

    public class Program
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Example: requested freeze row index (0‑based)
                int requestedFreezeRowIndex = 1050000; // adjust as needed

                // Determine the maximum number of rows supported by the worksheet.
                // Aspose.Cells supports up to 1,048,576 rows (Excel limit).
                const int maxRowsSupported = 1048576;

                // Validate the requested index against the maximum row count
                if (requestedFreezeRowIndex >= maxRowsSupported)
                {
                    throw new FreezeRowIndexOutOfRangeException(
                        $"Requested freeze row index {requestedFreezeRowIndex} exceeds the worksheet's maximum row count of {maxRowsSupported - 1}.");
                }

                // Apply freeze panes.
                // FreezePanes(row, column, totalRows, totalColumns) freezes rows above 'row' and columns left of 'column'.
                sheet.FreezePanes(requestedFreezeRowIndex, 0, requestedFreezeRowIndex, 0);

                // Define output file path
                string outputPath = "FreezeRowResult.xlsx";

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (FreezeRowIndexOutOfRangeException ex)
            {
                Console.WriteLine($"Validation error: {ex.Message}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An unexpected error occurred: {ex.Message}");
            }
        }
    }
}
