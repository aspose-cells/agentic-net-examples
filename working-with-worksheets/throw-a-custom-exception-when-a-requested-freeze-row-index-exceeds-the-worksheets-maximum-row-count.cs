// Title: Throw custom FreezeRowOutOfRangeException when freeze row index exceeds worksheet max rows – Aspose.Cells for .NET
// Description: Demonstrates how to read the worksheet's maximum row index via Workbook.Settings.MaxRow, compare it with a user‑supplied freeze row index, and raise a custom FreezeRowOutOfRangeException if the index is out of range. If validation succeeds, FreezePanes is applied and the workbook is saved.
// Keywords: Aspose.Cells freeze panes validation | custom exception FreezeRowOutOfRangeException | Settings.MaxRow Aspose.Cells | C# validate freeze row index | worksheet max row limit | Aspose.Cells .NET error handling
// Common Searches: Aspose.Cells validate freeze row index | throw custom exception when freeze row exceeds max rows | Settings.MaxRow usage in C# | how to check freeze pane limits Aspose.Cells | custom FreezeRowOutOfRangeException example
// Developer Intent: Check a requested freeze row index against the worksheet's maximum row count and throw a custom exception if it exceeds the limit.
// Use Cases: Prevent runtime errors by validating freeze pane parameters before calling FreezePanes. | Provide developers with a clear, domain‑specific exception when users request an invalid freeze row. | Integrate row‑limit validation into automated spreadsheet generation pipelines that use Aspose.Cells.
// AI Prompts: Generate C# code that validates a freeze row index with Aspose.Cells Settings.MaxRow and throws a custom FreezeRowOutOfRangeException. | Show how to catch FreezeRowOutOfRangeException around FreezePanes and log a detailed error message. | Extend the example to also validate freeze column index using Settings.MaxColumn and raise a corresponding custom exception.

using System;
using Aspose.Cells;

namespace FreezeRowValidationDemo
{
    // Custom exception to indicate that the requested freeze row index is out of range
    // Demonstrates how to read the worksheet's maximum row index via Workbook.Settings.MaxRow, compare it with a user‑supplied freeze row index, and raise a custom FreezeRowOutOfRangeException if the index is out of range. If validation succeeds, FreezePanes is applied and the workbook is saved.
    public class FreezeRowOutOfRangeException : Exception
    {
        public FreezeRowOutOfRangeException(string message) : base(message) { }
    }

    class Program
    {
        static void Main()
        {
            // Create a new workbook (uses Aspose.Cells create rule)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Example: requested freeze row index (zero‑based)
            int requestedFreezeRow = 70000;   // Change this value to test different scenarios
            int requestedFreezeColumn = 2;   // Example column index
            int freezedRows = requestedFreezeRow + 1;   // Number of rows to freeze (must be <= requestedFreezeRow)
            int freezedColumns = requestedFreezeColumn + 1; // Number of columns to freeze (must be <= requestedFreezeColumn)

            // Validate the requested freeze row against the workbook's maximum row index
            int maxRowIndex = workbook.Settings.MaxRow; // MaxRow is zero‑based

            if (requestedFreezeRow > maxRowIndex)
            {
                // Throw custom exception if the index exceeds the allowed range
                throw new FreezeRowOutOfRangeException(
                    $"Requested freeze row index ({requestedFreezeRow}) exceeds the worksheet's maximum row index ({maxRowIndex}).");
            }

            // If validation passes, apply the freeze panes
            worksheet.FreezePanes(requestedFreezeRow, requestedFreezeColumn, freezedRows, freezedColumns);

            // Save the workbook (uses Aspose.Cells save rule)
            workbook.Save("FreezeRowValidationResult.xlsx");

            Console.WriteLine("Workbook saved successfully.");
        }
    }
}
