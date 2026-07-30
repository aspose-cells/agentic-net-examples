// Title: Throw a custom FreezeRowOutOfRangeException when a freeze‑pane row exceeds the worksheet's max row in Aspose.Cells for .NET
// Description: The example creates a workbook, reads the maximum row index via Workbook.Settings.MaxRow, validates a requested freeze‑pane row, throws a user‑defined FreezeRowOutOfRangeException with a clear message if the index is out of bounds, and otherwise freezes the panes and saves the file.
// Keywords: Aspose.Cells | C# | FreezePanes | custom exception | FreezeRowOutOfRangeException | worksheet max row | Workbook.Settings.MaxRow | row index validation | Excel freeze pane error handling | Aspose.Cells .NET example
// Common Searches: validate freeze pane row index Aspose.Cells | custom exception for invalid freeze row in Aspose.Cells | maximum row count in Aspose.Cells workbook | C# freeze panes with range check | throw exception when freeze row exceeds max row Aspose.Cells
// Developer Intent: Check that the requested freeze‑pane row is within the worksheet limits and raise a meaningful error if it is not.
// Use Cases: Prevent runtime failures in reporting tools that programmatically freeze rows. | Provide end‑users with a precise error message when they attempt to freeze beyond the sheet size. | Integrate row‑index validation into automated Excel generation pipelines to enforce size constraints.
// AI Prompts: Generate C# code that checks a freeze‑pane row against Workbook.Settings.MaxRow and throws a custom FreezeRowOutOfRangeException with a detailed message. | Show how to catch FreezeRowOutOfRangeException in a console app, log the error, and continue processing other tasks. | Explain how to extend the sample to also validate a freeze‑pane column index using a similar custom exception.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Custom exception for invalid freeze row index
    // The example creates a workbook, reads the maximum row index via Workbook.Settings.MaxRow, validates a requested freeze‑pane row, throws a user‑defined FreezeRowOutOfRangeException with a clear message if the index is out of bounds, and otherwise freezes the panes and saves the file.
    public class FreezeRowOutOfRangeException : Exception
    {
        public FreezeRowOutOfRangeException(string message) : base(message) { }
    }

    public class FreezePanesWithValidation
    {
        public static void Run()
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Example: request a freeze row index beyond the worksheet's maximum row
            int requestedFreezeRow = workbook.Settings.MaxRow + 5; // intentionally out of range

            // Validate the requested row against the maximum allowed row index
            if (requestedFreezeRow > workbook.Settings.MaxRow)
            {
                // Throw custom exception if validation fails
                throw new FreezeRowOutOfRangeException(
                    $"Requested freeze row index {requestedFreezeRow} exceeds worksheet maximum row index {workbook.Settings.MaxRow}.");
            }

            // If validation passes, freeze panes (row, column, frozenRows, frozenColumns)
            worksheet.FreezePanes(requestedFreezeRow, 0, requestedFreezeRow, 0);

            // Save the workbook (lifecycle rule: save)
            workbook.Save("FreezePanesValidated.xlsx");
        }
    }

    class Program
    {
        static void Main()
        {
            try
            {
                FreezePanesWithValidation.Run();
            }
            catch (FreezeRowOutOfRangeException ex)
            {
                Console.WriteLine("Custom exception caught: " + ex.Message);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Unexpected error: " + ex.Message);
            }
        }
    }
}
