// Title: C# – Handle Invalid Custom Paper Size with Aspose.Cells PageSetup Exception Handling
// Description: Shows how to catch a CellsException when Worksheet.PageSetup.CustomPaperSize receives negative or zero dimensions, log a descriptive error, and still save the workbook using the default paper size.
// Keywords: Aspose.Cells | C# | custom paper size | exception handling | CellsException | PageSetup | .NET | invalid dimensions | error logging | Workbook.Save
// Common Searches: Aspose.Cells catch CellsException custom paper size | invalid custom paper dimensions Aspose.Cells C# | how to log page setup errors in Aspose.Cells | exception handling for Worksheet.PageSetup.CustomPaperSize | fallback to default paper size when custom size fails Aspose.Cells
// Developer Intent: Add try‑catch logic around Worksheet.PageSetup.CustomPaperSize to detect invalid width/height values, log clear error messages, and ensure the workbook is saved.
// Use Cases: Validate width and height before calling CustomPaperSize; if invalid, log and continue with default size. | Catch CellsException filtered by ExceptionType.PageSetup to differentiate page‑setup errors from other failures. | Save the workbook after handling the exception so the file is still generated with standard paper settings.
// AI Prompts: Generate C# code that validates custom paper width and height before calling Worksheet.PageSetup.CustomPaperSize and logs a CellsException with ExceptionType.PageSetup. | Provide an Aspose.Cells example that sets a custom paper size, includes exception handling for invalid dimensions, and falls back to the default size. | Explain how to distinguish page‑setup errors from other exceptions when using CustomPaperSize in Aspose.Cells for .NET.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Shows how to catch a CellsException when Worksheet.PageSetup.CustomPaperSize receives negative or zero dimensions, log a descriptive error, and still save the workbook using the default paper size.
    public class CustomPaperSizeExceptionHandling
    {
        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            try
            {
                // Example of invalid dimensions (negative or zero values)
                double widthInInches = -1.0;   // Invalid width
                double heightInInches = 0.0;   // Invalid height

                // Attempt to set a custom paper size; this may throw a CellsException
                worksheet.PageSetup.CustomPaperSize(widthInInches, heightInInches);

                Console.WriteLine($"Custom paper size set to {widthInInches} x {heightInInches} inches.");
            }
            catch (CellsException ex) when (ex.Code == ExceptionType.PageSetup)
            {
                // Handle specific page setup errors and log a descriptive message
                Console.WriteLine($"PageSetup error (Code {ex.Code}): {ex.Message}");
            }
            catch (Exception ex)
            {
                // Handle any other unexpected errors
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }

            // Save the workbook (will use default paper size if custom size failed)
            workbook.Save("CustomPaperSizeHandled.xlsx");
        }

        // Entry point for the application
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Fatal error: {ex.Message}");
            }
        }
    }
}
