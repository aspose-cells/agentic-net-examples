// Title: Handle Invalid Custom Paper Size with Aspose.Cells PageSetup in C#
// Description: Shows how to wrap Worksheet.PageSetup.CustomPaperSize in try‑catch blocks, capture CellsException for page‑setup errors, log clear messages, and still save the workbook when width or height are zero or negative.
// Keywords: Aspose.Cells | C# custom paper size | PageSetup.CustomPaperSize | exception handling | CellsException | invalid dimensions | error logging | negative width height | Excel export | page layout errors
// Common Searches: Aspose.Cells set custom paper size error | C# catch CellsException page setup | validate paper dimensions before CustomPaperSize | log page setup failures Aspose.Cells | exception thrown for negative page size
// Developer Intent: Add robust error handling around custom paper size settings to prevent crashes and provide informative logs.
// Use Cases: User enters custom page size in a UI; code validates and handles invalid values gracefully. | Automated report generation calculates dimensions; fallback to default size when values are out of range. | Batch processing of many workbooks where some contain zero or negative sizes; continue processing without interruption.
// AI Prompts: Generate C# code that checks width and height before calling PageSetup.CustomPaperSize and writes an error to a log file if they are non‑positive. | Provide an example of catching CellsException with ExceptionType.PageSetup, distinguishing it from other exceptions, and recording the stack trace. | Create a reusable method SetCustomPaperSize(Worksheet ws, double w, double h) that returns a bool indicating success and logs detailed messages on failure.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Shows how to wrap Worksheet.PageSetup.CustomPaperSize in try‑catch blocks, capture CellsException for page‑setup errors, log clear messages, and still save the workbook when width or height are zero or negative.
class CustomPaperSizeDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        PageSetup pageSetup = worksheet.PageSetup;

        // Example of invalid dimensions (zero or negative values)
        double widthInInches = 0.0;
        double heightInInches = -1.0;

        try
        {
            // Attempt to set a custom paper size; this may throw a CellsException
            pageSetup.CustomPaperSize(widthInInches, heightInInches);
            Console.WriteLine($"Custom paper size set to {widthInInches} x {heightInInches} inches.");
        }
        catch (CellsException ex) when (ex.Code == ExceptionType.PageSetup)
        {
            // Handle specific page‑setup related errors
            Console.WriteLine($"PageSetup error: {ex.Message}");
        }
        catch (Exception ex)
        {
            // Handle any other unexpected errors
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }

        // Save the workbook (valid even if custom size was not applied)
        workbook.Save("CustomPaperSizeDemo.xlsx");
    }
}
