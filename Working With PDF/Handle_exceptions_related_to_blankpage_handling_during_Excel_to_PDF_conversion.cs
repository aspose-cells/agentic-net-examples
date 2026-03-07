using System;
using Aspose.Cells;

class BlankPageHandlingDemo
{
    static void Main()
    {
        // Create a new workbook and add some sample data
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Header");
        sheet.Cells["A2"].PutValue("Row 1");
        sheet.Cells["A3"].PutValue("Row 2");

        // Add a hidden worksheet to simulate a sheet with nothing to print
        Worksheet hidden = workbook.Worksheets.Add("HiddenSheet");
        hidden.IsVisible = false;

        // Configure PDF save options:
        // - IgnoreError hides rendering errors (shapes, charts, etc.)
        // - OutputBlankPageWhenNothingToPrint prevents a blank page for empty sheets
        // - PrintingPageType.IgnoreBlank skips completely blank pages in the PDF
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            IgnoreError = true,
            OutputBlankPageWhenNothingToPrint = false,
            PrintingPageType = PrintingPageType.IgnoreBlank
        };

        try
        {
            // Save the workbook as PDF using the configured options
            workbook.Save("Output.pdf", pdfOptions);
            Console.WriteLine("PDF saved successfully.");
        }
        catch (CellsException ex) when (ex.Code == ExceptionType.UnsupportedFeature)
        {
            // Handle specific unsupported feature errors that could cause blank pages
            Console.WriteLine("Unsupported feature encountered: " + ex.Message);
        }
        catch (CellsException ex)
        {
            // General Aspose.Cells exception handling
            Console.WriteLine($"CellsException (Code {ex.Code}): {ex.Message}");
        }
        catch (Exception ex)
        {
            // Fallback for any other unexpected errors
            Console.WriteLine("Unexpected error: " + ex.Message);
        }
    }
}