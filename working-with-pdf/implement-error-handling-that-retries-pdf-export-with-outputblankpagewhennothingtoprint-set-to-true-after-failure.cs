// Title: Implement retry logic for PDF export with OutputBlankPageWhenNothingToPrint using Aspose.Cells in C#
// Common Searches: C# Aspose.Cells retry PDF save after exception with OutputBlankPageWhenNothingToPrint | How to catch workbook.Save failure and retry with blank page option in Aspose.Cells | Aspose.Cells PDF export fallback to blank page when nothing to print in C# | Exception handling for Aspose.Cells PdfSaveOptions OutputBlankPageWhenNothingToPrint | Retrying PDF generation in Aspose.Cells after first save fails
// Tags: Aspose.Cells PDF export retry logic | PdfSaveOptions OutputBlankPageWhenNothingToPrint usage | C# workbook.Save exception handling | fallback PDF generation with blank page | retry PDF save after failure Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The example creates a workbook, attempts to save it as a PDF, catches any exception, logs the error, and then retries the export using PdfSaveOptions with OutputBlankPageWhenNothingToPrint set to true, logging success or rethrowing if the retry also fails.
class PdfExportWithRetry
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook();               // empty workbook
        // Workbook workbook = new Workbook("input.xlsx"); // load from file if needed

        // First attempt to export PDF with default settings
        try
        {
            PdfSaveOptions options = new PdfSaveOptions();
            // By default OutputBlankPageWhenNothingToPrint is false
            workbook.Save("output.pdf", options);
        }
        catch (Exception ex)
        {
            // Log the failure of the first attempt
            Console.WriteLine("First PDF export failed: " + ex.Message);

            // Retry export with OutputBlankPageWhenNothingToPrint set to true
            try
            {
                PdfSaveOptions retryOptions = new PdfSaveOptions
                {
                    OutputBlankPageWhenNothingToPrint = true
                };
                workbook.Save("output.pdf", retryOptions);
                Console.WriteLine("PDF export succeeded on retry with blank page option.");
            }
            catch (Exception retryEx)
            {
                // Log the failure of the retry and rethrow if necessary
                Console.WriteLine("Retry PDF export also failed: " + retryEx.Message);
                throw;
            }
        }
    }
}
