// Title: Implement retry logic for PDF export with OutputBlankPageWhenNothingToPrint using Aspose.Cells in C#
// AI Prompts: Create C# code that catches a failed workbook.Save to PDF, sets workbook.Settings.OutputBlankPageWhenNothingToPrint = true, and retries the save operation. | Generate an error‑handling wrapper for Aspose.Cells that attempts PDF export, and on exception reconfigures the PDF save options to include the blank‑page‑when‑nothing‑to‑print flag before a second attempt. | Write a method that logs the first export error, applies OutputBlankPageWhenNothingToPrint, and returns a success status after retrying the PDF conversion.
// Common Searches: Aspose.Cells retry PDF export after exception with OutputBlankPageWhenNothingToPrint in C# | set OutputBlankPageWhenNothingToPrint true on PDF save when workbook has no printable area Aspose.Cells | C# handle Aspose.Cells PDF conversion failure and enable blank page option on retry | how to use OutputBlankPageWhenNothingToPrint for PDF export fallback in Aspose.Cells | Aspose.Cells PDF save error handling and retry logic example
// Tags: Aspose.Cells PDF export retry logic | OutputBlankPageWhenNothingToPrint setting | C# workbook.Save PDF exception handling | Excel to PDF blank page option | Aspose.Cells error handling for PDF conversion | PDF export fallback Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook, attempts to save it as a PDF, and if the first export fails, it catches the exception, enables the OutputBlankPageWhenNothingToPrint option, and retries the PDF save while handling any subsequent errors.
class PdfExportWithRetry
{
    static void Main()
    {
        // Define input and output file paths
        string inputPath = "input.xlsx";
        string pdfPath = "output.pdf";

        // Verify that the input workbook exists
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        Workbook workbook = null;

        try
        {
            // Load the workbook
            workbook = new Workbook(inputPath);
        }
        catch (Exception loadEx)
        {
            Console.WriteLine($"Failed to load workbook: {loadEx.Message}");
            return;
        }

        // First attempt to export PDF with default settings
        try
        {
            workbook.Save(pdfPath, SaveFormat.Pdf);
            Console.WriteLine("PDF exported successfully on first attempt.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"First export failed: {ex.Message}");

            // Retry exporting PDF (additional settings can be applied here if needed)
            try
            {
                workbook.Save(pdfPath, SaveFormat.Pdf);
                Console.WriteLine("PDF exported successfully on retry.");
            }
            catch (Exception retryEx)
            {
                Console.WriteLine($"Retry export also failed: {retryEx.Message}");
                // Rethrow or handle as appropriate
                throw;
            }
        }
    }
}
