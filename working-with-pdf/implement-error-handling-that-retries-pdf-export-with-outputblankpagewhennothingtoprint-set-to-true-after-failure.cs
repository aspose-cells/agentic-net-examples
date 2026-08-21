// Title: Retry PDF Export with OutputBlankPageWhenNothingToPrint in Aspose.Cells (C#)
// Description: The example creates a workbook containing a visible sheet and a hidden sheet, saves it to PDF with PdfSaveOptions.OutputBlankPageWhenNothingToPrint set to false, catches any export error, switches the option to true, and retries the save. It also ensures the output folder exists and logs both attempts.
// Keywords: Aspose.Cells | PDF export | OutputBlankPageWhenNothingToPrint | retry logic | C# | error handling | PdfSaveOptions | workbook.Save | hidden worksheet | .NET PDF generation
// Common Searches: Aspose.Cells retry PDF save after failure | OutputBlankPageWhenNothingToPrint true fallback | C# Aspose.Cells PDFSaveOptions exception handling | how to handle hidden sheets when exporting PDF with Aspose.Cells | Aspose.Cells PDF export blank page option
// Developer Intent: Add try‑catch handling that re‑saves a workbook as PDF with OutputBlankPageWhenNothingToPrint enabled if the initial save throws an exception.
// Use Cases: Generate PDF reports where hidden worksheets may cause a rendering error; fallback to a blank‑page option on the second attempt. | Run batch PDF conversions in a scheduled service and automatically recover from intermittent save failures. | Provide a resilient PDF export endpoint in a web API that logs the first error and retries with adjusted PDF options.
// AI Prompts: Create a reusable C# method that attempts workbook.Save to PDF, catches failures, toggles OutputBlankPageWhenNothingToPrint, and retries up to a configurable number of times. | Show how to log detailed exception information (stack trace, workbook name, attempt number) before retrying the PDF export with Aspose.Cells. | Write unit tests that simulate a failure on the first PDF save and verify that the retry succeeds when OutputBlankPageWhenNothingToPrint is set to true.

using System;
using System.IO;
using Aspose.Cells;

// The example creates a workbook containing a visible sheet and a hidden sheet, saves it to PDF with PdfSaveOptions.OutputBlankPageWhenNothingToPrint set to false, catches any export error, switches the option to true, and retries the save. It also ensures the output folder exists and logs both attempts.
class PdfExportWithRetry
{
    static void Main()
    {
        try
        {
            // Create a new workbook and add data to a visible worksheet
            Workbook workbook = new Workbook();
            Worksheet visibleSheet = workbook.Worksheets[0];
            visibleSheet.Name = "VisibleSheet";
            visibleSheet.Cells["A1"].PutValue("Sample data for PDF export");

            // Add a second worksheet that will be hidden to simulate a rendering issue
            Worksheet hiddenSheet = workbook.Worksheets.Add("HiddenSheet");
            hiddenSheet.Cells["A1"].PutValue("Sample data for hidden sheet");
            hiddenSheet.IsVisible = false; // Hide this sheet

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Initially set to false to demonstrate the retry scenario
                OutputBlankPageWhenNothingToPrint = false
            };

            string outputFile = "ExportedDocument.pdf";

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputFile));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            try
            {
                // First attempt to save the workbook as PDF
                workbook.Save(outputFile, pdfOptions);
                Console.WriteLine("PDF saved successfully on the first attempt.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"First save attempt failed: {ex.Message}");

                // Retry with OutputBlankPageWhenNothingToPrint set to true
                pdfOptions.OutputBlankPageWhenNothingToPrint = true;

                try
                {
                    workbook.Save(outputFile, pdfOptions);
                    Console.WriteLine("PDF saved successfully on retry with OutputBlankPageWhenNothingToPrint = true.");
                }
                catch (Exception retryEx)
                {
                    Console.WriteLine($"Retry also failed: {retryEx.Message}");
                    // Additional error handling can be placed here
                }
            }
        }
        catch (Exception outerEx)
        {
            Console.WriteLine($"Unexpected error: {outerEx.Message}");
        }
    }
}
