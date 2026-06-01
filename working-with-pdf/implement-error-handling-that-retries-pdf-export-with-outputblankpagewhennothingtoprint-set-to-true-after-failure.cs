using System;
using Aspose.Cells;

namespace AsposeCellsPdfExportRetry
{
    class Program
    {
        static void Main()
        {
            // Path to the source Excel file (replace with actual path if needed)
            string excelPath = "input.xlsx";

            // Path for the output PDF file
            string pdfPath = "output.pdf";

            // Load the workbook (using the standard load rule)
            Workbook workbook = new Workbook(excelPath);

            // Create PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Initially try without forcing a blank page when nothing is printed
                OutputBlankPageWhenNothingToPrint = false,

                // Optional: hide errors during rendering to avoid immediate failure
                IgnoreError = true
            };

            try
            {
                // First attempt to save the workbook as PDF
                workbook.Save(pdfPath, pdfOptions);
                Console.WriteLine("PDF saved successfully on first attempt.");
            }
            catch (Exception ex)
            {
                // Log the error
                Console.WriteLine($"First PDF export failed: {ex.Message}");

                // Retry with OutputBlankPageWhenNothingToPrint set to true
                pdfOptions.OutputBlankPageWhenNothingToPrint = true;

                try
                {
                    workbook.Save(pdfPath, pdfOptions);
                    Console.WriteLine("PDF saved successfully on retry with OutputBlankPageWhenNothingToPrint = true.");
                }
                catch (Exception retryEx)
                {
                    // If it still fails, report the error
                    Console.WriteLine($"Retry PDF export also failed: {retryEx.Message}");
                }
            }
        }
    }
}