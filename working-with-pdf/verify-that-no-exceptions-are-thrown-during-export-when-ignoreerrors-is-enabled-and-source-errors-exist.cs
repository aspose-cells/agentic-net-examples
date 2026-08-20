// Title: Aspose.Cells .NET: Export PDF without exceptions using PdfSaveOptions.IgnoreError for formula errors
// Description: Shows how to create a workbook with an invalid formula, enable IgnoreError in both calculation and PDF save options, and save to PDF without any thrown exceptions.
// Keywords: Aspose.Cells | .NET | PdfSaveOptions | IgnoreError | formula error handling | PDF export | exception‑free conversion | calculate formula ignore errors | error‑tolerant PDF generation
// Common Searches: Aspose.Cells export PDF ignore errors example | PdfSaveOptions.IgnoreError .NET | how to suppress formula errors during PDF conversion Aspose.Cells | prevent exceptions when saving Excel to PDF with invalid formulas | ignore calculation errors Aspose.Cells PDF
// Developer Intent: Confirm that enabling IgnoreError stops exceptions from being raised when exporting a workbook containing formula errors to PDF.
// Use Cases: Generate PDF reports from spreadsheets that may contain user‑entered or legacy formulas that are no longer valid. | Batch‑process large numbers of Excel files in a service where some files have broken formulas, without interrupting the conversion pipeline. | Provide a web‑based upload feature that converts user files to PDF while gracefully handling any formula errors.
// AI Prompts: Write C# code using Aspose.Cells to export a workbook to PDF with PdfSaveOptions.IgnoreError set to true and log hidden errors. | Create a unit test in C# that asserts no exception is thrown when saving a workbook containing a #NAME? formula using PdfSaveOptions.IgnoreError. | Explain the difference between CalculationOptions.IgnoreError and PdfSaveOptions.IgnoreError in Aspose.Cells and how they affect PDF rendering.

using System;
using Aspose.Cells;

namespace AsposeCellsIgnoreErrorDemo
{
    // Shows how to create a workbook with an invalid formula, enable IgnoreError in both calculation and PDF save options, and save to PDF without any thrown exceptions.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Insert a formula that will cause an error (#NAME? because the function does not exist)
            sheet.Cells["A1"].Formula = "=NONEXISTENTFUNC()";

            // Set calculation options to ignore errors during formula evaluation
            workbook.CalculateFormula(new CalculationOptions { IgnoreError = true });

            // Prepare PDF save options and enable IgnoreError to hide rendering errors
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // This property comes from PaginatedSaveOptions base class
                IgnoreError = true
            };

            // Export the workbook to PDF inside a try-catch block to verify no exception is thrown
            try
            {
                workbook.Save("ExportWithIgnoreError.pdf", pdfOptions);
                Console.WriteLine("Export completed successfully without throwing exceptions.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An unexpected exception occurred during export:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
