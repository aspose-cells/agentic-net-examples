// Title: Catch CellsException When Exporting an Empty Workbook to PDF – Aspose.Cells C#
// Description: Demonstrates how to configure PdfSaveOptions to suppress blank pages, hide worksheets, and wrap the PDF export in a try‑catch block that specifically handles Aspose.Cells CellsException and generic errors.
// Keywords: Aspose.Cells PDF export | CellsException handling | C# empty worksheet PDF | PdfSaveOptions OutputBlankPageWhenNothingToPrint | ignore rendering errors Aspose.Cells | Aspose.Cells try catch example
// Common Searches: how to catch CellsException during PDF export | Aspose.Cells prevent blank page when nothing to print | export empty workbook to PDF without error | PdfSaveOptions ignore errors Aspose.Cells C# | exception thrown when all sheets are hidden Aspose.Cells
// Developer Intent: Identify and handle the CellsException that occurs when a workbook with no printable content is saved as a PDF.
// Use Cases: Log a warning when PDF generation yields no pages because all worksheets are hidden or empty. | Show a user-friendly message instead of a blank PDF file. | Prevent application crashes in automated report generation pipelines.
// AI Prompts: Write C# code that saves an Aspose.Cells workbook to PDF, disables blank‑page generation, and catches CellsException for custom logging. | Show how to set PdfSaveOptions.IgnoreError and OutputBlankPageWhenNothingToPrint for an empty workbook export. | Provide a try‑catch example that distinguishes Aspose.Cells CellsException from other exceptions during PDF conversion.

using System;
using Aspose.Cells;

// Demonstrates how to configure PdfSaveOptions to suppress blank pages, hide worksheets, and wrap the PDF export in a try‑catch block that specifically handles Aspose.Cells CellsException and generic errors.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Add an empty worksheet first (must remain visible)
            int emptySheetIndex = workbook.Worksheets.Add();
            Worksheet emptySheet = workbook.Worksheets[emptySheetIndex];
            // No data is added; the sheet stays effectively empty

            // Hide the original worksheet to simulate a scenario with no printable content
            Worksheet originalSheet = workbook.Worksheets[0];
            originalSheet.IsVisible = false;

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Do not generate a blank page when there is nothing to print
                OutputBlankPageWhenNothingToPrint = false,
                // Hide rendering errors (optional, but useful when content is missing)
                IgnoreError = true
            };

            // Save the workbook as PDF
            workbook.Save("output.pdf", pdfOptions);
            Console.WriteLine("PDF saved successfully.");
        }
        catch (CellsException ex)
        {
            // Handle Aspose.Cells specific exceptions
            Console.WriteLine($"CellsException caught: {ex.Message}");
        }
        catch (Exception ex)
        {
            // Handle any other unexpected exceptions
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
