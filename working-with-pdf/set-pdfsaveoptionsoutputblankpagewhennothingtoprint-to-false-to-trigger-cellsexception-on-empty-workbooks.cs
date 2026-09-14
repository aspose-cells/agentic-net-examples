// Title: Trigger CellsException by saving an empty Aspose.Cells workbook to PDF with OutputBlankPageWhenNothingToPrint set to false
// AI Prompts: Generate C# code that creates a workbook without worksheets, sets PdfSaveOptions.OutputBlankPageWhenNothingToPrint = false, attempts to save to PDF, and catches the resulting CellsException. | Explain how to configure Aspose.Cells PDF export to suppress blank-page generation and cause an exception for empty workbooks in .NET.
// Common Searches: Aspose.Cells how to prevent blank page when exporting empty workbook to PDF | PdfSaveOptions OutputBlankPageWhenNothingToPrint false cause exception | C# catch CellsException for empty workbook PDF conversion | Saving empty workbook to PDF throws CellsException Aspose.Cells
// Tags: PdfSaveOptions disable blank page Aspose.Cells | empty workbook PDF export exception | CellsException handling Aspose.Cells PDF | Aspose.Cells PDF conversion error handling | C# Aspose.Cells output blank page setting

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// // Demonstrates creating an empty Workbook, configuring PdfSaveOptions to not output a blank page, attempting to save to PDF, and catching the resulting CellsException.
class Program
{
    static void Main()
    {
        // Create an empty workbook (no worksheets added)
        Workbook workbook = new Workbook();

        // Configure PDF save options to NOT output a blank page when nothing is printed
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        pdfOptions.OutputBlankPageWhenNothingToPrint = false; // This will cause a CellsException for empty workbooks

        try
        {
            // Attempt to save the empty workbook to PDF
            workbook.Save("EmptyWorkbook.pdf", pdfOptions);
            Console.WriteLine("PDF saved successfully (unexpected).");
        }
        catch (CellsException ex)
        {
            // Expected exception due to empty workbook and OutputBlankPageWhenNothingToPrint = false
            Console.WriteLine($"CellsException caught as expected: {ex.Message}");
        }
        catch (Exception ex)
        {
            // Any other unexpected exceptions
            Console.WriteLine($"Unexpected exception: {ex.Message}");
        }
    }
}
