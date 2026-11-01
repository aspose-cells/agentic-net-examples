// Title: How to trigger a CellsException by disabling blank page output when saving an empty workbook to PDF with Aspose.Cells (C#)
// AI Prompts: Generate C# code that creates a new Workbook, clears all worksheets, sets PdfSaveOptions.OutputBlankPageWhenNothingToPrint = false, and attempts to save the workbook as a PDF while catching the expected CellsException. | Explain how to catch and process the CellsException that occurs when Aspose.Cells tries to export an empty workbook to PDF with the blank‑page output option turned off, including sample try‑catch logic.
// Common Searches: Aspose.Cells C# save empty workbook to PDF without generating a blank page | PdfSaveOptions OutputBlankPageWhenNothingToPrint false throws CellsException | How to catch CellsException when exporting an empty Excel file to PDF using Aspose.Cells | Prevent blank page creation for empty workbook PDF export Aspose.Cells
// Tags: Aspose.Cells PDF blank page setting | empty workbook PDF export exception Aspose.Cells | C# handle CellsException during PDF save | disable blank page generation Aspose.Cells | Aspose.Cells export empty workbook as PDF

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Demonstrates creating an empty Workbook, removing all worksheets, configuring PdfSaveOptions to suppress blank page output, attempting to save the workbook as a PDF, and catching the expected CellsException.
class Program
{
    static void Main()
    {
        // Create a new, empty workbook.
        Workbook workbook = new Workbook();

        // Remove all worksheets to make the workbook truly empty.
        workbook.Worksheets.Clear();

        // Configure PDF save options to NOT output a blank page when there is nothing to print.
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        pdfOptions.OutputBlankPageWhenNothingToPrint = false; // This will cause a CellsException on empty workbooks.

        try
        {
            // Attempt to save the empty workbook as PDF.
            // Expect a CellsException because there is nothing to render.
            workbook.Save("EmptyWorkbook.pdf", pdfOptions);
            Console.WriteLine("PDF saved successfully (unexpected).");
        }
        catch (CellsException ex)
        {
            // Expected exception for an empty workbook when OutputBlankPageWhenNothingToPrint is false.
            Console.WriteLine($"CellsException caught as expected: {ex.Message}");
        }
        catch (Exception ex)
        {
            // Any other unexpected exceptions.
            Console.WriteLine($"Unexpected exception: {ex.Message}");
        }
    }
}
