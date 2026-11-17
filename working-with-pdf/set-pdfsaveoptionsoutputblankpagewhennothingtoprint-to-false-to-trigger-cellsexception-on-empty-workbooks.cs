// Title: How to disable blank page generation and force a CellsException when exporting an empty workbook to PDF using Aspose.Cells in C#
// Common Searches: Aspose.Cells C# export empty workbook to PDF without adding a blank page and get an exception | set PdfSaveOptions.OutputBlankPageWhenNothingToPrint false to trigger CellsException on PDF save | prevent blank page creation when saving an empty worksheet as PDF with Aspose.Cells | catch CellsException while saving an empty workbook to PDF using Aspose.Cells C# API | configure PDF save options for empty content handling in Aspose.Cells C#
// Tags: PdfSaveOptions.OutputBlankPageWhenNothingToPrint configuration | Aspose.Cells empty workbook PDF export | CellsException handling on PDF save | C# disable blank page generation Aspose.Cells | Aspose.Cells PDF save options for empty content

using System;
using Aspose.Cells;

// Demonstrates creating an empty Workbook, clearing its default sheet, setting PdfSaveOptions.OutputBlankPageWhenNothingToPrint to false, and attempting to save to PDF, which throws a CellsException that is caught and logged.
class Program
{
    static void Main()
    {
        // Create a new workbook (empty)
        Workbook workbook = new Workbook();

        // Clear the default worksheet so there is nothing to print
        workbook.Worksheets[0].Cells.Clear();

        // Set PDF save options: do NOT output a blank page when nothing to print
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        pdfOptions.OutputBlankPageWhenNothingToPrint = false;

        try
        {
            // Attempt to save the empty workbook to PDF.
            // This should throw a CellsException because there is nothing to render.
            workbook.Save("empty.pdf", pdfOptions);
            Console.WriteLine("PDF saved successfully (unexpected).");
        }
        catch (CellsException ex)
        {
            // Expected exception handling
            Console.WriteLine($"Expected exception caught: {ex.Message}");
        }
    }
}
