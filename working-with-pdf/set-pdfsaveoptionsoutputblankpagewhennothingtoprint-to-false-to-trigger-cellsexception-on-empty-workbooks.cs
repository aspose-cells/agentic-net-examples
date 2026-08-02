using System;
using Aspose.Cells;

// Author: Aspose.Cells .NET example – demonstrates triggering CellsException by disabling blank page output.

class Program
{
    static void Main()
    {
        // Create a new, empty workbook.
        Workbook workbook = new Workbook();

        // Configure PDF save options to *not* output a blank page when nothing is printable.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            OutputBlankPageWhenNothingToPrint = false
        };

        try
        {
            // Attempt to save the empty workbook as PDF.
            // With the option set to false, Aspose.Cells throws a CellsException.
            workbook.Save("EmptyWorkbook.pdf", pdfOptions);
            Console.WriteLine("Workbook saved without exception (unexpected).");
        }
        catch (CellsException ex)
        {
            // Expected path: capture and display the exception details.
            Console.WriteLine($"Caught expected CellsException: {ex.Message}");
        }
    }
}