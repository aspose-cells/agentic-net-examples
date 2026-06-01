using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsOutputBlankPageDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new empty workbook
                Workbook workbook = new Workbook();

                // Hide the only worksheet to simulate a workbook with nothing to print
                Worksheet sheet = workbook.Worksheets[0];
                sheet.IsVisible = false;

                // Configure PDF save options: do NOT output a blank page when nothing to print
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    OutputBlankPageWhenNothingToPrint = false
                };

                // Attempt to save the workbook as PDF.
                // This should throw a CellsException because the workbook has no printable content.
                workbook.Save("EmptyWorkbook.pdf", pdfOptions);
                Console.WriteLine("PDF saved successfully (unexpected).");
            }
            catch (CellsException ex)
            {
                // Expected exception handling when there is nothing to print
                Console.WriteLine("Caught CellsException as expected:");
                Console.WriteLine($"Message: {ex.Message}");
            }
            catch (Exception ex)
            {
                // Any other unexpected exceptions
                Console.WriteLine("An unexpected exception occurred:");
                Console.WriteLine(ex);
            }
        }
    }
}