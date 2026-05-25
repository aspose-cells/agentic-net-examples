using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsDemo
{
    public class PdfSaveWithExceptionHandling
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook (no external template required)
                Workbook workbook = new Workbook();

                // Hide the only worksheet to simulate a scenario where nothing is printable
                workbook.Worksheets[0].IsVisible = false;

                // Configure PDF save options
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    // Do not generate a blank page when there is nothing to print
                    OutputBlankPageWhenNothingToPrint = false,
                    // Hide rendering errors (prevents other exceptions)
                    IgnoreError = true
                };

                // Define output path
                string outputPath = "output.pdf";

                // Save the workbook as PDF
                workbook.Save(outputPath, pdfOptions);
                Console.WriteLine("PDF saved successfully.");
            }
            catch (CellsException ex)
            {
                // Handle Aspose.Cells specific exceptions (e.g., nothing printed)
                Console.WriteLine($"CellsException caught: {ex.Message}");
            }
            catch (Exception ex)
            {
                // Handle any other unexpected exceptions
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            PdfSaveWithExceptionHandling.Run();
        }
    }
}