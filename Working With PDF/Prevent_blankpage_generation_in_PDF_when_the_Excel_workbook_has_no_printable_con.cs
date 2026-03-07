using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsPdfNoBlankPageDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (empty)
            Workbook workbook = new Workbook();

            // Add some content to avoid "nothing to print" exception
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sample Text");

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Do not generate a blank page when there is nothing to print
                OutputBlankPageWhenNothingToPrint = false,

                // Ignore completely blank pages if any appear
                PrintingPageType = PrintingPageType.IgnoreBlank
            };

            // Define output path (adjust as needed)
            string outputPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "NoBlankPageOutput.pdf");

            // Save the workbook to PDF using the configured options
            workbook.Save(outputPath, pdfOptions);

            Console.WriteLine($"PDF saved to: {outputPath}");
        }
    }
}