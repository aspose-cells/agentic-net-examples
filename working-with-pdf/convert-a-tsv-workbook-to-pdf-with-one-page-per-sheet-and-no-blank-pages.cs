using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsTsvToPdf
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source TSV file
            string tsvPath = "input.tsv";

            // Load the TSV workbook
            // Aspose.Cells automatically detects the format based on the file extension
            Workbook workbook = new Workbook(tsvPath);

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Ensure each worksheet is rendered on a single page
                OnePagePerSheet = true,

                // Omit pages that contain only blank cells
                PrintingPageType = PrintingPageType.IgnoreBlank,

                // Do not generate a blank page when a sheet has nothing to print
                OutputBlankPageWhenNothingToPrint = false
            };

            // Save the workbook as PDF
            string pdfPath = "output.pdf";
            workbook.Save(pdfPath, pdfOptions);

            Console.WriteLine($"TSV workbook converted to PDF successfully: {pdfPath}");
        }
    }
}