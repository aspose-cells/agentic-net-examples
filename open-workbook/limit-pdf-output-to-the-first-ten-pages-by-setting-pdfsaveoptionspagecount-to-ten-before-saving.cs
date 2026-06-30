using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load the Excel workbook (replace with your source file)
        Workbook workbook = new Workbook("input.xlsx");

        // Configure PDF save options to limit output to the first 10 pages
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            PageCount = 10 // Only the first ten pages will be saved
        };

        // Save the workbook as a PDF using the configured options
        workbook.Save("output.pdf", pdfOptions);

        // Author note: This example demonstrates limiting PDF output to ten pages via PdfSaveOptions.PageCount.
    }
}