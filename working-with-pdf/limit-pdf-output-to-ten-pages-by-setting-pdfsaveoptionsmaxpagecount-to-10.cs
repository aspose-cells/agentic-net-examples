using System;
using Aspose.Cells;

class LimitPdfPages
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook(); // new Workbook("input.xlsx");

        // Initialize PDF save options
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Limit the PDF output to a maximum of 10 pages
        pdfOptions.PageCount = 10;

        // Save the workbook as a PDF using the configured options
        workbook.Save("output_limited.pdf", pdfOptions);
    }
}