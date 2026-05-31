using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load an existing Excel workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Create PDF save options
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Limit the conversion to a maximum of 5 pages
        pdfOptions.PageCount = 5; // Equivalent to a MaxPages setting

        // Save the workbook to PDF using the configured options
        workbook.Save("output.pdf", pdfOptions);
    }
}