using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class HtmlToPdfConverter
{
    static void Main()
    {
        // Load the HTML file into a workbook.
        // The Workbook constructor automatically detects the format based on the file extension.
        Workbook workbook = new Workbook("input.html");

        // Create PDF save options and set the compliance level to PDF 1.7.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            Compliance = PdfCompliance.Pdf17
        };

        // Save the workbook as a PDF file using the specified options.
        workbook.Save("output.pdf", pdfOptions);
    }
}