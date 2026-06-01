using System;
using Aspose.Cells;
using Aspose.Cells.Saving;

namespace AsposeCellsHtmlToPdf
{
    class Program
    {
        static void Main()
        {
            // Load the HTML file into a workbook
            Workbook workbook = new Workbook("input.html");

            // Set a footer that displays the page number in Roman numerals
            // Section 1 = center section; &R prints the page number in Roman numerals
            workbook.Worksheets[0].PageSetup.SetFooter(1, "Page &R");

            // Configure PDF save options (optional: set a default font)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();
            pdfOptions.DefaultFont = "Arial";

            // Save the workbook as a PDF file; the footer will appear on each page
            workbook.Save("output.pdf", pdfOptions);
        }
    }
}