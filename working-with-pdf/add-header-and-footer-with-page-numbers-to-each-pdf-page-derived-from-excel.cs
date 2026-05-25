using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsHeaderFooterPdf
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Get the PageSetup object to configure headers and footers
            PageSetup pageSetup = worksheet.PageSetup;

            // Set header sections:
            // Left: file name without path
            // Center: page number of current page and total pages
            // Right: current date
            pageSetup.SetHeader(0, "&F");                 // Left section
            pageSetup.SetHeader(1, "Page &P of &N");      // Center section
            pageSetup.SetHeader(2, "&D");                 // Right section

            // Set footer sections (optional, similar to header)
            pageSetup.SetFooter(0, "&A");                 // Left: sheet name
            pageSetup.SetFooter(1, "Confidential");       // Center: custom text
            pageSetup.SetFooter(2, "&T");                 // Right: current time

            // Optionally, ensure that the same header/footer appears on all pages
            pageSetup.IsHFDiffFirst = false;
            pageSetup.IsHFDiffOddEven = false;

            // Save the workbook as PDF
            PdfSaveOptions pdfOptions = new PdfSaveOptions();
            workbook.Save("WorkbookWithHeaderFooter.pdf", pdfOptions);

            Console.WriteLine("PDF saved with headers and footers.");
        }
    }
}