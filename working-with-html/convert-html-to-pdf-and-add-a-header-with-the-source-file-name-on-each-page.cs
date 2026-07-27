// Title: C# – Convert HTML to PDF with filename header using Aspose.Cells
// Description: Loads an HTML file into an Aspose.Cells Workbook, sets a left header to the source filename (&F) and a centered page‑number header, enables document‑title display, and saves the workbook as a PDF.
// Keywords: Aspose.Cells | C# HTML to PDF | filename header PDF | page header &F placeholder | PdfSaveOptions DisplayDocTitle | convert HTML workbook to PDF | set page header Aspose.Cells | C# PDF generation Aspose
// Common Searches: Aspose.Cells add file name to PDF header | C# convert html to pdf with header | How to use &F placeholder in Aspose.Cells | Set page numbers in PDF generated from HTML | Display document title when saving PDF with Aspose.Cells
// Developer Intent: Generate a PDF from an HTML document and automatically include the HTML file name in the page header.
// Use Cases: Produce printable reports from HTML templates where each PDF page shows the source file name for traceability. | Batch‑convert a collection of HTML files to PDFs with consistent filename headers and page numbers for archival purposes. | Create invoice PDFs from HTML where the invoice number is embedded in the filename and displayed on every page.
// AI Prompts: Write C# code with Aspose.Cells that converts an HTML file to PDF and adds a left header showing the source filename on each page. | Show how to configure a centered page‑number header and enable DisplayDocTitle in PdfSaveOptions when saving a workbook as PDF. | Explain how to replace the &F placeholder with a custom text or full file path while converting HTML to PDF using Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace Example
{
    // Loads an HTML file into an Aspose.Cells Workbook, sets a left header to the source filename (&F) and a centered page‑number header, enables document‑title display, and saves the workbook as a PDF.
    class HtmlToPdfWithHeader
    {
        static void Main()
        {
            // Input HTML file path
            string htmlPath = "source.html";

            // Output PDF file path
            string pdfPath = "output.pdf";

            try
            {
                // Verify that the HTML file exists before loading
                if (!File.Exists(htmlPath))
                {
                    Console.WriteLine($"Input file not found: {htmlPath}");
                    return;
                }

                // Load the HTML file into a workbook
                Workbook workbook = new Workbook(htmlPath);

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Set left header to file name without path (&F placeholder)
                sheet.PageSetup.SetHeader(0, "&F");

                // Center header with page numbering
                sheet.PageSetup.SetHeader(1, "Page &P of &N");

                // Configure PDF save options
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    DisplayDocTitle = true
                };

                // Save the workbook as a PDF file
                workbook.Save(pdfPath, pdfOptions);
                Console.WriteLine($"PDF saved successfully to {pdfPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
