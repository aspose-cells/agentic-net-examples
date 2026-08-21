// Title: Convert HTML to PDF with Aspose.Cells for .NET and add the source file name as a page header
// Description: Loads an HTML file into an Aspose.Cells Workbook, sets the left header to the file name using the &F placeholder, adds page number and date, configures PDF save options, and saves the workbook as a PDF where the header appears on every page.
// Keywords: Aspose.Cells | HTML to PDF conversion | C# .NET | page header | file name placeholder &F | PdfSaveOptions | add page numbers | add current date | batch HTML to PDF | document title in PDF
// Common Searches: Aspose.Cells add file name to PDF header | C# convert HTML to PDF with header | How to use &F placeholder in Aspose.Cells | Set page header when saving PDF from HTML Aspose.Cells | Add page numbers and date to PDF generated from HTML
// Developer Intent: Create a PDF from an HTML document and automatically include the HTML file name (plus optional page number and date) in the header of each PDF page using Aspose.Cells for .NET.
// Use Cases: Generate printable PDFs from HTML templates while showing the template name in the header for traceability. | Batch‑process a folder of HTML files into PDFs, each PDF displaying its source filename on every page. | Produce documentation PDFs that need consistent branding: file name on the left, page X of Y in the center, and the generation date on the right.
// AI Prompts: Write C# code that loads an HTML file into an Aspose.Cells Workbook, sets a left header with the file name (&F), adds page number and date, and saves the result as a PDF. | Provide a reusable method AcceptHtmlAndPdf(string htmlPath, string pdfPath) that configures PageSetup headers (file name, page number, total pages, date) and uses PdfSaveOptions to export the workbook. | Explain how to customize left, center, and right header sections when converting HTML to PDF with Aspose.Cells, including using placeholders like &F, &P, &N, &D and adding static text or images.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Saving;

// Loads an HTML file into an Aspose.Cells Workbook, sets the left header to the file name using the &F placeholder, adds page number and date, configures PDF save options, and saves the workbook as a PDF where the header appears on every page.
class HtmlToPdfWithHeader
{
    static void Main()
    {
        try
        {
            // Path of the source HTML file
            string htmlPath = "source.html";

            // Ensure the HTML file exists; create a simple one if missing
            if (!File.Exists(htmlPath))
            {
                string sampleHtml = "<html><body><h1>Sample HTML Content</h1></body></html>";
                File.WriteAllText(htmlPath, sampleHtml);
                Console.WriteLine($"Created placeholder HTML file at '{htmlPath}'.");
            }

            // Load the HTML file into a workbook
            Workbook workbook = new Workbook(htmlPath);

            // Use the first worksheet (or iterate all worksheets if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Set a page header that displays the source file name on every page
            // &F inserts the file name without the path
            sheet.PageSetup.SetHeader(0, "&F"); // Left section
            // Optional: add page number in the center and date on the right
            sheet.PageSetup.SetHeader(1, "Page &P of &N");
            sheet.PageSetup.SetHeader(2, "&D");

            // Configure PDF save options (optional: display document title in the PDF window)
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                DisplayDocTitle = true
            };

            // Save the workbook as a PDF file; the header will appear on each page
            string pdfPath = "output.pdf";

            // Ensure the output directory exists
            string pdfDir = Path.GetDirectoryName(Path.GetFullPath(pdfPath));
            if (!Directory.Exists(pdfDir))
            {
                Directory.CreateDirectory(pdfDir);
            }

            workbook.Save(pdfPath, pdfOptions);

            Console.WriteLine($"HTML file '{htmlPath}' has been converted to PDF with header on each page: {pdfPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
