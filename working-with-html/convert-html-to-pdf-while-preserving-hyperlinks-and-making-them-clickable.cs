// Title: C# – Convert HTML to PDF with Clickable Links using Aspose.Cells
// Description: Demonstrates loading an HTML file into an Aspose.Cells Workbook and saving it as a PDF. Hyperlinks from the source HTML are automatically retained as clickable links in the generated PDF, with basic PdfSaveOptions configuration.
// Keywords: Aspose.Cells | C# | HTML to PDF | clickable hyperlinks | PdfSaveOptions | preserve links | convert HTML workbook to PDF | Aspose.Cells .NET | HTML report to PDF
// Common Searches: Aspose.Cells convert HTML to PDF C# | preserve hyperlinks when saving PDF Aspose.Cells | C# HTML to PDF clickable links | PdfSaveOptions preserve links | how to keep HTML links in PDF using Aspose
// Developer Intent: Generate a PDF from an HTML document while keeping every original hyperlink active, using Aspose.Cells for .NET.
// Use Cases: Transform web‑based reports into printable PDFs that retain functional URLs for distribution. | Automate batch conversion of multiple HTML files to PDFs with active links for archival or compliance purposes. | Create PDF versions of HTML email templates where recipients can still click embedded links.
// AI Prompts: Provide C# code that loads an HTML file into an Aspose.Cells Workbook and saves it as a PDF with clickable hyperlinks. | Explain how to customize PdfSaveOptions (e.g., page size, image quality) while ensuring hyperlinks remain functional in the PDF output. | Show a C# script that batch processes a folder of HTML files into PDFs, preserving active links using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Saving;

// Demonstrates loading an HTML file into an Aspose.Cells Workbook and saving it as a PDF. Hyperlinks from the source HTML are automatically retained as clickable links in the generated PDF, with basic PdfSaveOptions configuration.
class HtmlToPdfConverter
{
    static void Main()
    {
        // Load the HTML file into a workbook.
        // The constructor automatically detects the format based on the file extension.
        Workbook workbook = new Workbook("input.html");

        // Create PDF save options. Hyperlinks are preserved by default,
        // but you can customize other PDF settings here if needed.
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Save the workbook as a PDF file. The resulting PDF will contain
        // clickable hyperlinks that were present in the original HTML.
        workbook.Save("output.pdf", pdfOptions);
    }
}
