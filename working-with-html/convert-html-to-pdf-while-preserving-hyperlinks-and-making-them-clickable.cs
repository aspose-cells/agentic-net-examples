// Title: C# – Convert HTML to PDF with clickable hyperlinks using Aspose.Cells
// Description: Load an HTML file into an Aspose.Cells Workbook and export it as a PDF. All anchor tags are retained, so the resulting PDF contains active, clickable links.
// Keywords: Aspose.Cells HTML to PDF | C# PDF conversion | preserve hyperlinks | clickable PDF links | SaveFormat.Pdf | HTML workbook Aspose.Cells | convert HTML file to PDF C#
// Common Searches: Aspose.Cells keep links when converting HTML to PDF | C# convert HTML file to PDF with active hyperlinks | How to export HTML as PDF with clickable links using Aspose.Cells | HTML to PDF conversion example Aspose.Cells .NET
// Developer Intent: Generate a PDF from an HTML document while preserving functional hyperlinks, using Aspose.Cells for .NET.
// Use Cases: Transform marketing email HTML templates into PDF brochures that retain all call‑to‑action links. | Archive web pages as PDFs for offline reading, keeping navigation links usable. | Automate report pipelines where HTML tables with URLs are rendered as PDFs with active links.
// AI Prompts: Show a C# snippet that loads an HTML file into an Aspose.Cells Workbook and saves it as a PDF with clickable hyperlinks. | Explain how Aspose.Cells maps HTML anchor tags to PDF link annotations during export and whether any extra settings are needed. | Provide code to batch‑process a folder of HTML files, converting each to a PDF while preserving all hyperlinks.

using System;
using Aspose.Cells;

// Load an HTML file into an Aspose.Cells Workbook and export it as a PDF. All anchor tags are retained, so the resulting PDF contains active, clickable links.
class HtmlToPdfConverter
{
    static void Main()
    {
        // Load the source HTML file into a workbook.
        // Aspose.Cells can parse HTML and represent it as an Excel workbook.
        Workbook workbook = new Workbook("input.html");

        // Save the workbook as PDF.
        // Hyperlinks present in the HTML are retained and become clickable in the PDF.
        workbook.Save("output.pdf", SaveFormat.Pdf);
    }
}
