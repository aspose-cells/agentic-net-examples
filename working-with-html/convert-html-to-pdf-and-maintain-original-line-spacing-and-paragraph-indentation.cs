// Title: Convert HTML to PDF in C# with Aspose.Cells while preserving line spacing and paragraph indentation
// AI Prompts: Generate C# code that loads an HTML file into an Aspose.Cells Workbook using HtmlLoadOptions and saves it as a PDF, keeping the original line spacing and paragraph indentation. | Show how to configure PdfSaveOptions in Aspose.Cells to disable OnePagePerSheet so the PDF pagination follows the HTML layout. | Provide a concise example that converts input.html to output.pdf with formatting retained, using Aspose.Cells for .NET.
// Common Searches: Aspose.Cells preserve line spacing during HTML to PDF conversion C# | C# convert HTML file to PDF keep paragraph indentation Aspose.Cells | How to use HtmlLoadOptions to retain formatting when saving PDF with Aspose.Cells | Disable OnePagePerSheet for HTML to PDF conversion in Aspose.Cells .NET | Aspose.Cells HTML to PDF layout fidelity line spacing indentation
// Tags: HTML-to-PDF conversion Aspose.Cells | HtmlLoadOptions preserve formatting | PdfSaveOptions OnePagePerSheet false | retain line spacing Aspose.Cells | keep paragraph indentation C#

using System;
using Aspose.Cells;

// // Loads an HTML file into an Aspose.Cells Workbook with HtmlLoadOptions to retain original formatting, then saves it as a PDF using PdfSaveOptions with OnePagePerSheet set to false, preserving line spacing and paragraph indentation.
class HtmlToPdfConverter
{
    static void Main()
    {
        // Path to the source HTML file
        string htmlPath = "input.html";

        // Path for the resulting PDF file
        string pdfPath = "output.pdf";

        // Load the HTML document into an Aspose.Cells Workbook.
        // HtmlLoadOptions preserves the original formatting of the HTML,
        // including line spacing and paragraph indentation.
        HtmlLoadOptions loadOptions = new HtmlLoadOptions();
        Workbook workbook = new Workbook(htmlPath, loadOptions);

        // Configure PDF save options to keep the layout as close as possible
        // to the original HTML rendering.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Do not force each sheet onto a single page; allow natural pagination.
            OnePagePerSheet = false,

            // Preserve the original text layout.
            // (Aspose.Cells handles line spacing and indentation automatically
            // when converting from HTML, so no additional settings are required.)
        };

        // Save the workbook as a PDF file.
        workbook.Save(pdfPath, pdfOptions);
    }
}
