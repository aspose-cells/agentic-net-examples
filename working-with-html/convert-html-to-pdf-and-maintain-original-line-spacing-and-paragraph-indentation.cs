// Title: C# – Convert HTML to PDF while keeping exact line spacing and paragraph indentation with Aspose.Cells
// Description: This example shows how to load an HTML file into an Aspose.Cells Workbook using HtmlLoadOptions.DeleteRedundantSpaces = false, then save it as a PDF with PdfSaveOptions. The original whitespace, line breaks, and paragraph indentation are retained in the output PDF, making the conversion faithful to the source HTML layout.
// Keywords: Aspose.Cells HTML to PDF | DeleteRedundantSpaces false | preserve whitespace Aspose.Cells | C# HTML to PDF conversion | maintain paragraph indentation PDF | Aspose.Cells .NET example | HTML layout retention PDF
// Common Searches: Aspose.Cells keep spaces when converting HTML to PDF | HtmlLoadOptions DeleteRedundantSpaces example C# | Convert HTML file to PDF without losing indentation | C# preserve line breaks HTML to PDF Aspose | How to retain original formatting in PDF generated from HTML
// Developer Intent: Create a PDF from an HTML document that mirrors the source's spacing and indentation using Aspose.Cells for .NET.
// Use Cases: Archiving marketing‑email HTML templates as printable PDFs with exact layout. | Generating compliance‑ready PDFs from HTML reports where whitespace matters. | Batch‑processing web‑page snapshots into PDFs that must preserve original formatting.
// AI Prompts: Write C# code that uses Aspose.Cells to convert an HTML file to PDF, ensuring whitespace and paragraph indentation are unchanged. | Explain how HtmlLoadOptions.DeleteRedundantSpaces influences HTML‑to‑PDF conversion and how to tweak PdfSaveOptions for page size or orientation. | Provide a loop that processes multiple HTML files, converting each to a PDF while preserving formatting with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// This example shows how to load an HTML file into an Aspose.Cells Workbook using HtmlLoadOptions.DeleteRedundantSpaces = false, then save it as a PDF with PdfSaveOptions. The original whitespace, line breaks, and paragraph indentation are retained in the output PDF, making the conversion faithful to the source HTML layout.
class HtmlToPdfConverter
{
    static void Main()
    {
        // Path to the source HTML file
        string htmlFile = "input.html";

        // Path for the resulting PDF file
        string pdfFile = "output.pdf";

        // Load the HTML into a workbook.
        // DeleteRedundantSpaces = false preserves original spaces and line breaks.
        HtmlLoadOptions loadOptions = new HtmlLoadOptions
        {
            DeleteRedundantSpaces = false
        };
        Workbook workbook = new Workbook(htmlFile, loadOptions);

        // Configure PDF save options if needed (default settings keep formatting).
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Save the workbook as PDF, preserving line spacing and paragraph indentation.
        workbook.Save(pdfFile, pdfOptions);

        Console.WriteLine("HTML successfully converted to PDF with original formatting.");
    }
}
