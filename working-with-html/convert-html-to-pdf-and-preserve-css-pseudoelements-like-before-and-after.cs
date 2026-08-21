// Title: C# – Convert HTML to PDF with Aspose.Cells while retaining CSS ::before and ::after
// Description: Load an HTML file (including CSS ::before/::after pseudo‑elements) into an Aspose.Cells Workbook and export it to PDF using PdfSaveOptions, preserving the original visual layout.
// Keywords: Aspose.Cells HTML to PDF | C# convert HTML PDF | preserve CSS pseudo‑elements | PdfSaveOptions | Aspose.Cells workbook load HTML
// Common Searches: Aspose.Cells convert HTML to PDF with pseudo elements | C# keep ::before ::after when exporting HTML to PDF | How to render CSS pseudo‑elements in PDF using Aspose.Cells | Save HTML workbook as PDF Aspose.Cells C#
// Developer Intent: Create a PDF from an HTML document that contains CSS ::before and ::after pseudo‑elements using Aspose.Cells for .NET.
// Use Cases: Archive marketing emails that use icon fonts via ::before into printable PDFs. | Generate printable reports from HTML templates that rely on ::after footnote markers. | Batch‑convert a collection of web pages to PDFs while maintaining all CSS visual effects.
// AI Prompts: Show how to set OnePagePerSheet = true in PdfSaveOptions without losing CSS pseudo‑elements. | Demonstrate adding a custom web font so text in ::before and ::after renders correctly in the PDF. | Explain how to reference external CSS files when loading HTML into an Aspose.Cells Workbook for PDF export.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Load an HTML file (including CSS ::before/::after pseudo‑elements) into an Aspose.Cells Workbook and export it to PDF using PdfSaveOptions, preserving the original visual layout.
class HtmlToPdfConverter
{
    static void Main()
    {
        // Input HTML file that may contain CSS pseudo‑elements (::before, ::after)
        string htmlFile = "input.html";

        // Output PDF file
        string pdfFile = "output.pdf";

        // Load the HTML document into a Workbook.
        // Aspose.Cells automatically parses the HTML and creates the corresponding worksheet.
        Workbook workbook = new Workbook(htmlFile);

        // Configure PDF save options.
        // The default rendering preserves the visual appearance of the HTML,
        // including CSS pseudo‑elements, as they are interpreted during the load phase.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Example: keep each worksheet on its own page (optional)
            OnePagePerSheet = false
        };

        // Save the workbook as a PDF file.
        workbook.Save(pdfFile, pdfOptions);

        Console.WriteLine($"HTML file \"{htmlFile}\" has been converted to PDF \"{pdfFile}\".");
    }
}
