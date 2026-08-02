// Title: C# – Convert HTML to a compressed PDF with Aspose.Cells (Flate stream compression)
// Description: Loads an HTML file into an Aspose.Cells Workbook, applies Flate object‑stream compression and MinimumSize optimization via PdfSaveOptions, and saves the result as a reduced‑size PDF.
// Keywords: Aspose.Cells | C# | .NET | HTML to PDF conversion | PDF compression | Flate compression | object stream compression | PdfSaveOptions | MinimumSize optimization | reduce PDF size
// Common Searches: Aspose.Cells convert HTML to PDF C# | How to compress PDF with Flate using Aspose.Cells | PDF object stream compression Aspose .NET | Reduce PDF file size Aspose.Cells | Save HTML as PDF with MinimumSize optimization
// Developer Intent: Load an HTML document into a Workbook and export it as a PDF that uses Flate object‑stream compression to minimize file size.
// Use Cases: Create email‑ready PDF reports from HTML templates while keeping attachments lightweight. | Archive web pages as compact PDFs to conserve storage in document repositories. | Batch‑process multiple HTML files into size‑optimized PDFs for a document‑management workflow.
// AI Prompts: Generate C# code using Aspose.Cells to convert an HTML file to PDF with Flate compression and MinimumSize optimization. | Explain how PdfCompressionCore.Flate and PdfOptimizationType.MinimumSize affect PDF size in Aspose.Cells. | Show best‑practice error handling when loading HTML into a Workbook and saving it as a compressed PDF.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Loads an HTML file into an Aspose.Cells Workbook, applies Flate object‑stream compression and MinimumSize optimization via PdfSaveOptions, and saves the result as a reduced‑size PDF.
class HtmlToPdfCompressed
{
    static void Main()
    {
        // Load the HTML file into a workbook
        Workbook workbook = new Workbook("input.html");

        // Configure PDF save options with object stream compression
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        pdfOptions.PdfCompression = PdfCompressionCore.Flate; // Apply Flate compression to PDF streams
        pdfOptions.OptimizationType = PdfOptimizationType.MinimumSize; // Optional: minimize file size

        // Save the workbook as a compressed PDF
        workbook.Save("output.pdf", pdfOptions);
    }
}
