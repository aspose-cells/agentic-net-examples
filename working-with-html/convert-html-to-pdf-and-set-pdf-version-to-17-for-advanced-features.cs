// Title: Convert HTML to PDF with Aspose.Cells (.NET) – PDF 1.7 compliance
// Description: Loads an HTML file into an Aspose.Cells Workbook, configures PdfSaveOptions with PdfCompliance.Pdf17, and saves the workbook as a PDF that meets PDF 1.7 specifications.
// Keywords: Aspose.Cells | HTML to PDF | C# PDF 1.7 | PdfSaveOptions | PdfCompliance.Pdf17 | .NET conversion | PDF version control
// Common Searches: Aspose.Cells convert HTML to PDF C# | set PDF version to 1.7 Aspose.Cells | PdfSaveOptions compliance PDF 1.7 example | C# convert HTML file to PDF with Aspose | how to enforce PDF 1.7 when saving workbook
// Developer Intent: Create a PDF from an HTML source using Aspose.Cells and enforce PDF 1.7 compliance.
// Use Cases: Generate printable PDFs from web‑based reports while preserving transparency and layers. | Archive HTML invoices as PDF 1.7 documents for legal compliance. | Batch‑convert marketing HTML pages to PDF with a consistent version for downstream processing.
// AI Prompts: Provide code to embed a custom TrueType font while keeping PdfCompliance.Pdf17. | Show how to switch the compliance level to PDF/A‑2b in the same conversion flow. | Explain how to handle missing or malformed HTML input during the conversion. | Demonstrate converting HTML with external CSS and images to PDF using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsHtmlToPdf
{
    // Loads an HTML file into an Aspose.Cells Workbook, configures PdfSaveOptions with PdfCompliance.Pdf17, and saves the workbook as a PDF that meets PDF 1.7 specifications.
    class Program
    {
        static void Main()
        {
            // Load the HTML file into a workbook
            // (Assumes "input.html" exists in the application directory)
            Workbook workbook = new Workbook("input.html");

            // Create PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Set PDF compliance level to PDF 1.7 (enables advanced features)
            pdfOptions.Compliance = PdfCompliance.Pdf17;

            // Save the workbook as a PDF file with the specified compliance level
            workbook.Save("output.pdf", pdfOptions);

            Console.WriteLine("HTML successfully converted to PDF with PDF 1.7 compliance.");
        }
    }
}
