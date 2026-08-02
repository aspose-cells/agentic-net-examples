// Title: Convert HTML to PDF at 300 DPI with Aspose.Cells for .NET
// Description: Loads an HTML file into an Aspose.Cells Workbook, sets the rendering DPI to 300, configures PdfSaveOptions to resample images at 300 PPI with full JPEG quality, and saves the output as a high‑quality PDF.
// Keywords: Aspose.Cells HTML to PDF | 300 DPI PDF | PdfSaveOptions SetImageResample | CellsHelper DPI | high quality PDF .NET | C# Aspose.Cells rendering | image resample 300 PPI
// Common Searches: Aspose.Cells convert HTML to PDF 300 DPI | Set DPI for PDF output using Aspose.Cells .NET | PdfSaveOptions image resample example | How to keep images sharp when saving HTML as PDF | C# generate print‑ready PDF from HTML with Aspose.Cells
// Developer Intent: Generate a PDF from an HTML workbook with 300 DPI graphics using Aspose.Cells.
// Use Cases: Produce print‑ready PDFs from web‑based reports while preserving image clarity. | Batch‑process multiple HTML files into PDFs with a consistent 300 DPI resolution for archival. | Expose an API that accepts HTML content and returns a 300 DPI PDF for downstream workflows.
// AI Prompts: Show how to change the page size while retaining the 300 DPI image quality. | Provide code that streams HTML from a URL instead of a local file and saves it as a 300 DPI PDF. | Explain how to lower the JPEG quality in SetImageResample to reduce file size without noticeable loss.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsHtmlToPdf
{
    // Loads an HTML file into an Aspose.Cells Workbook, sets the rendering DPI to 300, configures PdfSaveOptions to resample images at 300 PPI with full JPEG quality, and saves the output as a high‑quality PDF.
    class Program
    {
        static void Main()
        {
            // Path to the source HTML file
            string htmlPath = "input.html";

            // Load the HTML file into a workbook
            // Aspose.Cells automatically detects the format based on the file extension
            Workbook workbook = new Workbook(htmlPath);

            // Set the DPI to 300 for high‑quality graphics rendering
            CellsHelper.DPI = 300;

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Resample images to 300 PPI and use maximum JPEG quality (100%)
            // This ensures that images in the resulting PDF retain high resolution
            pdfOptions.SetImageResample(300, 100);

            // Save the workbook as a PDF file
            string pdfPath = "output.pdf";
            workbook.Save(pdfPath, pdfOptions);

            Console.WriteLine($"HTML file '{htmlPath}' has been converted to PDF with 300 DPI at '{pdfPath}'.");
        }
    }
}
