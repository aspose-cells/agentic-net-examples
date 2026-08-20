// Title: C# – Convert HTML with Base64 images to high‑resolution PDF using Aspose.Cells
// Description: Load an HTML file that embeds Base64‑encoded images into an Aspose.Cells Workbook, configure PdfSaveOptions.SetImageResample(300, 100) for 300 PPI and full JPEG quality, and save as PDF so the images keep their native resolution.
// Keywords: Aspose.Cells HTML to PDF | Base64 images PDF C# | SetImageResample | high resolution PDF conversion | preserve image quality Aspose.Cells | C# PDF generation from HTML
// Common Searches: Aspose.Cells keep image resolution when converting HTML to PDF | SetImageResample example C# | Convert HTML with embedded Base64 images to PDF | high‑quality PDF from HTML using Aspose.Cells | C# Aspose.Cells PDFSaveOptions image resample
// Developer Intent: Create a PDF from an HTML document that contains Base64‑encoded images without degrading image quality.
// Use Cases: Generate print‑ready PDFs from marketing emails that embed images as Base64. | Produce PDF reports from HTML templates with chart graphics encoded in Base64, ensuring crisp visuals. | Batch‑process HTML invoices containing Base64 logos to PDFs while maintaining logo clarity.
// AI Prompts: Write C# code with Aspose.Cells to load an HTML file containing Base64 images and export it to a PDF at 300 PPI and JPEG quality 100. | Explain how PdfSaveOptions.SetImageResample affects image scaling and how to adjust its parameters for different resolution needs. | Provide a step‑by‑step guide for batch converting multiple HTML files with Base64 images to high‑resolution PDFs using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsHtmlToPdf
{
    // Load an HTML file that embeds Base64‑encoded images into an Aspose.Cells Workbook, configure PdfSaveOptions.SetImageResample(300, 100) for 300 PPI and full JPEG quality, and save as PDF so the images keep their native resolution.
    class Program
    {
        static void Main()
        {
            // Load the HTML file that contains Base64‑encoded images.
            // The Workbook constructor automatically detects the format.
            Workbook workbook = new Workbook("input.html");

            // Create PDF save options.
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Set a high PPI (e.g., 300) and maximum JPEG quality (100) so that
            // images are not down‑sampled and retain their original resolution.
            pdfOptions.SetImageResample(300, 100);

            // Save the workbook as a PDF file.
            workbook.Save("output.pdf", pdfOptions);
        }
    }
}
