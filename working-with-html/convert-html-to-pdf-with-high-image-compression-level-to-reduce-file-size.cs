// Title: C# – Convert HTML to PDF with High Image Compression Using Aspose.Cells
// Description: Loads an HTML file into an Aspose.Cells Workbook, configures PdfSaveOptions for Flate compression, MinimumSize optimization, and image resampling (96 PPI, 70 % JPEG quality), then saves a compact PDF. Ideal for reducing PDF size while preserving readable image quality.
// Keywords: Aspose.Cells | HTML to PDF conversion | PDF compression | image resample | MinimumSize PDF | Flate compression | SetImageResample | PdfSaveOptions | C# example | reduce PDF file size
// Common Searches: Aspose.Cells convert HTML to PDF with compression | how to shrink PDF size in Aspose.Cells C# | set JPEG quality for PDF images Aspose.Cells | minimum size PDF option Aspose.Cells | resample images to 96 PPI Aspose.Cells PDF export
// Developer Intent: The developer needs to turn an HTML document into a PDF while applying aggressive image compression to keep the output file as small as possible.
// Use Cases: Email‑friendly PDF reports generated from HTML templates. | Archiving web pages or dashboards with minimal storage impact. | Creating printable PDFs for mobile devices where bandwidth is limited.
// AI Prompts: Generate C# code that uses Aspose.Cells to convert HTML to a PDF with maximum compression and custom image resampling. | Explain the effect of PdfOptimizationType.MinimumSize and SetImageResample on PDF file size in Aspose.Cells. | Suggest alternative Aspose.Cells settings for even smaller PDFs, such as different JPEG quality levels or image downsampling strategies.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Loads an HTML file into an Aspose.Cells Workbook, configures PdfSaveOptions for Flate compression, MinimumSize optimization, and image resampling (96 PPI, 70 % JPEG quality), then saves a compact PDF. Ideal for reducing PDF size while preserving readable image quality.
class HtmlToPdfConverter
{
    static void Main()
    {
        // Load the HTML file into a workbook
        Workbook workbook = new Workbook("input.html");

        // Create PDF save options
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Use Flate compression for non‑image content
        pdfOptions.PdfCompression = PdfCompressionCore.Flate;

        // Optimize for minimum file size (higher compression)
        pdfOptions.OptimizationType = PdfOptimizationType.MinimumSize;

        // Resample images to 96 PPI and set JPEG quality to 70 %
        pdfOptions.SetImageResample(96, 70);

        // Save the workbook as a PDF with the configured options
        workbook.Save("output.pdf", pdfOptions);
    }
}
