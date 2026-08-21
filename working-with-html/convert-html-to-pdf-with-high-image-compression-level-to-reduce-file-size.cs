// Title: C# – Convert HTML to PDF with high image compression using Aspose.Cells
// Description: Load an HTML file into an Aspose.Cells Workbook, configure PdfSaveOptions for Flate compression, select MinimumSize optimization, and resample images to 96 dpi with 50 % JPEG quality before saving a compact PDF.
// Keywords: Aspose.Cells | HTML to PDF C# | PDF compression Flate | MinimumSize PDF | SetImageResample | strong image compression | reduce PDF file size | .NET PDF generation
// Common Searches: Aspose.Cells compress PDF images when converting HTML | C# set PDF optimization to minimum size Aspose | How to resample images for PDF output with Aspose.Cells | Flate compression option in Aspose.Cells PDF save | Batch HTML to compressed PDF using Aspose.Cells .NET
// Developer Intent: Create a PDF from an HTML source while applying aggressive image compression to achieve the smallest possible file size.
// Use Cases: Email‑ready PDFs of web reports that stay under attachment limits. | Archiving large numbers of web pages as space‑efficient PDFs. | Generating PDFs for mobile apps where bandwidth and storage are limited.
// AI Prompts: Show how to adjust DPI and JPEG quality with SetImageResample in Aspose.Cells PDF conversion. | Explain the effect of PdfCompressionCore.Flate and PdfOptimizationType.MinimumSize on PDF size and quality. | Provide a C# loop that converts multiple HTML files to compressed PDFs using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace HtmlToPdfConversion
{
    // Load an HTML file into an Aspose.Cells Workbook, configure PdfSaveOptions for Flate compression, select MinimumSize optimization, and resample images to 96 dpi with 50 % JPEG quality before saving a compact PDF.
    class Program
    {
        static void Main()
        {
            // Load the HTML file into a workbook.
            // The constructor automatically detects the format based on the file extension.
            Workbook workbook = new Workbook("input.html");

            // Create PDF save options.
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Use Flate compression for the PDF core content.
            pdfOptions.PdfCompression = PdfCompressionCore.Flate;

            // Optimize for minimum file size (prioritizes size over print quality).
            pdfOptions.OptimizationType = PdfOptimizationType.MinimumSize;

            // Resample images to a lower DPI and reduce JPEG quality to achieve higher compression.
            // DesiredPPI of 96 is suitable for email/web, JPEG quality of 50% balances quality and size.
            pdfOptions.SetImageResample(96, 50);

            // Save the workbook as a PDF with the specified compression settings.
            workbook.Save("output.pdf", pdfOptions);
        }
    }
}
