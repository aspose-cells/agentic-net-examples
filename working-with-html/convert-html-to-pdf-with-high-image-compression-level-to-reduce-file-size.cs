using System;
using Aspose.Cells;
using Aspose.Cells.Rendering; // for PdfOptimizationType, PdfCompressionCore

class HtmlToPdfConverter
{
    static void Main()
    {
        // Load the source HTML file into a workbook
        Workbook workbook = new Workbook("input.html");

        // Create PDF save options
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Optimize for minimum file size (prioritizes size over print quality)
        pdfOptions.OptimizationType = PdfOptimizationType.MinimumSize;

        // Use Flate compression for PDF content (good compression ratio)
        pdfOptions.PdfCompression = PdfCompressionCore.Flate;

        // Resample images to lower resolution and JPEG quality to further reduce size
        // Desired PPI: 96 (screen quality), JPEG quality: 50%
        pdfOptions.SetImageResample(96, 50);

        // Save the workbook as a PDF with the specified compression settings
        workbook.Save("output.pdf", pdfOptions);
    }
}