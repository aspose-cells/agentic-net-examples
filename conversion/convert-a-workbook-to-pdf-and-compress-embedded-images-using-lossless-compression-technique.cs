using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class ConvertToPdfWithLosslessImageCompression
{
    static void Main()
    {
        // Load the source workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Enable automatic lossless compression of embedded pictures in the workbook
        workbook.Settings.AutoCompressPictures = true;

        // Create PDF save options
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Apply Flate compression (lossless) to PDF content (excluding images)
        pdfOptions.PdfCompression = PdfCompressionCore.Flate;

        // Optimize for minimum file size while keeping quality (optional)
        pdfOptions.OptimizationType = PdfOptimizationType.MinimumSize;

        // Save the workbook as a PDF using the configured options
        workbook.Save("output.pdf", pdfOptions);
    }
}