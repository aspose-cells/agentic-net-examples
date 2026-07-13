using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class Program
{
    static void Main()
    {
        // Create a new workbook and add some sample content
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        worksheet.Cells["A1"].PutValue("Sample data for PDF compression");

        // Configure PDF save options
        PdfSaveOptions pdfSaveOptions = new PdfSaveOptions();

        // Apply core compression (Flate) to reduce file size
        pdfSaveOptions.PdfCompression = PdfCompressionCore.Flate;

        // Resample images to 150 PPI with 90% JPEG quality to keep image quality acceptable
        pdfSaveOptions.SetImageResample(150, 90);

        // Optimize for minimum file size while preserving reasonable quality
        pdfSaveOptions.OptimizationType = PdfOptimizationType.MinimumSize;

        // Save the workbook as a PDF with the specified compression settings
        workbook.Save("CompressedOutput.pdf", pdfSaveOptions);
    }
}
// Author: Aspose.Cells .NET example – custom PDF compression with image quality control.