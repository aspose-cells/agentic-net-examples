using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class Program
{
    static void Main()
    {
        // Create a new workbook and add some sample data
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        worksheet.Cells["A1"].PutValue("Sample data for PDF compression demo");
        worksheet.Cells["A2"].PutValue(DateTime.Now);

        // Create PDF save options with high compression settings
        PdfSaveOptions pdfSaveOptions = new PdfSaveOptions();

        // Set the core compression algorithm to Flate (high compression)
        pdfSaveOptions.PdfCompression = PdfCompressionCore.Flate;

        // Optimize for minimum file size while preserving visual fidelity
        pdfSaveOptions.OptimizationType = PdfOptimizationType.MinimumSize;

        // Save the workbook as a PDF using the configured options
        workbook.Save("CompressedOutput.pdf", pdfSaveOptions);
    }
}