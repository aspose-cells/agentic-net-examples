using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Add some sample data
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Compression Level Demo");
        sheet.Cells["A2"].PutValue("This PDF uses maximum compression.");

        // Create PDF save options with maximum compression settings
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Use Flate compression for the core PDF content (best compression)
        pdfOptions.PdfCompression = PdfCompressionCore.Flate;

        // Optimize the PDF for minimum file size
        pdfOptions.OptimizationType = PdfOptimizationType.MinimumSize;

        // Save the workbook as a PDF using the configured options
        workbook.Save("CompressedOutput.pdf", pdfOptions);
    }
}