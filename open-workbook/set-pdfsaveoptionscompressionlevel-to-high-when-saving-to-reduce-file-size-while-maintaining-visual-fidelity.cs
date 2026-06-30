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
        worksheet.Cells["A1"].PutValue("Demo for high PDF compression");
        worksheet.Cells["A2"].PutValue(DateTime.Now);

        // Configure PDF save options for high compression
        PdfSaveOptions saveOptions = new PdfSaveOptions();

        // Use Flate compression for all non‑image content (high compression)
        saveOptions.PdfCompression = PdfCompressionCore.Flate;

        // Optimize the PDF for minimum file size (includes image resampling if needed)
        saveOptions.OptimizationType = PdfOptimizationType.MinimumSize;

        // Save the workbook as a PDF with the specified compression settings
        workbook.Save("HighCompressionOutput.pdf", saveOptions);
    }
}
// Author: Aspose.Cells .NET example – high PDF compression settings.