using System;
using Aspose.Cells;
using Aspose.Cells.Rendering; // for PdfSaveOptions

// Author: Aspose.Cells .NET example – increase PDF image resampling to 300 DPI
class Program
{
    static void Main()
    {
        // Load an existing workbook (replace with your source file)
        Workbook workbook = new Workbook("input.xlsx");

        // Configure PDF save options to resample images at 300 PPI with high JPEG quality
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        pdfOptions.SetImageResample(300, 100); // 300 dpi, 100 % JPEG quality

        // Save the workbook as a PDF using the configured options
        workbook.Save("output.pdf", pdfOptions);
    }
}