using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load the workbook (replace with your source file)
        Workbook workbook = new Workbook("input.xlsx");

        // Create PDF save options and set image resampling to 150 DPI (screen quality)
        // with a JPEG quality of 80% to balance size and clarity.
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        pdfOptions.SetImageResample(150, 80); // desiredPPI = 150, jpegQuality = 80

        // Save the workbook as PDF using the configured options.
        workbook.Save("output.pdf", pdfOptions);
    }
}

// Author: Aspose.Cells .NET example – image resample configuration.