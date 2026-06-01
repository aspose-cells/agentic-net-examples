using System;
using Aspose.Cells;

namespace AsposeCellsPdfResampleDemo
{
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Create PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Set image resampling to 150 DPI with a reasonable JPEG quality (e.g., 80%)
            // This reduces PDF size while keeping image clarity suitable for screen viewing
            pdfOptions.SetImageResample(150, 80);

            // Save the workbook as a PDF using the configured options
            workbook.Save("output_resampled.pdf", pdfOptions);
        }
    }
}