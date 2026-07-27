using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsPdfExport
{
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Create PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Set image resampling to achieve high resolution (e.g., 300 DPI, 90% JPEG quality)
            pdfOptions.SetImageResample(300, 90);

            // Optionally, you can enable other high‑quality settings, such as optimization type
            // pdfOptions.OptimizationType = PdfOptimizationType.Standard;

            // Save the workbook as a PDF using the options
            workbook.Save("output.pdf", pdfOptions);

            Console.WriteLine("Workbook has been exported to PDF with high‑resolution settings.");
        }
    }
}