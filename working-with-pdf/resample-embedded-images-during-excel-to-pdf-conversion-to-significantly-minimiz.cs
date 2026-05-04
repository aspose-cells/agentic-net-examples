using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class ResampleImagesPdf
{
    static void Main()
    {
        // Load the source Excel workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Create PDF save options
        PdfSaveOptions pdfSaveOptions = new PdfSaveOptions();

        // Resample all embedded images to 96 PPI with 80% JPEG quality
        // This converts images to JPEG and reduces their resolution, shrinking the PDF size
        pdfSaveOptions.SetImageResample(96, 80);

        // Further minimize file size by using the MinimumSize optimization type
        pdfSaveOptions.OptimizationType = PdfOptimizationType.MinimumSize;

        // Save the workbook as a PDF using the configured options
        workbook.Save("output.pdf", pdfSaveOptions);
    }
}