using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class Program
{
    static void Main()
    {
        // Load the source workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Create PDF save options
        PdfSaveOptions pdfSaveOptions = new PdfSaveOptions();

        // Set image resampling: 150 PPI for reasonable resolution and 80% JPEG quality
        pdfSaveOptions.SetImageResample(150, 80);

        // Save the workbook as PDF using the configured options
        workbook.Save("output.pdf", pdfSaveOptions);
    }
}