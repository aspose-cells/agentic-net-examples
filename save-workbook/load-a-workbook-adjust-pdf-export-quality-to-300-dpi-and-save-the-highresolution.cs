using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class Program
{
    static void Main()
    {
        // Set the global DPI to 300 for high‑resolution rendering
        CellsHelper.DPI = 300;

        // Load the source workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Create PDF save options and configure image resampling to 300 PPI with maximum JPEG quality
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        pdfOptions.SetImageResample(300, 100); // 300 PPI, 100 % JPEG quality

        // Save the workbook as a high‑resolution PDF
        workbook.Save("output_high_res.pdf", pdfOptions);
    }
}