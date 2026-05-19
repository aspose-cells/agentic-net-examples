using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class Program
{
    static void Main()
    {
        // Load the existing workbook from file
        Workbook workbook = new Workbook("input.xlsx");

        // Create PDF save options instance
        PdfSaveOptions pdfSaveOptions = new PdfSaveOptions();

        // Set the desired image resolution to 300 DPI and maximum JPEG quality (100%)
        // This ensures that images inside the PDF are rendered at high resolution.
        pdfSaveOptions.SetImageResample(300, 100);

        // Save the workbook as a high‑resolution PDF using the configured options
        workbook.Save("output_300dpi.pdf", pdfSaveOptions);
    }
}