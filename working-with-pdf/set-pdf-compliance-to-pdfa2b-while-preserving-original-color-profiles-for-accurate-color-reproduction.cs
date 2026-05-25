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
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Set PDF compliance to PDF/A‑2b
        pdfOptions.Compliance = PdfCompliance.PdfA2b;

        // Aspose.Cells preserves the original color profiles (ICC) by default,
        // so no additional settings are required for accurate color reproduction.

        // Save the workbook as a PDF with the specified compliance level
        workbook.Save("output.pdf", pdfOptions);
    }
}