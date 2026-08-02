using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class Program
{
    static void Main()
    {
        // Load the source workbook (adjust the path as needed)
        Workbook workbook = new Workbook("input.xlsx");

        // Configure PDF save options
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Set PDF/A‑2b compliance
            Compliance = PdfCompliance.PdfA2b

            // Aspose.Cells retains the original color profiles during PDF conversion,
            // so no additional setting is required to preserve accurate colors.
        };

        // Save the workbook as a PDF with the specified compliance level
        workbook.Save("output.pdf", pdfOptions);
    }
}

// Author: Aspose.Cells .NET example – PDF/A‑2b compliance with color profile preservation.