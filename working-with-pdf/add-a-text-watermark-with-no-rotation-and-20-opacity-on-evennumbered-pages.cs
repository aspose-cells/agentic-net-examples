using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class Program
{
    static void Main()
    {
        // Create a new workbook with three worksheets to generate three pages in PDF
        Workbook workbook = new Workbook();
        workbook.Worksheets[0].Cells["A1"].PutValue("Page 1");

        int sheet2 = workbook.Worksheets.Add();
        workbook.Worksheets[sheet2].Cells["A1"].PutValue("Page 2");

        int sheet3 = workbook.Worksheets.Add();
        workbook.Worksheets[sheet3].Cells["A1"].PutValue("Page 3");

        // Create rendering font for watermark text
        RenderingFont font = new RenderingFont("Arial", 36);

        // Create text watermark with no rotation and 20% opacity
        RenderingWatermark watermark = new RenderingWatermark("CONFIDENTIAL", font);
        watermark.Rotation = 0f;      // No rotation
        watermark.Opacity = 0.2f;     // 20% opacity

        // Aspose.Cells does not currently expose a property to limit a watermark to even‑numbered pages.
        // The watermark will be applied to all pages. If future versions add page‑range support,
        // the appropriate property can be set here.

        // Assign watermark to PDF save options
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        pdfOptions.Watermark = watermark;

        // Save workbook as PDF
        workbook.Save("EvenPagesWatermark.pdf", pdfOptions);
    }
}

// Author: Aspose.Cells .NET example