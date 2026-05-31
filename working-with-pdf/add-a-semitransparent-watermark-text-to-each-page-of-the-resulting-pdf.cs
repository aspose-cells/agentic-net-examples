using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class Program
{
    static void Main()
    {
        // Create a new workbook and add sample data to three worksheets
        Workbook wb = new Workbook();
        wb.Worksheets[0].Cells["A1"].PutValue("Page 1");
        int idx = wb.Worksheets.Add();
        wb.Worksheets[idx].Cells["A1"].PutValue("Page 2");
        idx = wb.Worksheets.Add();
        wb.Worksheets[idx].Cells["A1"].PutValue("Page 3");

        // Define the font for the watermark text
        RenderingFont font = new RenderingFont("Arial", 72)
        {
            Bold = true,
            Color = Color.Gray   // Light color for a subtle effect
        };

        // Create a semi‑transparent text watermark
        RenderingWatermark watermark = new RenderingWatermark("CONFIDENTIAL", font)
        {
            HAlignment = TextAlignmentType.Center,
            VAlignment = TextAlignmentType.Center,
            Rotation = 45,            // Rotate 45 degrees
            Opacity = 0.25f,          // 25% opacity (semi‑transparent)
            IsBackground = true,      // Place behind page contents
            ScaleToPagePercent = 80   // Scale relative to page size
        };

        // Configure PDF save options to use the watermark
        PdfSaveOptions options = new PdfSaveOptions
        {
            Watermark = watermark
        };

        // Save the workbook as a PDF with the watermark applied to each page
        wb.Save("output_watermark.pdf", options);
    }
}