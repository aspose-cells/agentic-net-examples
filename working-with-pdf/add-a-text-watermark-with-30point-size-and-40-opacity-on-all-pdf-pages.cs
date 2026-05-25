using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class AddWatermark
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook();

        // Add sample data to demonstrate the watermark
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Sample content for PDF with watermark");

        // Create a rendering font with 30‑point size
        RenderingFont font = new RenderingFont("Arial", 30)
        {
            Bold = true,
            Color = Color.Gray
        };

        // Create a text watermark using the font
        RenderingWatermark watermark = new RenderingWatermark("CONFIDENTIAL", font)
        {
            Opacity = 0.4f,                     // 40% opacity
            IsBackground = true,                // place behind page contents
            HAlignment = TextAlignmentType.Center,
            VAlignment = TextAlignmentType.Center,
            Rotation = 0                         // no rotation
        };

        // Configure PDF save options with the watermark
        PdfSaveOptions options = new PdfSaveOptions
        {
            Watermark = watermark
        };

        // Save the workbook as a PDF with the watermark applied to all pages
        workbook.Save("output_watermarked.pdf", options);
    }
}