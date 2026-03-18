using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsWatermarkDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Add some sample data to the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("This workbook will be saved as PDF with a text watermark.");

            // Create a rendering font for the watermark text
            RenderingFont watermarkFont = new RenderingFont("Calibri", 68)
            {
                Italic = true,
                Bold = true,
                Color = Color.Blue
            };

            // Create a text watermark using the font
            RenderingWatermark watermark = new RenderingWatermark("CONFIDENTIAL", watermarkFont)
            {
                // Center the watermark on the page
                HAlignment = TextAlignmentType.Center,
                VAlignment = TextAlignmentType.Center,
                // Rotate for visual effect
                Rotation = 45,
                // Make it semi‑transparent
                Opacity = 0.3f,
                // Scale relative to the page size
                ScaleToPagePercent = 75,
                // Place it behind the page contents
                IsBackground = true
            };

            // Configure PDF save options and assign the watermark
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                Watermark = watermark
            };

            // Save the workbook as a PDF with the watermark applied
            workbook.Save("output_watermark.pdf", pdfOptions);
        }
    }
}