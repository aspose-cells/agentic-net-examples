using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsWatermarkDemo
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sample data for PDF with watermark");

            // Define the font for the text watermark
            RenderingFont font = new RenderingFont("Calibri", 68)
            {
                Bold = true,
                Italic = true,
                Color = Color.Blue
            };

            // Create a text watermark using the font
            RenderingWatermark watermark = new RenderingWatermark("CONFIDENTIAL", font)
            {
                // Center the watermark on the page
                HAlignment = TextAlignmentType.Center,
                VAlignment = TextAlignmentType.Center,
                // Rotate for visual effect
                Rotation = 45,
                // Make it semi‑transparent
                Opacity = 0.3f,
                // Scale relative to page size
                ScaleToPagePercent = 75,
                // Place it behind the page content
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