using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsWatermarkExample
{
    class Program
    {
        static void Main()
        {
            // Load an existing workbook from file
            // Replace "input.xlsx" with the actual path to your Excel file
            Workbook workbook = new Workbook("input.xlsx");

            // Create a font for the watermark text
            RenderingFont watermarkFont = new RenderingFont("Calibri", 68)
            {
                Italic = true,
                Bold = true,
                Color = Color.Blue
            };

            // Create a text watermark using the font
            RenderingWatermark watermark = new RenderingWatermark("CONFIDENTIAL", watermarkFont)
            {
                // Center the watermark on each page
                HAlignment = TextAlignmentType.Center,
                VAlignment = TextAlignmentType.Center,
                // Rotate for visual effect
                Rotation = 45,
                // Set opacity (0 = fully transparent, 1 = fully opaque)
                Opacity = 0.3f,
                // Scale watermark relative to page size (percentage)
                ScaleToPagePercent = 75,
                // Place the watermark behind the page content
                IsBackground = true
            };

            // Configure PDF save options and assign the watermark
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                Watermark = watermark
            };

            // Save the workbook as a PDF with the watermark applied to every page
            // Replace "output_watermark.pdf" with your desired output path
            workbook.Save("output_watermark.pdf", pdfOptions);
        }
    }
}