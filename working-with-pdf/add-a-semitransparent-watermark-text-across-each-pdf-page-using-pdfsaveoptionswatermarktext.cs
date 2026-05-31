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
            // Create a new workbook and add some sample data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sample data for PDF with watermark");
            sheet.Cells["B2"].PutValue("Another cell");

            // Define the font for the watermark text
            RenderingFont watermarkFont = new RenderingFont("Arial", 72)
            {
                Bold = true,
                Color = Color.Gray   // Light gray for a subtle appearance
            };

            // Create a text watermark with the desired properties
            RenderingWatermark watermark = new RenderingWatermark("CONFIDENTIAL", watermarkFont)
            {
                HAlignment = TextAlignmentType.Center,   // Center horizontally
                VAlignment = TextAlignmentType.Center,   // Center vertically
                Rotation = 45f,                          // Diagonal orientation
                Opacity = 0.3f,                          // Semi‑transparent
                ScaleToPagePercent = 75,                 // Scale relative to page size
                IsBackground = true                      // Place behind page contents
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