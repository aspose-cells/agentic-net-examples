using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace WatermarkPdfDemo
{
    // Author: Aspose.Cells .NET example
    class Program
    {
        static void Main()
        {
            // Create a new workbook (in-memory)
            Workbook workbook = new Workbook();

            // Optionally add some content to demonstrate the watermark effect
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sample Data");
            sheet.Cells["A2"].PutValue(DateTime.Now);

            // Create a rendering font for the watermark text
            RenderingFont watermarkFont = new RenderingFont("Arial", 48);
            watermarkFont.Bold = true;
            watermarkFont.Color = Color.Gray; // Light gray for subtle appearance

            // Initialize the text watermark with the desired text and font
            RenderingWatermark watermark = new RenderingWatermark("CONFIDENTIAL", watermarkFont);

            // Configure watermark appearance
            watermark.Rotation = 45f;          // Diagonal orientation
            watermark.Opacity = 0.3f;         // Semi‑transparent (30% opacity)
            watermark.IsBackground = true;   // Place behind page contents
            watermark.ScaleToPagePercent = 150; // Scale relative to page size
            watermark.HAlignment = TextAlignmentType.Center;
            watermark.VAlignment = TextAlignmentType.Center;

            // Set up PDF save options and assign the watermark
            PdfSaveOptions pdfOptions = new PdfSaveOptions();
            pdfOptions.Watermark = watermark;

            // Save the workbook as PDF with the watermark applied to each page
            workbook.Save("Output_With_Watermark.pdf", pdfOptions);
        }
    }
}