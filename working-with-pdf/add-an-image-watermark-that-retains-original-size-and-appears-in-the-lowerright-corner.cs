using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using System.Drawing;

namespace AsposeCellsWatermarkExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (empty workbook with a default worksheet)
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sample data to demonstrate watermark positioning.");

            // Load the image that will be used as the watermark
            string imagePath = "watermark.png";
            if (!File.Exists(imagePath))
            {
                Console.WriteLine($"Image file not found: {imagePath}");
                return;
            }
            byte[] imageData = File.ReadAllBytes(imagePath);

            // Create an image‑based RenderingWatermark
            RenderingWatermark watermark = new RenderingWatermark(imageData)
            {
                // Keep the original size of the image (no scaling)
                ScaleToPagePercent = 100,

                // Position the watermark in the lower‑right corner
                HAlignment = TextAlignmentType.Right,
                VAlignment = TextAlignmentType.Bottom,

                // Optional: make the watermark fully opaque and placed in front of content
                Opacity = 1.0f,
                IsBackground = false,

                // No rotation
                Rotation = 0
            };

            // Configure PDF save options to include the watermark
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                Watermark = watermark
            };

            // Save the workbook as PDF with the configured watermark
            workbook.Save("WatermarkedOutput.pdf", pdfOptions);
        }
    }
}