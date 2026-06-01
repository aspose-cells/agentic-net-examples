using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsWatermarkDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (empty worksheet)
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sample content for watermark demonstration");

            // Load the image that will be used as the watermark
            string imagePath = "watermark.png";
            if (!File.Exists(imagePath))
            {
                Console.WriteLine($"Image file not found: {imagePath}");
                return;
            }
            byte[] imageData = File.ReadAllBytes(imagePath);

            // Create an image‑based watermark
            RenderingWatermark watermark = new RenderingWatermark(imageData)
            {
                // Scale to fill the entire page (maintains aspect ratio)
                ScaleToPagePercent = 100,

                // Center the watermark horizontally and vertically
                HAlignment = TextAlignmentType.Center,
                VAlignment = TextAlignmentType.Center,

                // Optional visual settings
                Opacity = 0.3f,          // semi‑transparent
                IsBackground = true,    // place behind page contents
                Rotation = 0f           // no rotation
            };

            // Configure PDF save options with the watermark
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                Watermark = watermark
            };

            // Save the workbook as a PDF with the full‑page image watermark
            string outputPath = "WatermarkedOutput.pdf";
            workbook.Save(outputPath, pdfOptions);

            Console.WriteLine($"Workbook saved with watermark to: {outputPath}");
        }
    }
}